"""
Forecast vs Actuals Dashboard
=============================
A production-grade Streamlit application that replicates the Excel
"Forecast vs Actuals" pivot dashboard and adds an interactive
Seasonal-IQR outlier-detection tab for sales-history time series.

Author: Built for Anthropic / Claude
Run:    streamlit run app.py
"""

from __future__ import annotations

import hashlib
import hmac
import io
import re
from datetime import datetime
from typing import List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ---------------------------------------------------------------------------
# Page configuration
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Forecast vs Actuals Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)


def inject_dark_theme_css() -> None:
    """Force a dark theme for all UI elements (Change 1).

    A ``.streamlit/config.toml`` already sets the base dark theme; this CSS
    guarantees consistent dark styling for widgets that the config alone
    doesn't fully cover (multiselects, dropdown menus, the file uploader,
    text inputs, tabs) and keeps text high-contrast/legible.
    """
    st.markdown(
        """
        <style>
        /* ---- App background ---- */
        .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
            background-color: #0e1117;
            color: #e6edf3;
        }
        [data-testid="stHeader"] { background: rgba(0,0,0,0); }

        /* ---- Generic text ---- */
        .stApp, .stMarkdown, label, p, span, div,
        h1, h2, h3, h4, h5, h6 { color: #e6edf3; }
        .stCaption, [data-testid="stCaptionContainer"] { color: #9aa4b2 !important; }

        /* ---- Bordered containers / cards ---- */
        [data-testid="stContainer"], div[data-testid="stVerticalBlockBorderWrapper"] {
            background-color: #141925;
            border-color: #2a3040 !important;
        }

        /* ---- Inputs: text, multiselect, selectbox, slider ---- */
        [data-baseweb="input"], [data-baseweb="select"] > div,
        [data-baseweb="base-input"], .stTextInput input, .stNumberInput input {
            background-color: #1a1f2b !important;
            color: #e6edf3 !important;
            border-color: #2a3040 !important;
        }
        /* Multiselect selected "tags" */
        [data-baseweb="tag"] {
            background-color: #2d4b7a !important;
            color: #ffffff !important;
        }
        /* Dropdown popover menu */
        [data-baseweb="popover"], [data-baseweb="menu"], [role="listbox"] {
            background-color: #1a1f2b !important;
            color: #e6edf3 !important;
        }
        [role="option"] { color: #e6edf3 !important; }
        [role="option"]:hover { background-color: #2a3040 !important; }

        /* ---- File uploader ---- */
        [data-testid="stFileUploader"] section,
        [data-testid="stFileUploaderDropzone"] {
            background-color: #1a1f2b !important;
            border: 1px dashed #3a4255 !important;
            color: #e6edf3 !important;
        }
        [data-testid="stFileUploader"] * { color: #e6edf3 !important; }
        [data-testid="stFileUploader"] button {
            background-color: #2d4b7a !important;
            color: #ffffff !important;
            border: none !important;
        }

        /* ---- Buttons / download buttons ---- */
        .stButton button, .stDownloadButton button, .stFormSubmitButton button {
            background-color: #2d4b7a !important;
            color: #ffffff !important;
            border: 1px solid #3a5a8a !important;
        }
        .stButton button:hover, .stDownloadButton button:hover {
            background-color: #3a5a8a !important;
            border-color: #4da3ff !important;
        }

        /* ---- Tabs ---- */
        .stTabs [data-baseweb="tab-list"] { border-bottom-color: #2a3040; }
        .stTabs [data-baseweb="tab"] { color: #9aa4b2; }
        .stTabs [aria-selected="true"] { color: #4da3ff !important; }

        /* ---- Radio / checkbox labels ---- */
        .stRadio label, .stCheckbox label { color: #e6edf3 !important; }

        /* ---- Metric cards ---- */
        [data-testid="stMetric"] {
            background-color: #1a1f2b;
            border: 1px solid #2a3040;
            border-radius: 8px;
            padding: 12px 14px;
        }
        [data-testid="stMetricValue"] { color: #e6edf3 !important; }
        [data-testid="stMetricLabel"] { color: #9aa4b2 !important; }

        /* ---- Dataframe ---- */
        [data-testid="stDataFrame"] { background-color: #1a1f2b; }

        /* ---- Fullscreen chart overlay ----
           When a chart is expanded with the "Fullscreen" button, Streamlit
           renders it inside an overlay container that defaults to white.
           Force that overlay (and its wrapper) to the dark background so the
           chart keeps the same dark theme as the embedded view. */
        [data-testid="stFullScreenFrame"] > div,
        [data-testid="stFullScreenFrame"],
        div[data-testid="stFullScreenFrame"] iframe,
        .element-container:fullscreen,
        [data-testid="stFullScreenFrame"]:fullscreen,
        [data-testid="stFullScreenFrame"] > div:fullscreen {
            background-color: #0e1117 !important;
        }
        /* The vendored fullscreen wrapper used by st.plotly_chart */
        :fullscreen, ::backdrop {
            background-color: #0e1117 !important;
        }
        [data-testid="stFullScreenFrame"] [data-testid="stPlotlyChart"],
        [data-testid="stFullScreenFrame"] .js-plotly-plot,
        [data-testid="stFullScreenFrame"] .plot-container {
            background-color: #0e1117 !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

# ---------------------------------------------------------------------------
# Authentication
# ---------------------------------------------------------------------------
# Credentials are stored as SHA-256 hashes — never as plaintext. Each user's
# password follows the agreed pattern (first four letters of the username,
# first letter capitalised, followed by '@123'). The username lookup is
# case-insensitive; the password check is exact and timing-safe.
AUTHORIZED_USERS = {
    # username (lowercase) : sha256(password)
    "abhishek": "3946259a858b6ba2a76770041490d296ee6c8556a65f89d3f15286cbccd556b8",
    "tonny":    "7345bf2a3c10fd13589c62ac3d33222c3de15829ba59f92720390dacfe2096bb",
    "michael":  "1caed70443e860896d7289678ab817688a1d59ae85af100bcbfb482c65f96b58",
    "sara":     "cabe78b03175eeff37bf7e5f9c6e4691136967d1cd7272f177b952bee06000e9",
    "mercedes": "0e0fb6880ccc17217a82a016fd2bac3e4a40cbed30a5843b51ebeefce635f4b5",
    "susan":    "8fa7c5a2aed87ae68a92256da79704df110b5d3988bc9e810f7437d16edd8a6a",
    "ria":      "513e987f258e7c73b1310ee5f56076dfb1060788c330f2f044ca02c7f87d4d8d",
}

# Display names, keyed by the lowercase username, for the welcome message.
USER_DISPLAY_NAMES = {
    "abhishek": "Abhishek",
    "tonny": "Tonny",
    "michael": "Michael",
    "sara": "Sara",
    "mercedes": "Mercedes",
    "susan": "Susan",
    "ria": "Ria",
}


def verify_credentials(username: str, password: str) -> bool:
    """Return True iff the username/password pair is valid.

    * Username matching is case-insensitive and whitespace-trimmed.
    * Password is hashed with SHA-256 and compared with
      ``hmac.compare_digest`` (constant-time, prevents timing attacks).
    """
    if not username or not password:
        return False
    key = username.strip().lower()
    stored_hash = AUTHORIZED_USERS.get(key)
    if stored_hash is None:
        # Hash anyway so the timing profile is the same for unknown users.
        hashlib.sha256(password.encode("utf-8")).hexdigest()
        return False
    supplied_hash = hashlib.sha256(password.encode("utf-8")).hexdigest()
    return hmac.compare_digest(stored_hash, supplied_hash)


def render_login_screen() -> None:
    """Render the login form. Sets st.session_state['authenticated'] on success."""
    inject_dark_theme_css()
    st.title("🔐 Forecast vs Actuals — Sign in")
    st.caption("Please sign in to access the dashboard.")

    # Center the form in a narrower column for a cleaner look.
    _, mid, _ = st.columns([1, 1.2, 1])
    with mid:
        with st.form("login_form", clear_on_submit=False):
            username = st.text_input("Username", key="login::username",
                                     autocomplete="username")
            password = st.text_input("Password", type="password",
                                     key="login::password",
                                     autocomplete="current-password")
            submitted = st.form_submit_button("Sign in",
                                              use_container_width=True)

        if submitted:
            if verify_credentials(username, password):
                st.session_state["authenticated"] = True
                st.session_state["auth_user"] = username.strip().lower()
                # Don't keep the raw password in session state.
                if "login::password" in st.session_state:
                    del st.session_state["login::password"]
                st.rerun()
            else:
                st.error("❌ Invalid username or password. Please try again.")

    st.stop()


def logout() -> None:
    """Clear authentication and all app state, then rerun to the login page."""
    st.session_state["authenticated"] = False
    st.session_state.pop("auth_user", None)
    st.session_state["current_file_id"] = None
    reset_app_state()
    st.rerun()


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
ID_COLUMNS: List[str] = [
    "Key",
    "Business Line",
    "Material",
    "Material code",
    "Ship To Sub Region",
    "Arkieva ABC",
    "Arkieva Pattern",
    "Stat Flag",
    "Data",
]

# Order matches the Excel slicer order
FILTER_COLUMNS: List[str] = [
    "Business Line",
    "Arkieva ABC",
    "Ship To Sub Region",
    "Material",
    "Stat Flag",
    "Arkieva Pattern",
    "Data",
]

SALES_HISTORY_LABEL = "Sales History (kg)"
HISTORY_FOR_FORECAST_LABEL = "History For Forecast (kg)"
STAT_FORECAST_LABEL = "Statistical Forecast (kg)"

# Some input files label the statistical-forecast line as
# "Statistical Forecast Committed (kg)". Any of these aliases is normalised
# to the canonical STAT_FORECAST_LABEL when the file is loaded.
STAT_FORECAST_ALIASES = [
    "Statistical Forecast Committed (kg)",
    "Statistical Forecast (kg)",
]

# Categorisation of the four "Data" series
HISTORY_SERIES = ["Sales History (kg)", "History For Forecast (kg)"]
FORECAST_SERIES = ["Statistical Forecast (kg)", "Final Demand Plan Lag 1 (kg)"]

# Visual style: every series uses a SOLID line (Change 3). History and
# Forecast are distinguished by colour family and marker, not by dashing.
SERIES_STYLE = {
    "Sales History (kg)":           {"color": "#4da3ff", "dash": "solid", "category": "History"},
    "History For Forecast (kg)":    {"color": "#36cfc9", "dash": "solid", "category": "History"},
    "Statistical Forecast (kg)":    {"color": "#ffa94d", "dash": "solid", "category": "Forecast"},
    "Final Demand Plan Lag 1 (kg)": {"color": "#ff6b6b", "dash": "solid", "category": "Forecast"},
}

# ---------------------------------------------------------------------------
# Dark theme palette (Change 1)
# ---------------------------------------------------------------------------
DARK_BG = "#0e1117"           # page background
DARK_PANEL = "#1a1f2b"        # cards / panels
DARK_GRID = "#2a3040"         # chart gridlines
DARK_TEXT = "#e6edf3"         # primary text (high contrast on dark)
DARK_MUTED = "#9aa4b2"        # secondary/muted text
ACCENT = "#4da3ff"            # primary accent

# History / Forecast background bands (tuned for a dark canvas)
HISTORY_BAND_COLOR = "rgba(77, 163, 255, 0.10)"
FORECAST_BAND_COLOR = "rgba(255, 169, 77, 0.10)"

# Trend-line colours (Change 4)
HISTORY_TREND_COLOR = "#82c91e"
FORECAST_TREND_COLOR = "#f783ac"


def dark_layout(fig: "go.Figure", **overrides) -> "go.Figure":
    """Apply the shared dark-theme layout to a Plotly figure.

    Keeps text light and legible on the dark canvas, adds an interactive
    range slider + range-selector buttons, unified hover and spike lines.
    Any keyword overrides are merged into update_layout.
    """
    base = dict(
        template="plotly_dark",
        # Solid dark backgrounds (not transparent): a transparent paper/plot
        # background lets Streamlit's white fullscreen overlay show through
        # when the chart is expanded. Using the app's dark colour keeps the
        # chart dark in BOTH the embedded and fullscreen views, while still
        # blending seamlessly with the dark page background.
        paper_bgcolor=DARK_BG,
        plot_bgcolor=DARK_BG,
        font=dict(color=DARK_TEXT, size=13),
        hovermode="x unified",
        hoverlabel=dict(bgcolor=DARK_PANEL, font_size=12,
                        font_color=DARK_TEXT, bordercolor=DARK_GRID),
        legend=dict(
            orientation="h", yanchor="top", y=-0.32,
            xanchor="center", x=0.5,
            bgcolor="rgba(0,0,0,0)", font=dict(color=DARK_TEXT),
        ),
        margin=dict(t=70, l=60, r=30, b=120),
    )
    base.update(overrides)
    fig.update_layout(**base)
    fig.update_xaxes(
        showgrid=True, gridcolor=DARK_GRID, zeroline=False,
        showspikes=True, spikemode="across", spikethickness=1,
        spikecolor=DARK_MUTED, color=DARK_TEXT,
    )
    fig.update_yaxes(
        showgrid=True, gridcolor=DARK_GRID, zeroline=False,
        color=DARK_TEXT, tickformat=",.0f",
    )
    return fig


def add_time_controls(fig: "go.Figure") -> "go.Figure":
    """Add an interactive range slider + quick range-selector buttons."""
    fig.update_xaxes(
        rangeslider=dict(visible=True, thickness=0.06,
                         bgcolor=DARK_PANEL),
        rangeselector=dict(
            bgcolor=DARK_PANEL, activecolor=ACCENT,
            font=dict(color=DARK_TEXT),
            buttons=[
                dict(count=6, label="6m", step="month", stepmode="backward"),
                dict(count=1, label="1y", step="year", stepmode="backward"),
                dict(count=2, label="2y", step="year", stepmode="backward"),
                dict(step="all", label="All"),
            ],
        ),
        tickformat="%b %Y",
    )
    return fig


# ---------------------------------------------------------------------------
# Data loading & validation
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_excel(file_bytes: bytes, file_name: str) -> Tuple[pd.DataFrame, List[pd.Timestamp]]:
    """Load and reshape the uploaded workbook.

    Returns a long-format DataFrame ['Key', ..., 'Date', 'Value'] together
    with the sorted list of months found in the file. Raises ValueError
    with a user-friendly message if the file is invalid.
    """
    try:
        raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0)
    except Exception as exc:  # noqa: BLE001
        raise ValueError(f"Could not read the Excel file: {exc}") from exc

    missing = [c for c in ID_COLUMNS if c not in raw.columns]
    if missing:
        raise ValueError(
            "The uploaded file is missing required column(s): " + ", ".join(missing)
        )

    month_cols = [c for c in raw.columns if c not in ID_COLUMNS]
    parsed_months = []
    for c in month_cols:
        if isinstance(c, (pd.Timestamp, datetime)):
            parsed_months.append((c, pd.Timestamp(c)))
        else:
            try:
                parsed_months.append((c, pd.Timestamp(c)))
            except (ValueError, TypeError):
                continue

    if not parsed_months:
        raise ValueError(
            "No month/date columns found. Expected dated columns such as "
            "'2023-04-01', '2023-05-01', ..."
        )

    keep_cols = ID_COLUMNS + [orig for orig, _ in parsed_months]
    df = raw[keep_cols].copy()
    df = df.rename(columns={orig: ts for orig, ts in parsed_months})

    long_df = df.melt(
        id_vars=ID_COLUMNS,
        value_vars=[ts for _, ts in parsed_months],
        var_name="Date",
        value_name="Value",
    )
    long_df["Date"] = pd.to_datetime(long_df["Date"])
    long_df["Value"] = pd.to_numeric(long_df["Value"], errors="coerce")

    # Change 1: normalise the statistical-forecast label. Some files use
    # "Statistical Forecast Committed (kg)"; map every alias to the canonical
    # "Statistical Forecast (kg)" so the rest of the app is label-agnostic.
    long_df["Data"] = long_df["Data"].astype(str).str.strip()
    long_df["Data"] = long_df["Data"].replace(
        "Statistical Forecast Committed (kg)", STAT_FORECAST_LABEL
    )

    for c in FILTER_COLUMNS:
        if long_df[c].dtype == object:
            long_df[c] = long_df[c].astype(str).str.strip()
        else:
            long_df[c] = long_df[c].astype(str)

    sorted_months = sorted({ts for _, ts in parsed_months})
    return long_df, sorted_months


def extract_month_from_filename(file_name: str) -> Optional[str]:
    """Pick the trailing month token from the file stem (e.g. '...April.xlsx')."""
    if not file_name:
        return None
    stem = re.sub(r"\.xlsx?$", "", file_name, flags=re.IGNORECASE)
    last_token = stem.split("_")[-1].split(" ")[-1].strip()
    months = {
        "january", "february", "march", "april", "may", "june",
        "july", "august", "september", "october", "november", "december",
    }
    return last_token if last_token.lower() in months else None


# ---------------------------------------------------------------------------
# Filtering helpers
# ---------------------------------------------------------------------------
def unique_sorted(series: pd.Series) -> List[str]:
    return sorted(series.dropna().astype(str).unique().tolist())


def apply_filters(df: pd.DataFrame, selections: dict) -> pd.DataFrame:
    out = df
    for col, sel in selections.items():
        if sel:  # non-empty list ⇒ active filter
            out = out[out[col].isin(sel)]
    return out


def cascading_options(
    df: pd.DataFrame,
    columns: List[str],
    current_selections: dict,
) -> dict:
    """For each filter column, compute the option universe consistent with
    the current selections in **all other** filter columns.
    """
    options: dict = {}
    for col in columns:
        other_selections = {c: v for c, v in current_selections.items() if c != col}
        narrowed = apply_filters(df, other_selections)
        options[col] = unique_sorted(narrowed[col])
    return options


def reconcile_selection(state_key: str, valid_options: List[str]) -> List[str]:
    """Drop any value from session_state[state_key] that's no longer valid."""
    current = st.session_state.get(state_key, [])
    if not isinstance(current, list):
        current = []
    cleaned = [v for v in current if v in valid_options]
    if cleaned != current:
        st.session_state[state_key] = cleaned
    return cleaned


def render_filter_strip(
    df: pd.DataFrame,
    columns: List[str],
    key_prefix: str,
    n_per_row: int = 4,
) -> dict:
    """Render a cascading multi-select filter strip above the chart.

    Behavior matches Excel slicers:
    * Each filter shows only values that are consistent with the
      selections in the *other* filters.
    * Selecting values in one filter narrows the options in every other.
    * Stale selections are silently pruned so the UI never crashes.

    The pruning is single-pass and proceeds in the declared filter order:
    earlier filters take precedence over later ones when they conflict.

    Each tab passes a unique ``key_prefix`` so filter state is
    independent across tabs and Streamlit doesn't see duplicate keys.
    """
    # Step 1: read current selections, dropping any value that doesn't even
    # exist in the underlying data (defensive against schema changes).
    current_selections: dict = {}
    for col in columns:
        state_key = f"{key_prefix}::{col}"
        current_selections[col] = reconcile_selection(
            state_key, unique_sorted(df[col])
        )

    # Step 2: walk filters in declared order. For each filter, compute its
    # valid options given only the already-pruned EARLIER filters. Drop any
    # of this filter's values that aren't in those options.
    for i, col in enumerate(columns):
        earlier = {c: current_selections[c] for c in columns[:i]}
        narrowed = apply_filters(df, earlier)
        valid_here = unique_sorted(narrowed[col])

        state_key = f"{key_prefix}::{col}"
        before = list(current_selections[col])
        cleaned = [v for v in before if v in valid_here]
        if cleaned != before:
            st.session_state[state_key] = cleaned
            current_selections[col] = cleaned

    # Step 3: compute the final option set for each filter.
    options_per_col = cascading_options(df, columns, current_selections)

    # Step 4: render the widgets.
    rows = [columns[i:i + n_per_row] for i in range(0, len(columns), n_per_row)]
    selections: dict = {}
    for row in rows:
        cols = st.columns(len(row))
        for slot, col_name in zip(cols, row):
            opts = options_per_col[col_name]
            full_count = len(unique_sorted(df[col_name]))
            placeholder = (
                f"All ({len(opts)})"
                if len(opts) == full_count
                else f"{len(opts)} of {full_count} (filtered)"
            )
            with slot:
                selections[col_name] = st.multiselect(
                    col_name,
                    options=opts,
                    key=f"{key_prefix}::{col_name}",
                    placeholder=placeholder,
                )

    spacer, btn_col = st.columns([8, 2])
    with btn_col:
        if st.button("🔄 Reset filters", key=f"{key_prefix}::reset",
                     use_container_width=True):
            for col_name in columns:
                st.session_state[f"{key_prefix}::{col_name}"] = []
            st.rerun()

    return selections


# ---------------------------------------------------------------------------
# Seasonal IQR outlier detection
# ---------------------------------------------------------------------------
def seasonal_iqr_outliers(
    series_df: pd.DataFrame,
    k: float = 1.5,
    min_points_per_month: int = 2,
) -> pd.DataFrame:
    """Detect outliers using the Seasonal IQR method.

    For each calendar month-of-year, historical Q1, Q3 and IQR are computed
    across all years. Points outside [Q1 − k·IQR, Q3 + k·IQR] are flagged.
    Months with fewer than ``min_points_per_month`` historical observations
    fall back to the *global* IQR for that series.
    """
    if series_df.empty:
        return pd.DataFrame(
            columns=["Date", "Value", "Month", "Q1", "Q3", "IQR",
                     "Lower", "Upper", "IsOutlier", "BoundsSource"]
        )

    work = series_df[["Date", "Value"]].dropna(subset=["Date"]).copy()
    work["Date"] = pd.to_datetime(work["Date"])
    work["Month"] = work["Date"].dt.month
    work = work.sort_values("Date").reset_index(drop=True)

    valid_mask = work["Value"].notna()
    valid_vals = work.loc[valid_mask, "Value"]
    if len(valid_vals) >= 4:
        gq1, gq3 = np.percentile(valid_vals, [25, 75])
    elif len(valid_vals) >= 1:
        gq1 = float(valid_vals.min())
        gq3 = float(valid_vals.max())
    else:
        gq1 = gq3 = np.nan
    g_iqr = gq3 - gq1 if pd.notna(gq1) else np.nan

    seasonal_stats = {}
    for m, grp in work[valid_mask].groupby("Month"):
        vals = grp["Value"].values
        if len(vals) >= min_points_per_month:
            q1, q3 = np.percentile(vals, [25, 75])
            seasonal_stats[m] = (q1, q3, q3 - q1, "seasonal")
        else:
            if pd.notna(g_iqr):
                seasonal_stats[m] = (gq1, gq3, g_iqr, "global")
            else:
                seasonal_stats[m] = (np.nan, np.nan, np.nan, "insufficient")

    rows = []
    for _, row in work.iterrows():
        m = row["Month"]
        q1, q3, iqr, src = seasonal_stats.get(
            m, (gq1, gq3, g_iqr, "global" if pd.notna(g_iqr) else "insufficient")
        )
        if pd.isna(iqr) or pd.isna(row["Value"]):
            lower = upper = np.nan
            is_out = False
        else:
            lower = q1 - k * iqr
            upper = q3 + k * iqr
            is_out = bool(row["Value"] < lower or row["Value"] > upper)
        rows.append({
            "Date": row["Date"], "Value": row["Value"], "Month": int(m),
            "Q1": q1, "Q3": q3, "IQR": iqr,
            "Lower": lower, "Upper": upper,
            "IsOutlier": is_out, "BoundsSource": src,
        })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Outlier detection & correction — IQR and Sigma methods (Change 2)
# ---------------------------------------------------------------------------
def detect_and_correct_outliers(
    series_df: pd.DataFrame,
    method: str = "IQR",
    k: float = 1.5,
    n_sigma: float = 3.0,
) -> pd.DataFrame:
    """Detect and correct outliers in a single time series.

    Two statistically valid methods are supported:

    * **IQR (Tukey's fences)** — compute Q1, Q3 and IQR = Q3 − Q1 over the
      series. Points outside ``[Q1 − k·IQR, Q3 + k·IQR]`` are outliers. Each
      outlier is corrected by replacing it with the series **median** (a
      robust central estimate that is unaffected by the outliers it
      replaces).
    * **Sigma (z-score / empirical rule)** — compute the mean μ and standard
      deviation σ. Points with ``|x − μ| > n_sigma·σ`` are outliers. Because
      the ordinary mean/σ are themselves inflated by outliers, the bounds are
      computed on a **robust, outlier-trimmed** basis: μ and σ are estimated
      after excluding points beyond the median ± n_sigma·(1.4826·MAD). Each
      outlier is corrected by replacing it with the (robust) mean.

    Returns a DataFrame with a consistent schema:
    Date, Value, Center, Lower, Upper, IsOutlier, Filtered.
    ``Filtered`` is the cleansed series (outliers replaced, others unchanged).
    """
    cols = ["Date", "Value", "Center", "Lower", "Upper", "IsOutlier", "Filtered"]
    if series_df.empty:
        return pd.DataFrame(columns=cols)

    work = series_df[["Date", "Value"]].dropna(subset=["Date"]).copy()
    work["Date"] = pd.to_datetime(work["Date"])
    work = work.sort_values("Date").reset_index(drop=True)

    vals = work["Value"].to_numpy(dtype="float64")
    valid = vals[~np.isnan(vals)]

    method_u = (method or "IQR").upper()

    if valid.size == 0:
        work["Center"] = np.nan
        work["Lower"] = np.nan
        work["Upper"] = np.nan
        work["IsOutlier"] = False
        work["Filtered"] = vals
        return work[cols]

    if method_u == "IQR":
        q1, q3 = np.percentile(valid, [25, 75])
        iqr = q3 - q1
        lower = q1 - k * iqr
        upper = q3 + k * iqr
        center = float(np.median(valid))   # correction target (robust)
    elif method_u == "SIGMA":
        # Robust pre-pass to keep the mean/σ from being inflated by outliers.
        med = float(np.median(valid))
        mad = float(np.median(np.abs(valid - med)))
        robust_sigma = 1.4826 * mad
        if robust_sigma > 0:
            keep = valid[np.abs(valid - med) <= n_sigma * robust_sigma]
            if keep.size < 2:
                keep = valid
        else:
            keep = valid
        mu = float(np.mean(keep))
        sigma = float(np.std(keep, ddof=1)) if keep.size > 1 else 0.0
        lower = mu - n_sigma * sigma
        upper = mu + n_sigma * sigma
        center = mu                          # correction target
    else:
        raise ValueError(f"Unknown method: {method!r} (expected 'IQR' or 'Sigma')")

    is_out = np.zeros(len(vals), dtype=bool)
    filtered = vals.copy()
    for i, v in enumerate(vals):
        if np.isnan(v):
            continue
        if v < lower or v > upper:
            is_out[i] = True
            filtered[i] = center            # statistically valid correction

    work["Center"] = center
    work["Lower"] = lower
    work["Upper"] = upper
    work["IsOutlier"] = is_out
    work["Filtered"] = filtered
    return work[cols]


# ---------------------------------------------------------------------------
# Hampel filter — outlier detection AND correction
# ---------------------------------------------------------------------------
# Scale factor that makes the MAD a consistent estimator of the standard
# deviation for normally-distributed data (1 / Phi^-1(0.75) ≈ 1.4826).
_MAD_TO_SIGMA = 1.4826


def hampel_filter(
    series_df: pd.DataFrame,
    window_size: int = 5,
    n_sigma: float = 3.0,
) -> pd.DataFrame:
    """Apply the Hampel filter to a single time series for outlier detection
    and correction.

    Implements the algorithm described by Otero Pedrido
    (https://medium.com/@migueloteropedrido/hampel-filter-with-python-17db1d265375)
    and the ``hampel`` PyPI library:

    * A centred rolling window of half-width ``window_size`` slides over the
      series (so the full window spans ``2*window_size + 1`` points).
    * For each point the window **median** and the **MAD** (median absolute
      deviation) are computed.
    * The detection threshold at each point is ``n_sigma * 1.4826 * MAD``.
    * A point is an **outlier** when ``|x_i - median_i| > threshold_i``.
    * Each outlier is **corrected** by replacing it with the window median,
      producing the cleansed series ``Filtered``.

    Parameters
    ----------
    series_df : DataFrame with 'Date' and 'Value' columns (single Key, single
                Data label — typically Sales History).
    window_size : half-width of the rolling window (default 5, matching the
                library's default full window behaviour at the edges).
    n_sigma : tolerance in robust standard deviations (default 3.0).

    Returns
    -------
    DataFrame with columns: Date, Value, Median, MAD, Threshold, Lower, Upper,
    IsOutlier, Filtered.  ``Filtered`` is the Hampel-cleansed history.
    """
    cols = ["Date", "Value", "Median", "MAD", "Threshold",
            "Lower", "Upper", "IsOutlier", "Filtered"]
    if series_df.empty:
        return pd.DataFrame(columns=cols)

    work = series_df[["Date", "Value"]].dropna(subset=["Date"]).copy()
    work["Date"] = pd.to_datetime(work["Date"])
    work = work.sort_values("Date").reset_index(drop=True)

    n = len(work)
    values = work["Value"].to_numpy(dtype="float64")
    filtered = values.copy()

    medians = np.full(n, np.nan)
    mads = np.full(n, np.nan)
    thresholds = np.full(n, np.nan)
    is_outlier = np.zeros(n, dtype=bool)

    half = max(int(window_size), 1)
    for i in range(n):
        lo = max(0, i - half)
        hi = min(n, i + half + 1)
        window = values[lo:hi]
        window = window[~np.isnan(window)]
        if window.size == 0 or np.isnan(values[i]):
            # Cannot evaluate — leave as-is, not an outlier.
            continue
        med = np.median(window)
        mad = np.median(np.abs(window - med))
        sigma = _MAD_TO_SIGMA * mad
        thr = n_sigma * sigma

        medians[i] = med
        mads[i] = mad
        thresholds[i] = thr

        # When MAD == 0 the window is (near-)constant; only a strictly
        # different value can be an outlier.
        if abs(values[i] - med) > thr:
            is_outlier[i] = True
            filtered[i] = med  # correction: replace with window median

    out = pd.DataFrame({
        "Date": work["Date"].to_numpy(),
        "Value": values,
        "Median": medians,
        "MAD": mads,
        "Threshold": thresholds,
        "Lower": medians - thresholds,
        "Upper": medians + thresholds,
        "IsOutlier": is_outlier,
        "Filtered": filtered,
    })
    return out


# ---------------------------------------------------------------------------
# History / Forecast boundary
# ---------------------------------------------------------------------------
def history_forecast_boundary(filtered: pd.DataFrame) -> Optional[pd.Timestamp]:
    """Last date with non-null, non-zero Sales History — the History/Forecast cut-off."""
    sh = filtered[(filtered["Data"] == SALES_HISTORY_LABEL)
                  & filtered["Value"].notna()
                  & (filtered["Value"] != 0)]
    if sh.empty:
        sh = filtered[(filtered["Data"] == SALES_HISTORY_LABEL)
                      & filtered["Value"].notna()]
    if sh.empty:
        return None
    return sh["Date"].max()


# ---------------------------------------------------------------------------
# Trend analysis (Change 4)
# ---------------------------------------------------------------------------
def fit_trend(dates: pd.Series, values: pd.Series) -> Optional[dict]:
    """Fit a simple linear (OLS) trend to a time series.

    Returns a dict with the fitted endpoint line, the slope expressed *per
    month* and *per year*, the % change per year relative to the mean level,
    and the R² goodness-of-fit. Returns None when there are too few points.
    """
    d = pd.DataFrame({"Date": pd.to_datetime(dates), "Value": values}).dropna()
    d = d.sort_values("Date")
    if len(d) < 2:
        return None

    # Convert dates to a month index so the slope is "per month".
    t0 = d["Date"].min()
    x = ((d["Date"].dt.year - t0.year) * 12
         + (d["Date"].dt.month - t0.month)).to_numpy(dtype="float64")
    y = d["Value"].to_numpy(dtype="float64")
    if np.allclose(x, x[0]):
        return None

    slope, intercept = np.polyfit(x, y, 1)
    y_hat = slope * x + intercept
    ss_res = float(np.sum((y - y_hat) ** 2))
    ss_tot = float(np.sum((y - y.mean()) ** 2))
    r2 = 1 - ss_res / ss_tot if ss_tot > 0 else 0.0

    mean_level = float(np.mean(y)) if np.mean(y) != 0 else np.nan
    slope_per_year = slope * 12
    pct_per_year = (slope_per_year / mean_level * 100) if mean_level and not np.isnan(mean_level) else np.nan

    return {
        "slope_per_month": float(slope),
        "slope_per_year": float(slope_per_year),
        "pct_per_year": float(pct_per_year) if not np.isnan(pct_per_year) else None,
        "r2": float(r2),
        "x": x, "dates": d["Date"].to_numpy(),
        "fit_line": y_hat,
        "mean_level": mean_level,
    }


def interpret_trend(hist_trend: Optional[dict],
                    fcst_trend: Optional[dict]) -> Tuple[str, str]:
    """Produce a plain-language interpretation comparing history vs forecast
    trend. Returns (status, message) where status ∈ {ok, warn, info}."""
    if hist_trend is None and fcst_trend is None:
        return "info", "Not enough data to estimate a trend."
    if hist_trend is None:
        return "info", "Not enough Sales-History points to estimate a historical trend."
    if fcst_trend is None:
        return "info", "Not enough forecast points to estimate a forecast trend."

    h = hist_trend["pct_per_year"]
    f = fcst_trend["pct_per_year"]

    def direction(p):
        if p is None:
            return "flat"
        if p > 2:
            return "rising"
        if p < -2:
            return "falling"
        return "broadly flat"

    hd, fd = direction(h), direction(f)
    h_txt = f"{h:+.1f}%/yr" if h is not None else "n/a"
    f_txt = f"{f:+.1f}%/yr" if f is not None else "n/a"

    msg = (f"History is **{hd}** ({h_txt}); forecast is **{fd}** ({f_txt}). ")

    if hd == fd:
        status = "ok"
        msg += ("The forecast follows the historical trend direction, which "
                "suggests the trend assumption is consistent with the past.")
    elif "flat" in hd and "flat" in fd:
        status = "ok"
        msg += "Both are broadly flat — the trend assumption looks consistent."
    else:
        status = "warn"
        msg += ("The forecast trend direction **differs** from history. "
                "Review whether this inflection is intended (e.g. a known "
                "market change) or an artefact of the statistical model.")

    # Flag implausibly steep forecast growth relative to history.
    if h is not None and f is not None and abs(h) > 1e-6:
        ratio = f / h if h != 0 else np.inf
        if ratio > 3 and f > 0 and h > 0:
            status = "warn"
            msg += (" Note: the forecast grows much faster than history "
                    f"(~{ratio:.1f}× the historical rate).")

    return status, msg


# ---------------------------------------------------------------------------
# Seasonality analysis (Change 4)
# ---------------------------------------------------------------------------
def monthly_seasonality_profile(dates, values) -> Optional[pd.Series]:
    """Return a length-12 seasonal index (mean-normalised) indexed by month
    number 1..12, or None if insufficient data. A value of 1.10 means that
    month runs ~10% above the series average."""
    d = pd.DataFrame({"Date": pd.to_datetime(dates), "Value": values}).dropna()
    if d.empty:
        return None
    d["Month"] = d["Date"].dt.month
    monthly = d.groupby("Month")["Value"].mean()
    overall = d["Value"].mean()
    if overall is None or overall == 0 or np.isnan(overall):
        return None
    profile = monthly / overall
    return profile.reindex(range(1, 13))


def seasonality_analysis(sales_hist: pd.DataFrame,
                         forecast: pd.DataFrame,
                         boundary: Optional[pd.Timestamp],
                         n_years: int = 3) -> dict:
    """Compare the year-over-year seasonal shape of the last ``n_years`` of
    Sales History against the forecast period.

    Returns a dict with per-month profiles and a correlation that measures
    how closely the forecast's seasonal shape matches recent history.
    """
    out = {
        "hist_profile": None,
        "fcst_profile": None,
        "per_year": {},      # year -> length-12 profile (history)
        "correlation": None,
        "n_hist_years": 0,
    }

    if sales_hist.empty:
        return out

    sh = sales_hist.dropna(subset=["Value"]).copy()
    sh["Date"] = pd.to_datetime(sh["Date"])
    if boundary is not None:
        sh = sh[sh["Date"] <= boundary]
    sh = sh[sh["Value"] != 0]
    if sh.empty:
        return out

    # Restrict to the last n_years complete-ish years of history.
    last_year = sh["Date"].dt.year.max()
    recent_years = sorted(
        [y for y in sh["Date"].dt.year.unique() if y > last_year - n_years]
    )
    out["n_hist_years"] = len(recent_years)
    sh_recent = sh[sh["Date"].dt.year.isin(recent_years)]

    out["hist_profile"] = monthly_seasonality_profile(
        sh_recent["Date"], sh_recent["Value"])
    for y in recent_years:
        yr = sh_recent[sh_recent["Date"].dt.year == y]
        out["per_year"][int(y)] = monthly_seasonality_profile(yr["Date"], yr["Value"])

    if forecast is not None and not forecast.empty:
        fc = forecast.dropna(subset=["Value"]).copy()
        fc["Date"] = pd.to_datetime(fc["Date"])
        if boundary is not None:
            fc = fc[fc["Date"] > boundary]
        fc = fc[fc["Value"] != 0]
        out["fcst_profile"] = monthly_seasonality_profile(fc["Date"], fc["Value"])

    # Correlate the two seasonal shapes where both are defined.
    hp, fp = out["hist_profile"], out["fcst_profile"]
    if hp is not None and fp is not None:
        both = pd.concat([hp.rename("h"), fp.rename("f")], axis=1).dropna()
        if len(both) >= 3 and both["h"].std() > 0 and both["f"].std() > 0:
            out["correlation"] = float(both["h"].corr(both["f"]))

    return out


def interpret_seasonality(analysis: dict) -> Tuple[str, str]:
    """Plain-language interpretation of the seasonality comparison.
    Returns (status, message)."""
    hp = analysis.get("hist_profile")
    fp = analysis.get("fcst_profile")
    corr = analysis.get("correlation")

    if hp is None:
        return "info", "Not enough historical data to estimate seasonality."
    if fp is None:
        return "info", ("No forecast values available to compare seasonality "
                        "against history.")

    # Strength of seasonality = spread of the index around 1.0.
    hist_amp = float(np.nanstd(hp.values))
    fcst_amp = float(np.nanstd(fp.values))

    peak_month = int(hp.idxmax())
    trough_month = int(hp.idxmin())
    mname = lambda m: pd.Timestamp(2000, m, 1).strftime("%B")

    # Special case: forecast is essentially flat (no seasonal shape) so a
    # correlation can't be computed. If history IS seasonal, that's a problem.
    if corr is None:
        if hist_amp > 0.05 and fcst_amp <= 0.02:
            return "warn", (
                "The forecast is essentially **flat** (no seasonal shape), "
                f"while history is seasonal — peaking around **{mname(peak_month)}** "
                f"and troughing around **{mname(trough_month)}**. The forecast "
                "does not reproduce the historical seasonality; review whether "
                "it should.")
        return "info", ("Couldn't compute a seasonality correlation (the "
                        "seasonal shape may be flat or too sparse).")

    if corr >= 0.7:
        status = "ok"
        verdict = ("strongly matches recent history — the forecast preserves "
                   "the historical seasonal pattern")
    elif corr >= 0.3:
        status = "warn"
        verdict = ("only partially matches recent history — some seasonal "
                   "structure is preserved but there are notable differences")
    else:
        status = "warn"
        verdict = ("does **not** match recent history — the forecast's "
                   "seasonal shape diverges from the last few years")

    msg = (f"Seasonality correlation **{corr:+.2f}**: the forecast {verdict}. "
           f"Historically, demand peaks around **{mname(peak_month)}** and "
           f"troughs around **{mname(trough_month)}**.")
    if status == "warn":
        msg += (" If the business is genuinely seasonal, review whether the "
                "forecast should better reflect this shape.")
    return status, msg


# ---------------------------------------------------------------------------
# Reset on file removal
# ---------------------------------------------------------------------------
def reset_app_state() -> None:
    """Clear filter state, date range, and the data cache."""
    keys_to_clear = [k for k in list(st.session_state.keys())
                     if "::" in k or k.startswith("flt_")]
    for k in keys_to_clear:
        del st.session_state[k]
    load_excel.clear()


# ---------------------------------------------------------------------------
# UI – Tab 1: Forecast vs Actuals
# ---------------------------------------------------------------------------
def render_dashboard_tab(long_df: pd.DataFrame, file_name: str) -> None:
    st.subheader("📈 Forecast v/s Actuals")
    detected_month = extract_month_from_filename(file_name)
    meta_bits = [f"**File:** `{file_name}`"]
    if detected_month:
        meta_bits.append(f"**Reporting month:** {detected_month}")
    st.caption(" • ".join(meta_bits))

    # ---- Filter strip on top -----------------------------------------------
    with st.container(border=True):
        st.markdown("**🔎 Filters** — leave empty to include everything. Filters cascade: each one narrows the choices in the others (Excel-slicer style).")
        selections = render_filter_strip(long_df, FILTER_COLUMNS, key_prefix="dash")

        min_d = long_df["Date"].min().to_pydatetime()
        max_d = long_df["Date"].max().to_pydatetime()
        st.slider(
            "📅 Date range",
            min_value=min_d, max_value=max_d, value=(min_d, max_d),
            format="MMM YYYY", key="dash::date_range",
        )

    d_start, d_end = st.session_state.get("dash::date_range", (min_d, max_d))
    d_start, d_end = pd.Timestamp(d_start), pd.Timestamp(d_end)

    filtered = apply_filters(long_df, selections)
    filtered = filtered[(filtered["Date"] >= d_start) & (filtered["Date"] <= d_end)]

    st.caption(f"Rows after filter: **{len(filtered):,}**")

    if filtered.empty:
        st.warning("No data matches the current filter combination.")
        return

    # ---- KPI cards ----------------------------------------------------------
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Unique Keys", f"{filtered['Key'].nunique():,}")
    c2.metric("Materials", f"{filtered['Material'].nunique():,}")
    c3.metric("Months", f"{filtered['Date'].nunique():,}")
    c4.metric("Total volume (kg)", f"{filtered['Value'].sum(skipna=True):,.0f}")

    # ---- History / Forecast banner -----------------------------------------
    boundary = history_forecast_boundary(filtered)
    bcol1, bcol2 = st.columns(2)
    bcol1.markdown(
        "<div style='background:rgba(77,163,255,0.12); padding:10px 14px;"
        " border-left:4px solid #4da3ff; border-radius:4px;'>"
        "<b style='color:#e6edf3;'>📘 History</b><br>"
        "<span style='font-size:0.85em; color:#9aa4b2;'>"
        "Solid lines · Sales History, History For Forecast"
        "</span></div>",
        unsafe_allow_html=True,
    )
    bcol2.markdown(
        "<div style='background:rgba(255,169,77,0.12); padding:10px 14px;"
        " border-left:4px solid #ffa94d; border-radius:4px;'>"
        "<b style='color:#e6edf3;'>📕 Forecast</b><br>"
        "<span style='font-size:0.85em; color:#9aa4b2;'>"
        "Solid lines · Statistical Forecast, Final Demand Plan Lag 1"
        "</span></div>",
        unsafe_allow_html=True,
    )
    if boundary is not None:
        st.caption(
            f"History / Forecast boundary detected at "
            f"**{boundary.strftime('%b %Y')}** "
            "(last month with Sales History data)."
        )

    # ---- Trend toggle -------------------------------------------------------
    show_trend = st.checkbox(
        "📈 Show trend lines (history & forecast)", value=True,
        key="dash::show_trend",
        help="Fits a linear trend separately to the history and forecast "
             "periods so you can see whether the forecast follows the "
             "historical trajectory.",
    )

    # ---- Pivot & line chart -------------------------------------------------
    pivot = (
        filtered.dropna(subset=["Value"])
        .groupby(["Date", "Data"], as_index=False)["Value"].sum()
        .pivot(index="Date", columns="Data", values="Value")
        .sort_index()
    )

    fig = go.Figure()

    # Background bands for History / Forecast regions
    if boundary is not None and len(pivot.index) > 0:
        x_min = pivot.index.min()
        x_max = pivot.index.max()
        if x_min <= boundary:
            fig.add_vrect(
                x0=x_min, x1=boundary,
                fillcolor=HISTORY_BAND_COLOR, line_width=0, layer="below",
                annotation_text="History", annotation_position="top left",
                annotation_font=dict(color="#4da3ff", size=12),
            )
        if x_max >= boundary:
            fig.add_vrect(
                x0=boundary, x1=x_max,
                fillcolor=FORECAST_BAND_COLOR, line_width=0, layer="below",
                annotation_text="Forecast", annotation_position="top right",
                annotation_font=dict(color="#ffa94d", size=12),
            )
        fig.add_vline(x=boundary, line=dict(color=DARK_MUTED, width=1, dash="dot"))

    # One trace per Data series — all SOLID (Change 3), distinguished by colour
    ordered = [s for s in HISTORY_SERIES + FORECAST_SERIES if s in pivot.columns]
    for col in ordered:
        style = SERIES_STYLE.get(
            col, {"color": "#888", "dash": "solid", "category": "Other"}
        )
        fig.add_trace(
            go.Scatter(
                x=pivot.index, y=pivot[col],
                name=col, mode="lines+markers",
                line=dict(color=style["color"], width=2.6, dash="solid",
                          shape="spline", smoothing=0.5),
                marker=dict(size=6,
                            symbol="circle" if style["category"] == "History" else "diamond"),
                legendgroup=style["category"],
                legendgrouptitle_text=style["category"],
                hovertemplate=(
                    f"<b>{col}</b><br>"
                    f"<i>{style['category']}</i><br>"
                    "%{x|%b %Y}<br>Volume: %{y:,.0f} kg<extra></extra>"
                ),
            )
        )

    # ---- Trend lines (Change 4) --------------------------------------------
    hist_trend = fcst_trend = None
    if boundary is not None:
        sales_series = (filtered[filtered["Data"] == SALES_HISTORY_LABEL]
                        .dropna(subset=["Value"]))
        sales_series = sales_series[sales_series["Date"] <= boundary]
        sales_grp = sales_series.groupby("Date")["Value"].sum()
        hist_trend = fit_trend(sales_grp.index, sales_grp.values)

        fcst_series = (filtered[filtered["Data"] == STAT_FORECAST_LABEL]
                       .dropna(subset=["Value"]))
        fcst_series = fcst_series[fcst_series["Date"] > boundary]
        fcst_grp = fcst_series.groupby("Date")["Value"].sum()
        fcst_trend = fit_trend(fcst_grp.index, fcst_grp.values)

    if show_trend:
        if hist_trend is not None:
            fig.add_trace(go.Scatter(
                x=hist_trend["dates"], y=hist_trend["fit_line"],
                mode="lines", name="History trend",
                line=dict(color=HISTORY_TREND_COLOR, width=2.4, dash="dash"),
                legendgroup="Trend", legendgrouptitle_text="Trend",
                hovertemplate="History trend<br>%{x|%b %Y}<br>%{y:,.0f} kg<extra></extra>",
            ))
        if fcst_trend is not None:
            fig.add_trace(go.Scatter(
                x=fcst_trend["dates"], y=fcst_trend["fit_line"],
                mode="lines", name="Forecast trend",
                line=dict(color=FORECAST_TREND_COLOR, width=2.4, dash="dash"),
                legendgroup="Trend", legendgrouptitle_text="Trend",
                hovertemplate="Forecast trend<br>%{x|%b %Y}<br>%{y:,.0f} kg<extra></extra>",
            ))

    dark_layout(
        fig,
        title="Forecast v/s Actuals",
        xaxis_title="Months", yaxis_title="Demand volume in kgs",
        height=560,
        legend=dict(
            orientation="h", yanchor="top", y=-0.36,
            xanchor="center", x=0.5, groupclick="toggleitem",
            bgcolor="rgba(0,0,0,0)", font=dict(color=DARK_TEXT),
        ),
        margin=dict(t=70, l=60, r=30, b=150),
    )
    add_time_controls(fig)

    st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False})

    # ---- Trend interpretation ----------------------------------------------
    if show_trend:
        status, message = interpret_trend(hist_trend, fcst_trend)
        st.markdown("#### 🧭 Trend interpretation")
        if status == "ok":
            st.success(message)
        elif status == "warn":
            st.warning(message)
        else:
            st.info(message)


# ---------------------------------------------------------------------------
# UI – Tab 2: Outlier Detection & Correction (Hampel filter)
# ---------------------------------------------------------------------------
def render_outlier_tab(long_df: pd.DataFrame) -> None:
    st.subheader("🚨 Outlier Detection & Correction")
    st.caption(
        "Detect and **correct** outliers on **Sales History (kg)**. Choose a "
        "method below. Both the original history and the cleansed history "
        "(after correction) are shown. *History For Forecast (kg)* is shown "
        "unchanged for comparison."
    )

    sales = long_df[long_df["Data"] == SALES_HISTORY_LABEL].copy()
    if sales.empty:
        st.warning(f"No rows with Data == '{SALES_HISTORY_LABEL}' were found.")
        return

    hist_ff = long_df[long_df["Data"] == HISTORY_FOR_FORECAST_LABEL].copy()

    # 'Data' filter is omitted: this view is intrinsically Sales History.
    outlier_filter_cols = [c for c in FILTER_COLUMNS if c != "Data"]

    with st.container(border=True):
        st.markdown("**🔎 Filters** — leave empty to include everything. Filters cascade: each one narrows the choices in the others (Excel-slicer style).")
        selections = render_filter_strip(
            sales, outlier_filter_cols, key_prefix="outlier",
        )

    sales_f = apply_filters(sales, selections)
    if sales_f.empty:
        st.warning("No Sales History rows match the current filter combination.")
        return

    st.caption(
        f"Rows after filter: **{len(sales_f):,}** "
        f"covering **{sales_f['Key'].nunique():,}** Keys"
    )

    # ---- Method selection (Change 2) ---------------------------------------
    mcol, pcol = st.columns([1, 1.3])
    with mcol:
        method = st.radio(
            "Detection & correction method",
            options=["IQR", "Sigma"],
            horizontal=True,
            key="outlier::method",
            help="IQR uses Tukey's fences (Q1−k·IQR / Q3+k·IQR) and corrects "
                 "to the median. Sigma uses mean ± n·σ (robustly estimated) "
                 "and corrects to the mean.",
        )
    with pcol:
        if method == "IQR":
            param = st.slider(
                "IQR multiplier (k)", min_value=0.5, max_value=3.0,
                value=1.5, step=0.1,
                help="Tukey's rule uses 1.5; larger values flag fewer points.",
                key="outlier::k",
            )
            k_val, sigma_val = param, 3.0
        else:
            param = st.slider(
                "Sigma threshold (n)", min_value=1.0, max_value=4.0,
                value=3.0, step=0.5,
                help="Points beyond mean ± n·σ are outliers. 3σ is typical.",
                key="outlier::nsigma",
            )
            k_val, sigma_val = 1.5, param

    @st.cache_data(show_spinner="Detecting & correcting outliers across all Keys…")
    def compute_all(sales_in: pd.DataFrame, method_in: str,
                    k_in: float, sigma_in: float) -> pd.DataFrame:
        frames = []
        for key, grp in sales_in.groupby("Key"):
            res = detect_and_correct_outliers(
                grp, method=method_in, k=k_in, n_sigma=sigma_in)
            if res.empty:
                continue
            res.insert(0, "Key", key)
            frames.append(res)
        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

    all_results = compute_all(sales_f, method, k_val, sigma_val)
    if all_results.empty:
        st.info("No data available for outlier computation.")
        return

    total_pts = len(all_results)
    total_outliers = int(all_results["IsOutlier"].sum())
    keys_with_outliers = all_results.loc[all_results["IsOutlier"], "Key"].nunique()
    total_keys = all_results["Key"].nunique()

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Keys analysed", f"{total_keys:,}")
    m2.metric("Keys with ≥1 outlier", f"{keys_with_outliers:,}")
    m3.metric("Total observations", f"{total_pts:,}")
    rate = (total_outliers / total_pts * 100) if total_pts else 0
    m4.metric("Outliers corrected", f"{total_outliers:,}", delta=f"{rate:.1f}% of points")

    # ---- Per-Key chooser ----------------------------------------------------
    st.markdown("### 🔬 Inspect a single Key")
    keys_sorted = sorted(all_results["Key"].unique().tolist())
    keys_with_out_set = set(all_results.loc[all_results["IsOutlier"], "Key"].unique())
    only_with_out = st.checkbox(
        "Show only Keys with at least one outlier",
        value=bool(keys_with_outliers),
        key="outlier::only_with",
    )
    key_options = (
        [k for k in keys_sorted if k in keys_with_out_set]
        if only_with_out else keys_sorted
    )

    if not key_options:
        st.info("No Keys match the current selection.")
    else:
        chosen_key = st.selectbox("Select a Key", key_options, index=0,
                                  key="outlier::chosen_key")
        key_df = all_results[all_results["Key"] == chosen_key].sort_values("Date")
        hff_key = hist_ff[hist_ff["Key"] == chosen_key].sort_values("Date") \
            if not hist_ff.empty else pd.DataFrame(columns=["Date", "Value"])
        _render_key_chart(chosen_key, key_df, hff_key, method,
                          k_val if method == "IQR" else sigma_val)

        # Per-Key cleansed-history download
        cleansed = key_df[["Date", "Value", "Filtered", "IsOutlier"]].copy()
        cleansed = cleansed.rename(columns={
            "Value": "Sales History (kg)",
            "Filtered": "Cleansed History (kg)",
        })
        cleansed["Date"] = pd.to_datetime(cleansed["Date"]).dt.strftime("%Y-%m-%d")
        cbuf = io.StringIO()
        cleansed.to_csv(cbuf, index=False)
        st.download_button(
            "⬇️ Download this Key's cleansed history (CSV)",
            data=cbuf.getvalue(),
            file_name=f"cleansed_{method.lower()}_{chosen_key}.csv",
            mime="text/csv",
            key="outlier::dl_key",
        )

    # ---- Full outlier listing -----------------------------------------------
    st.markdown("### 📋 All flagged & corrected outliers")
    outlier_df = all_results[all_results["IsOutlier"]].copy()
    if outlier_df.empty:
        st.success("No outliers detected with the current settings. 🎉")
    else:
        meta = (sales_f.drop_duplicates(subset=["Key"])
                [["Key", "Business Line", "Material", "Ship To Sub Region",
                  "Arkieva ABC", "Arkieva Pattern"]])
        outlier_df = outlier_df.merge(meta, on="Key", how="left")
        outlier_df["Date"] = pd.to_datetime(outlier_df["Date"]).dt.strftime("%Y-%m-%d")
        outlier_df = outlier_df.rename(columns={
            "Value": "Original", "Filtered": "Corrected"})
        cols = ["Key", "Business Line", "Material", "Ship To Sub Region",
                "Arkieva ABC", "Arkieva Pattern", "Date", "Original",
                "Corrected", "Center", "Lower", "Upper"]
        st.dataframe(
            outlier_df[cols].style.format(
                {c: "{:,.2f}" for c in
                 ["Original", "Corrected", "Center", "Lower", "Upper"]},
                na_rep="–",
            ),
            use_container_width=True,
        )
        buf = io.StringIO()
        outlier_df[cols].to_csv(buf, index=False)
        st.download_button(
            "⬇️ Download all corrected outliers (CSV)",
            data=buf.getvalue(),
            file_name=f"corrected_outliers_{method.lower()}.csv",
            mime="text/csv",
            key="outlier::dl_all",
        )

    # ---- Full cleansed-history export (all Keys) ----------------------------
    st.markdown("### 🧼 Cleansed history export")
    st.caption(
        "The full cleansed Sales History for every Key in the current filter "
        "selection — outliers replaced by the method's central estimate, all "
        "other points unchanged."
    )
    full_clean = all_results[["Key", "Date", "Value", "Filtered", "IsOutlier"]].copy()
    full_clean = full_clean.rename(columns={
        "Value": "Sales History (kg)",
        "Filtered": "Cleansed History (kg)",
    })
    full_clean["Date"] = pd.to_datetime(full_clean["Date"]).dt.strftime("%Y-%m-%d")
    fbuf = io.StringIO()
    full_clean.to_csv(fbuf, index=False)
    st.download_button(
        "⬇️ Download full cleansed history — all Keys (CSV)",
        data=fbuf.getvalue(),
        file_name=f"cleansed_history_all_keys_{method.lower()}.csv",
        mime="text/csv",
        key="outlier::dl_full_clean",
    )


def _render_key_chart(
    key: str,
    key_df: pd.DataFrame,
    hff_df: pd.DataFrame,
    method: str,
    param: float,
) -> None:
    fig = go.Figure()

    # Threshold band (center ± bound)
    bounds_df = key_df.dropna(subset=["Lower", "Upper"]).sort_values("Date")
    if not bounds_df.empty:
        fig.add_trace(go.Scatter(
            x=bounds_df["Date"], y=bounds_df["Upper"],
            mode="lines", name="Upper bound",
            line=dict(color="rgba(154,164,178,0.5)", width=1, dash="dot"),
            hovertemplate="Upper: %{y:,.2f}<extra></extra>",
        ))
        fig.add_trace(go.Scatter(
            x=bounds_df["Date"], y=bounds_df["Lower"],
            mode="lines", name="Lower bound",
            line=dict(color="rgba(154,164,178,0.5)", width=1, dash="dot"),
            fill="tonexty", fillcolor="rgba(154,164,178,0.10)",
            hovertemplate="Lower: %{y:,.2f}<extra></extra>",
        ))

    # Original Sales History (solid)
    fig.add_trace(go.Scatter(
        x=key_df["Date"], y=key_df["Value"],
        mode="lines+markers", name="Sales History (original)",
        line=dict(color=SERIES_STYLE[SALES_HISTORY_LABEL]["color"], width=2.4,
                  shape="spline", smoothing=0.5),
        marker=dict(size=6),
        hovertemplate="<b>Sales History</b><br>%{x|%b %Y}"
                      "<br>Value: %{y:,.2f} kg<extra></extra>",
    ))

    # Cleansed history (solid, distinct colour)
    fig.add_trace(go.Scatter(
        x=key_df["Date"], y=key_df["Filtered"],
        mode="lines+markers", name="Cleansed history (corrected)",
        line=dict(color="#51cf66", width=2.4, shape="spline", smoothing=0.5),
        marker=dict(size=5, symbol="square"),
        hovertemplate="<b>Cleansed</b><br>%{x|%b %Y}"
                      "<br>Value: %{y:,.2f} kg<extra></extra>",
    ))

    # History For Forecast (kept as-is) — only if available
    if hff_df is not None and not hff_df.empty:
        fig.add_trace(go.Scatter(
            x=hff_df["Date"], y=hff_df["Value"],
            mode="lines", name="History For Forecast (unchanged)",
            line=dict(color=SERIES_STYLE[HISTORY_FOR_FORECAST_LABEL]["color"],
                      width=1.8),
            hovertemplate="<b>History For Forecast</b><br>%{x|%b %Y}"
                          "<br>Value: %{y:,.2f} kg<extra></extra>",
        ))

    # Outlier markers (on the original series)
    out_df = key_df[key_df["IsOutlier"]]
    if not out_df.empty:
        fig.add_trace(go.Scatter(
            x=out_df["Date"], y=out_df["Value"],
            mode="markers", name="Outlier (corrected)",
            marker=dict(color="#ff6b6b", size=13, symbol="x-thin",
                        line=dict(color="#ff6b6b", width=3)),
            hovertemplate="<b>OUTLIER</b><br>%{x|%b %Y}<br>"
                          "Value: %{y:,.2f}<extra></extra>",
        ))

    param_label = f"k = {param}" if method == "IQR" else f"n_sigma = {param}"
    dark_layout(
        fig,
        title=f"{key} — {method} method ({param_label})",
        xaxis_title="Months", yaxis_title="Demand volume in kgs",
        height=540,
        legend=dict(orientation="h", yanchor="top", y=-0.32,
                    xanchor="center", x=0.5, bgcolor="rgba(0,0,0,0)",
                    font=dict(color=DARK_TEXT)),
        margin=dict(t=70, l=60, r=30, b=150),
    )
    add_time_controls(fig)
    st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False})

    n_out = int(key_df["IsOutlier"].sum())
    if n_out:
        st.warning(
            f"⚠️ {n_out} outlier observation(s) detected and corrected for this Key."
        )
    else:
        st.success("✅ No outliers detected for this Key.")


# ---------------------------------------------------------------------------
# UI – Tab 3: Year-over-Year Seasonality (Change 4)
# ---------------------------------------------------------------------------
def render_seasonality_tab(long_df: pd.DataFrame) -> None:
    st.subheader("📅 Year-over-Year Seasonality")
    st.caption(
        "Compare the **monthly seasonal shape** of the last 3 historical "
        "years of Sales History against the forecast period. Each line is a "
        "seasonal index: a value of 1.10 means that month runs ~10% above "
        "that period's average. Use it to check whether the forecast carries "
        "the same seasonality as recent history."
    )

    # ---- Filter strip (same six as the outlier tab) ------------------------
    season_filter_cols = [c for c in FILTER_COLUMNS if c != "Data"]
    with st.container(border=True):
        st.markdown("**🔎 Filters** — leave empty to include everything. Filters cascade: each one narrows the choices in the others (Excel-slicer style).")
        selections = render_filter_strip(
            long_df, season_filter_cols, key_prefix="season",
        )

    filtered = apply_filters(long_df, selections)
    if filtered.empty:
        st.warning("No data matches the current filter combination.")
        return

    boundary = history_forecast_boundary(filtered)

    sales = filtered[filtered["Data"] == SALES_HISTORY_LABEL]
    sales_grp = (sales.dropna(subset=["Value"])
                 .groupby("Date", as_index=False)["Value"].sum())
    forecast = filtered[filtered["Data"] == STAT_FORECAST_LABEL]
    fcst_grp = (forecast.dropna(subset=["Value"])
                .groupby("Date", as_index=False)["Value"].sum())

    analysis = seasonality_analysis(sales_grp, fcst_grp, boundary, n_years=3)

    if analysis["hist_profile"] is None:
        st.info("Not enough historical Sales-History data to compute seasonality.")
        return

    months = list(range(1, 13))
    month_labels = [pd.Timestamp(2000, m, 1).strftime("%b") for m in months]

    fig = go.Figure()

    # Per-year historical profiles (thin, muted)
    year_palette = ["#4dabf7", "#3bc9db", "#9775fa"]
    for i, (yr, prof) in enumerate(sorted(analysis["per_year"].items())):
        if prof is None:
            continue
        fig.add_trace(go.Scatter(
            x=month_labels, y=prof.reindex(months).values,
            mode="lines+markers", name=f"History {yr}",
            line=dict(color=year_palette[i % len(year_palette)], width=1.6),
            marker=dict(size=5),
            legendgroup="History",
            hovertemplate=f"History {yr}<br>%{{x}}<br>Index: %{{y:.2f}}<extra></extra>",
        ))

    # Average historical profile (bold)
    fig.add_trace(go.Scatter(
        x=month_labels, y=analysis["hist_profile"].reindex(months).values,
        mode="lines+markers", name="History (3-yr avg)",
        line=dict(color="#4da3ff", width=3.4),
        marker=dict(size=7),
        legendgroup="History",
        hovertemplate="History 3-yr avg<br>%{x}<br>Index: %{y:.2f}<extra></extra>",
    ))

    # Forecast profile (bold, contrasting)
    if analysis["fcst_profile"] is not None:
        fig.add_trace(go.Scatter(
            x=month_labels, y=analysis["fcst_profile"].reindex(months).values,
            mode="lines+markers", name="Forecast",
            line=dict(color="#ffa94d", width=3.4),
            marker=dict(size=7, symbol="diamond"),
            legendgroup="Forecast",
            hovertemplate="Forecast<br>%{x}<br>Index: %{y:.2f}<extra></extra>",
        ))

    # Reference line at index = 1.0 (the period average)
    fig.add_hline(y=1.0, line=dict(color=DARK_MUTED, width=1, dash="dot"),
                  annotation_text="period average",
                  annotation_font=dict(color=DARK_MUTED, size=11))

    dark_layout(
        fig,
        title="Monthly seasonal index — History vs Forecast",
        xaxis_title="Month", yaxis_title="Seasonal index (1.0 = average)",
        height=520,
        legend=dict(orientation="h", yanchor="top", y=-0.28,
                    xanchor="center", x=0.5, bgcolor="rgba(0,0,0,0)",
                    font=dict(color=DARK_TEXT)),
        margin=dict(t=70, l=60, r=30, b=120),
    )
    # Seasonality x-axis is categorical months — no rangeslider needed.
    fig.update_xaxes(showspikes=True)
    fig.update_yaxes(tickformat=".2f")
    st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False})

    # ---- Interpretation -----------------------------------------------------
    status, message = interpret_seasonality(analysis)
    st.markdown("#### 🧭 Seasonality interpretation")
    if status == "ok":
        st.success(message)
    elif status == "warn":
        st.warning(message)
    else:
        st.info(message)

    if analysis["correlation"] is not None:
        st.caption(
            f"Based on {analysis['n_hist_years']} historical year(s). "
            "Correlation ranges from −1 (opposite shape) to +1 (identical "
            "shape); values ≥ 0.7 indicate the forecast preserves the "
            "historical seasonal pattern well."
        )


# ---------------------------------------------------------------------------
# Main entry-point
# ---------------------------------------------------------------------------
def main() -> None:
    inject_dark_theme_css()

    # ---- Authentication gate ------------------------------------------------
    # Nothing below this block (including the file uploader) renders until
    # the user signs in with a valid username and password.
    if not st.session_state.get("authenticated", False):
        render_login_screen()
        return  # unreachable (render_login_screen calls st.stop), kept for clarity

    # ---- Signed-in header ---------------------------------------------------
    auth_user = st.session_state.get("auth_user", "")
    display_name = USER_DISPLAY_NAMES.get(auth_user, auth_user.capitalize())
    head_l, head_r = st.columns([8, 2])
    with head_l:
        st.title("📊 Forecast vs Actuals — Interactive Dashboard")
        st.caption(f"👤 Signed in as **{display_name}**")
    with head_r:
        st.write("")  # vertical spacing to align the button with the title
        if st.button("🚪 Log out", key="auth::logout", use_container_width=True):
            logout()

    if "current_file_id" not in st.session_state:
        st.session_state["current_file_id"] = None

    uploaded = st.file_uploader(
        "Upload the **Forecast vs Actuals** input Excel file (.xlsx)",
        type=["xlsx"], accept_multiple_files=False,
        help="The file name typically ends with the reporting month, "
             "e.g. 'Forecast_vs_Actuals_Cleaning_and_Personal_Care_April.xlsx'.",
    )

    if uploaded is None:
        if st.session_state["current_file_id"] is not None:
            st.session_state["current_file_id"] = None
            reset_app_state()
        st.info(
            "👆 Drag-and-drop or browse to upload the Forecast vs Actuals "
            "Excel file to get started."
        )
        st.markdown(
            "**This dashboard provides:**\n"
            "1. An interactive replica of the *Forecast vs Actuals* pivot "
            "chart with all original slicer filters on top — clearly "
            "splitting **History** vs **Forecast** series, with optional "
            "**trend lines** and an automatic trend interpretation.\n"
            "2. An **outlier detection & correction** view (choose **IQR** or "
            "**Sigma**) for the **Sales History** of every Key, producing a "
            "cleansed history alongside the original.\n"
            "3. A **year-over-year seasonality** view comparing the forecast's "
            "seasonal shape against the last 3 historical years, with an "
            "automatic interpretation of forecast quality."
        )
        st.stop()

    file_id = f"{uploaded.name}::{uploaded.size}"
    if st.session_state["current_file_id"] != file_id:
        reset_app_state()
        st.session_state["current_file_id"] = file_id

    file_bytes = uploaded.getvalue()
    try:
        long_df, _ = load_excel(file_bytes, uploaded.name)
    except ValueError as exc:
        st.error(f"❌ {exc}")
        st.stop()
    except Exception as exc:  # noqa: BLE001
        st.error(f"❌ Unexpected error while loading the file: {exc}")
        st.stop()

    if long_df.empty:
        st.error("The uploaded file contains no usable rows.")
        st.stop()

    tab1, tab2, tab3 = st.tabs([
        "📈 Forecast vs Actuals",
        "🚨 Outlier Detection & Correction",
        "📅 Seasonality (YoY)",
    ])
    with tab1:
        render_dashboard_tab(long_df, uploaded.name)
    with tab2:
        render_outlier_tab(long_df)
    with tab3:
        render_seasonality_tab(long_df)


if __name__ == "__main__":
    main()
