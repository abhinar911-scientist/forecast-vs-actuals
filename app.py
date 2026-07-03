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
# Identifier columns in the RAW Arkieva input file (first sheet). Note this
# raw file does NOT contain a pre-built "Key" or a separate "Material code"
# column — the app derives both (see load_excel). "Arkieva Review Req"
# replaces the old "Stat Flag", and "Arkieva Active Status" is new.
RAW_ID_COLUMNS: List[str] = [
    "Business Line",
    "Material",
    "Ship To Sub Region",
    "Arkieva ABC",
    "Arkieva Pattern",
    "Arkieva Active Status",
    "Arkieva Review Req",
    "Data",
]

# Columns present AFTER load_excel derives Key and Material code. "Key" and
# "Material code" are placed at the front of the frame.
ID_COLUMNS: List[str] = [
    "Key",
    "Business Line",
    "Material",
    "Material code",
    "Ship To Sub Region",
    "Arkieva ABC",
    "Arkieva Pattern",
    "Arkieva Active Status",
    "Arkieva Review Req",
    "Data",
]

ACTIVE_STATUS_COL = "Arkieva Active Status"
# By default the app shows only "Active" and "Sparse" keys.
DEFAULT_ACTIVE_STATUSES = ["Active", "Sparse"]

REVIEW_REQ_COL = "Arkieva Review Req"
# A material + Ship To Sub Region combination is considered to be running on
# the Statistical Forecast only when it does NOT require review and its
# active status is Active or Sparse.
ON_STAT_REVIEW_VALUE = "No"
ON_STAT_STATUSES = ["Active", "Sparse"]

# Order matches the Excel slicer order. "Stat Flag" is replaced by
# "Arkieva Review Req"; "Arkieva Active Status" is added (defaulted to
# Active + Sparse).
FILTER_COLUMNS: List[str] = [
    "Business Line",
    "Arkieva ABC",
    "Ship To Sub Region",
    "Material",
    "Arkieva Review Req",
    "Arkieva Active Status",
    "Arkieva Pattern",
    "Data",
]

SALES_HISTORY_LABEL = "Sales History (kg)"
HISTORY_FOR_FORECAST_LABEL = "History For Forecast (kg)"
STAT_FORECAST_LABEL = "Statistical Forecast (kg)"

# STF (Statistical Forecast Committed) current vs Lag 1 — used by the
# "STF Variation & Exceptions" tab. The current committed line is normalised
# to STAT_FORECAST_LABEL on load, so that label IS the current STF here. The
# Lag 1 line keeps its own label (it must NOT be normalised away).
STF_CURRENT_LABEL = STAT_FORECAST_LABEL
STF_LAG1_LABEL = "Statistical Forecast Committed Lag 1 (kg)"

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
def _extract_material_code(material) -> str:
    """Pull the numeric material code out of a Material string.

    In the raw Arkieva file the Material column embeds the code after a
    double underscore, e.g. '2-ETHYL HEXANOL BULK__3000924' -> '3000924'.
    Falls back to the whole string when there is no '__'.
    """
    s = str(material).strip()
    if "__" in s:
        return s.rsplit("__", 1)[-1].strip()
    return s


@st.cache_data(show_spinner=False)
def load_excel(file_bytes: bytes, file_name: str) -> Tuple[pd.DataFrame, List[pd.Timestamp]]:
    """Load and reshape the uploaded workbook.

    Accepts the raw Arkieva export (first sheet) whose identifier columns are
    RAW_ID_COLUMNS. The app derives:
      * **Material code** — extracted from the Material string (after '__').
      * **Key** — ``Business Line_<material code>_Ship To Sub Region`` —
        placed at the front of the frame, mirroring the legacy input's Key.

    Older files that already contain a built ``Key`` and a ``Stat Flag``
    column are also accepted (Stat Flag is mapped to ``Arkieva Review Req``;
    a missing ``Arkieva Active Status`` defaults to 'Active').

    Returns a long-format DataFrame ['Key', ..., 'Date', 'Value'] together
    with the sorted list of months found in the file. Raises ValueError
    with a user-friendly message if the file is invalid.
    """
    try:
        raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0)
    except Exception as exc:  # noqa: BLE001
        raise ValueError(f"Could not read the Excel file: {exc}") from exc

    raw = raw.copy()

    # ---- Backward compatibility shims --------------------------------------
    # Legacy files used "Stat Flag" instead of "Arkieva Review Req".
    if "Arkieva Review Req" not in raw.columns and "Stat Flag" in raw.columns:
        raw = raw.rename(columns={"Stat Flag": "Arkieva Review Req"})
    # Legacy files have no active-status column; treat every row as Active.
    if ACTIVE_STATUS_COL not in raw.columns:
        raw[ACTIVE_STATUS_COL] = "Active"

    # ---- Validate the raw identifier columns -------------------------------
    missing = [c for c in RAW_ID_COLUMNS if c not in raw.columns]
    if missing:
        raise ValueError(
            "The uploaded file is missing required column(s): "
            + ", ".join(missing)
            + ". Make sure you uploaded the raw Arkieva 'Forecast vs Actuals' "
            "export (data on the first sheet)."
        )

    # ---- Derive Material code and Key --------------------------------------
    if "Material code" not in raw.columns:
        raw["Material code"] = raw["Material"].apply(_extract_material_code)
    if "Key" not in raw.columns:
        raw["Key"] = (
            raw["Business Line"].astype(str).str.strip() + "_"
            + raw["Material code"].astype(str).str.strip() + "_"
            + raw["Ship To Sub Region"].astype(str).str.strip()
        )

    # ---- Identify month/date columns ---------------------------------------
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
            "'2023-06-01', '2023-07-01', ..."
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

    # Normalise the statistical-forecast label. Some files use
    # "Statistical Forecast Committed (kg)"; map every alias to the canonical
    # "Statistical Forecast (kg)" so the rest of the app is label-agnostic.
    long_df["Data"] = long_df["Data"].astype(str).str.strip()
    long_df["Data"] = long_df["Data"].replace(
        "Statistical Forecast Committed (kg)", STAT_FORECAST_LABEL
    )

    # Normalise the Arkieva Review Req values to readable strings (the raw
    # file stores booleans True/False).
    long_df["Arkieva Review Req"] = (
        long_df["Arkieva Review Req"].map(
            {True: "Yes", False: "No", "True": "Yes", "False": "No"}
        ).fillna(long_df["Arkieva Review Req"].astype(str).str.strip())
    )

    for c in FILTER_COLUMNS:
        if long_df[c].dtype == object:
            long_df[c] = long_df[c].astype(str).str.strip()
        else:
            long_df[c] = long_df[c].astype(str)

    # Performance: downcast the value column to a 32-bit float to roughly
    # halve memory and speed up the numeric aggregations every tab performs.
    # (String columns are left as object on purpose: converting them to
    # 'category' triggers pandas' "all categories" groupby behaviour, which
    # would make per-Key loops iterate over absent Keys.)
    long_df["Value"] = long_df["Value"].astype("float32")

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
    default_selections: Optional[dict] = None,
) -> dict:
    """Render a cascading multi-select filter strip above the chart.

    Behavior matches Excel slicers:
    * Each filter shows only values that are consistent with the
      selections in the *other* filters.
    * Selecting values in one filter narrows the options in every other.
    * Stale selections are silently pruned so the UI never crashes.

    The pruning is single-pass and proceeds in the declared filter order:
    earlier filters take precedence over later ones when they conflict.

    ``default_selections`` maps a column name to a list of values that should
    be pre-selected on the FIRST render for this ``key_prefix`` (e.g. the
    Arkieva Active Status defaulting to Active + Sparse). The defaults are
    applied only once; afterwards the user's choices — including clearing a
    filter — are respected. The 🔄 Reset button restores the defaults.

    Each tab passes a unique ``key_prefix`` so filter state is
    independent across tabs and Streamlit doesn't see duplicate keys.
    """
    default_selections = default_selections or {}

    # Apply one-time defaults: only seed a filter's state the very first time
    # this strip is rendered for this key_prefix (tracked by an init flag).
    init_flag = f"{key_prefix}::__initialized__"
    if not st.session_state.get(init_flag, False):
        for col, default_vals in default_selections.items():
            state_key = f"{key_prefix}::{col}"
            if state_key not in st.session_state:
                # keep only defaults that actually exist in the data
                valid = [v for v in default_vals if v in set(df[col].astype(str))]
                st.session_state[state_key] = valid
        st.session_state[init_flag] = True

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
                # Reset restores defaults (empty if none configured).
                default_vals = default_selections.get(col_name, [])
                valid = [v for v in default_vals
                         if v in set(df[col_name].astype(str))]
                st.session_state[f"{key_prefix}::{col_name}"] = valid
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
    # Avoid a redundant pd.to_datetime when the input is already datetime
    # (it is, post-load); this function is called thousands of times in the
    # per-Key anomaly scan, so the conversion cost adds up.
    dt = dates if pd.api.types.is_datetime64_any_dtype(
        getattr(dates, "dtype", None)) else pd.to_datetime(dates)
    vals = np.asarray(values, dtype="float64")
    mask = ~np.isnan(vals)
    if not mask.any():
        return None
    months = pd.DatetimeIndex(dt)[mask].month
    vals = vals[mask]
    overall = vals.mean()
    if overall == 0 or np.isnan(overall):
        return None
    monthly = pd.Series(vals, index=months).groupby(level=0).mean()
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
# Anomaly Summary engine
# ---------------------------------------------------------------------------
# Reason codes and the Demand-Planner guidance that goes with each. The
# guidance is intentionally action-oriented: root cause + recommended next
# step from a demand-planning point of view.
REASON_CODES = {
    "TREND_REVERSAL": {
        "label": "Trend reversal",
        "short": "Forecast trends opposite to history.",
        "root_cause": (
            "Forecast trends opposite to recent history. Usually caused by a "
            "few recent months breaking the pattern, end-of-history outliers "
            "pulling the fit, or a model change at regeneration."
        ),
        "next_steps": (
            "Confirm the reversal on the chart, check the last 3–6 months for "
            "outliers/one-off events; if not backed by a real event, override "
            "toward the historical direction or re-run after cleansing."
        ),
    },
    "TREND_MISMATCH": {
        "label": "Trend magnitude mismatch",
        "short": "Same direction, very different slope.",
        "root_cause": (
            "Same direction as history but a disproportionate rate of change. "
            "Often short-history overfitting, aggressive trend smoothing, or "
            "recent spikes amplifying the slope."
        ),
        "next_steps": (
            "Compare history vs forecast slope, sanity-check end-of-horizon "
            "volume; if unrealistic, dampen the trend or cap growth and re-run."
        ),
    },
    "SEASONALITY_LOSS": {
        "label": "Seasonality not captured",
        "short": "History seasonal, forecast (near) flat.",
        "root_cause": (
            "History repeats an intra-year pattern the forecast doesn't "
            "reproduce. Usually too little history, a non-seasonal model, or "
            "seasonality washed out by noise/outliers."
        ),
        "next_steps": (
            "Confirm peak/trough months on the Seasonality tab; if history is "
            "reliable, apply the seasonal profile or switch to a seasonal "
            "model and re-run."
        ),
    },
    "IQR_OUTLIERS": {
        "label": "Outliers in sales history (IQR)",
        "short": "Several points violate IQR fences.",
        "root_cause": (
            "Sales history has points outside the IQR fences — promotions, "
            "bulk orders, data errors, stock-outs or returns — which can "
            "distort the fitted level and trend."
        ),
        "next_steps": (
            "Review flagged points on the Outlier tab against known events, "
            "replace genuine anomalies with the cleansed value and re-run; "
            "model recurring causes explicitly."
        ),
    },
}


def analyze_key_anomalies(
    key_long_df: pd.DataFrame,
    boundary: Optional[pd.Timestamp],
    iqr_k: float = 1.5,
    trend_mismatch_ratio: float = 3.0,
    seasonality_corr_threshold: float = 0.3,
    seasonality_min_amplitude: float = 0.10,
) -> Optional[dict]:
    """Analyse a single Key (already filtered to one Key) and return its
    anomaly profile: which reason codes apply, a severity score, and the
    supporting metrics used in the drill-down.

    Returns None if there isn't enough data to assess the Key.
    """
    if key_long_df.empty or boundary is None:
        return None

    sales = (key_long_df[key_long_df["Data"] == SALES_HISTORY_LABEL]
             .dropna(subset=["Value"]))
    sales = sales[sales["Date"] <= boundary]
    sales_grp = sales.groupby("Date", as_index=False)["Value"].sum()

    fcst = (key_long_df[key_long_df["Data"] == STAT_FORECAST_LABEL]
            .dropna(subset=["Value"]))
    fcst = fcst[fcst["Date"] > boundary]
    fcst_grp = fcst.groupby("Date", as_index=False)["Value"].sum()

    # Need some history to say anything meaningful.
    nonzero_hist = sales_grp[sales_grp["Value"] != 0]
    if len(nonzero_hist) < 6:
        return None

    reasons = []          # list of dicts: {code, severity, detail}
    metrics = {}

    # ---- Trend analysis ----------------------------------------------------
    hist_trend = fit_trend(sales_grp["Date"], sales_grp["Value"])
    fcst_trend = fit_trend(fcst_grp["Date"], fcst_grp["Value"]) if len(fcst_grp) >= 2 else None
    metrics["hist_trend"] = hist_trend
    metrics["fcst_trend"] = fcst_trend

    if hist_trend is not None and fcst_trend is not None:
        h = hist_trend["pct_per_year"]
        f = fcst_trend["pct_per_year"]
        metrics["hist_pct_per_year"] = h
        metrics["fcst_pct_per_year"] = f

        if h is not None and f is not None:
            flat = 2.0  # within ±2%/yr is "flat"
            h_dir = 0 if abs(h) <= flat else (1 if h > 0 else -1)
            f_dir = 0 if abs(f) <= flat else (1 if f > 0 else -1)

            # Trend reversal: directions strictly opposite (one up, one down)
            if h_dir != 0 and f_dir != 0 and h_dir != f_dir:
                # severity scales with the combined magnitude of the swing
                sev = min(100.0, (abs(h) + abs(f)) / 2.0)
                reasons.append({
                    "code": "TREND_REVERSAL",
                    "severity": 60.0 + 0.4 * sev,   # high base — most serious
                    "detail": f"History {h:+.1f}%/yr vs forecast {f:+.1f}%/yr",
                })
            else:
                # Trend magnitude mismatch (same direction but disparate rate)
                if abs(h) > 1e-6:
                    ratio = abs(f) / abs(h) if h != 0 else np.inf
                    metrics["trend_ratio"] = ratio
                    if ratio >= trend_mismatch_ratio or (ratio > 0 and ratio <= 1.0 / trend_mismatch_ratio):
                        # How far from parity (1.0), in log space, capped
                        far = abs(np.log10(ratio)) if ratio > 0 else 2.0
                        sev = min(100.0, far * 40.0)
                        reasons.append({
                            "code": "TREND_MISMATCH",
                            "severity": 30.0 + 0.4 * sev,
                            "detail": (f"Forecast slope {f:+.1f}%/yr vs history "
                                       f"{h:+.1f}%/yr (~{ratio:.1f}× rate)"),
                        })

    # ---- Seasonality analysis ----------------------------------------------
    analysis = seasonality_analysis(sales_grp, fcst_grp, boundary, n_years=3)
    metrics["seasonality"] = analysis
    hp = analysis["hist_profile"]
    fp = analysis["fcst_profile"]
    if hp is not None:
        hist_amp = float(np.nanstd(hp.values))
        metrics["hist_seasonal_amplitude"] = hist_amp
        corr = analysis["correlation"]
        metrics["seasonal_corr"] = corr
        fcst_amp = float(np.nanstd(fp.values)) if fp is not None else 0.0
        metrics["fcst_seasonal_amplitude"] = fcst_amp

        # History is meaningfully seasonal …
        if hist_amp >= seasonality_min_amplitude and fp is not None:
            seasonality_lost = (
                (corr is None and fcst_amp <= 0.02) or
                (corr is not None and corr < seasonality_corr_threshold)
            )
            if seasonality_lost:
                # severity scales with how strong history seasonality is and
                # how poor the match is.
                match_gap = 1.0 - (corr if corr is not None else 0.0)
                match_gap = max(0.0, min(2.0, match_gap))
                sev = min(100.0, hist_amp * 100.0 * match_gap)
                reasons.append({
                    "code": "SEASONALITY_LOSS",
                    "severity": 25.0 + 0.5 * sev,
                    "detail": (f"History seasonal amplitude {hist_amp:.2f}, "
                               f"forecast match "
                               f"{('corr ' + format(corr, '+.2f')) if corr is not None else 'flat'}"),
                })

    # ---- IQR outliers ------------------------------------------------------
    iqr_res = detect_and_correct_outliers(sales_grp, method="IQR", k=iqr_k)
    n_out = int(iqr_res["IsOutlier"].sum())
    metrics["iqr_outlier_count"] = n_out
    metrics["iqr_result"] = iqr_res
    if n_out > 0:
        # relative magnitude of the worst outlier vs the series median
        med = float(np.median(sales_grp["Value"].replace(0, np.nan).dropna())) \
            if len(sales_grp) else 0.0
        worst = 0.0
        if med and not np.isnan(med):
            outvals = iqr_res.loc[iqr_res["IsOutlier"], "Value"].to_numpy()
            if outvals.size:
                worst = float(np.max(np.abs(outvals - med)) / abs(med))
        metrics["iqr_worst_rel"] = worst
        sev = min(100.0, n_out * 12.0 + worst * 20.0)
        reasons.append({
            "code": "IQR_OUTLIERS",
            "severity": 20.0 + 0.5 * sev,
            "detail": f"{n_out} point(s) violate IQR fences",
        })

    if not reasons:
        return None

    # Primary reason = highest-severity; overall score = max severity, with a
    # small bonus for having multiple concurrent issues.
    reasons_sorted = sorted(reasons, key=lambda r: r["severity"], reverse=True)
    primary = reasons_sorted[0]
    overall = primary["severity"] + 3.0 * (len(reasons_sorted) - 1)

    return {
        "reasons": reasons_sorted,
        "primary_code": primary["code"],
        "primary_detail": primary["detail"],
        "all_codes": [r["code"] for r in reasons_sorted],
        "severity": round(float(overall), 1),
        "metrics": metrics,
        "sales_grp": sales_grp,
        "fcst_grp": fcst_grp,
    }


@st.cache_data(show_spinner="Scanning Keys for anomalies…")
def build_anomaly_summary(filtered_df: pd.DataFrame,
                          top_n: int = 20) -> pd.DataFrame:
    """Build the per-Key anomaly table for the current filter selection.

    Returns one row per anomalous Key with its Ship To Sub Region, primary
    reason code, all reason codes, severity, supporting detail, and the key
    metrics used in the drill-down. Only the top ``top_n`` Keys (by severity)
    per Ship To Sub Region are returned.
    """
    boundary = history_forecast_boundary(filtered_df)
    if boundary is None:
        return pd.DataFrame()

    # ---- Performance pre-filter -------------------------------------------
    # analyze_key_anomalies needs >= 6 non-zero Sales-History months (up to the
    # boundary) to say anything. Compute that in one vectorised pass and skip
    # the expensive per-Key trend/seasonality/IQR work for Keys that can't
    # qualify — this is the dominant cost in the scan.
    sh = filtered_df[(filtered_df["Data"] == SALES_HISTORY_LABEL)
                     & (filtered_df["Date"] <= boundary)]
    sh = sh[sh["Value"].notna() & (sh["Value"] != 0)]
    eligible = (sh.groupby("Key")["Date"].nunique())
    eligible_keys = set(eligible[eligible >= 6].index)
    if not eligible_keys:
        return pd.DataFrame()
    work = filtered_df[filtered_df["Key"].isin(eligible_keys)]

    rows = []
    for key, grp in work.groupby("Key"):
        profile = analyze_key_anomalies(grp, boundary)
        if profile is None:
            continue
        first = grp.iloc[0]
        row = {
            "Key": key,
            "Ship To Sub Region": first.get("Ship To Sub Region", "—"),
            "Business Line": first.get("Business Line", "—"),
            "Material": first.get("Material", "—"),
            "Arkieva ABC": first.get("Arkieva ABC", "—"),
            "Arkieva Pattern": first.get("Arkieva Pattern", "—"),
            "Primary reason": REASON_CODES[profile["primary_code"]]["label"],
            "Primary code": profile["primary_code"],
            "All reasons": ", ".join(
                REASON_CODES[c]["label"] for c in profile["all_codes"]),
            "All codes": profile["all_codes"],
            "Severity": profile["severity"],
            "Detail": profile["primary_detail"],
            "IQR outliers": profile["metrics"].get("iqr_outlier_count", 0),
            "Hist %/yr": profile["metrics"].get("hist_pct_per_year"),
            "Fcst %/yr": profile["metrics"].get("fcst_pct_per_year"),
            "Seasonal corr": profile["metrics"].get("seasonal_corr"),
        }
        rows.append(row)

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)
    # Keep the top_n most severe per Ship To Sub Region.
    df = (df.sort_values(["Ship To Sub Region", "Severity"],
                         ascending=[True, False])
            .groupby("Ship To Sub Region", group_keys=False)
            .head(top_n)
            .reset_index(drop=True))
    return df


# ---------------------------------------------------------------------------
# Statistical Forecast Adoption
# ---------------------------------------------------------------------------
def active_year_from_history(filtered_df: pd.DataFrame) -> Optional[int]:
    """Determine the *active year* — the calendar year of the last active
    month — by looking at the Sales History (kg) series.

    The last active month is the last month that has real (non-null,
    non-zero) Sales History. Adoption is computed using only the Statistical
    Forecast of that year.
    """
    boundary = history_forecast_boundary(filtered_df)
    if boundary is None:
        return None
    return int(pd.Timestamp(boundary).year)


@st.cache_data(show_spinner=False)
def compute_stat_adoption(
    filtered_df: pd.DataFrame,
    group_col: Optional[str] = None,
    active_year: Optional[int] = None,
) -> pd.DataFrame:
    """Compute the Statistical-Forecast adoption %, by fiscal year, for the
    **active year onwards**, rolled up from the granular
    BL + Material + Ship To Sub Region + Arkieva Active Status level.

    A granular combination is "on the Statistical Forecast" when
    ``Arkieva Review Req == 'No'`` AND ``Arkieva Active Status`` is Active or
    Sparse. Adoption is always measured against the **Ship To Sub Region**
    total forecast for that year:

        Adoption = Σ on-stat forecast
                   ──────────────────────────────
                   Σ total forecast (same scope)

    The denominator is the total Statistical Forecast of the scope being
    reported (overall, per Business Line, or per Ship To Sub Region); higher
    levels roll up the granular numerators and denominators.

    Only years **>= the active year** are included. The active year is the
    calendar year of the last active month in the Sales History (derived via
    ``active_year_from_history`` when not supplied) — past years are excluded
    because their forecast is historical and not relevant to adoption going
    forward.

    Returns a tidy DataFrame:
    [Year, (group_col), OnStatVolume, TotalVolume, AdoptionPct, ActiveYear].
    """
    base_cols = ["Year"] + ([group_col] if group_col else []) + \
        ["OnStatVolume", "TotalVolume", "AdoptionPct", "ActiveYear"]

    if active_year is None:
        active_year = active_year_from_history(filtered_df)

    stat = filtered_df[filtered_df["Data"] == STAT_FORECAST_LABEL].copy()
    stat = stat.dropna(subset=["Value"])
    if stat.empty or active_year is None:
        return pd.DataFrame(columns=base_cols)

    # Keep the active year and every future year (drop past years).
    stat["Year"] = pd.to_datetime(stat["Date"]).dt.year
    stat = stat[stat["Year"] >= active_year]
    if stat.empty:
        return pd.DataFrame(columns=base_cols)

    stat["_on_stat"] = (
        (stat[REVIEW_REQ_COL] == ON_STAT_REVIEW_VALUE)
        & (stat[ACTIVE_STATUS_COL].isin(ON_STAT_STATUSES))
    )
    stat["_on_stat_vol"] = np.where(stat["_on_stat"], stat["Value"], 0.0)

    group_keys = ["Year"] + ([group_col] if group_col else [])
    agg = stat.groupby(group_keys, as_index=False).agg(
        OnStatVolume=("_on_stat_vol", "sum"),
        TotalVolume=("Value", "sum"),
    )
    agg["AdoptionPct"] = np.where(
        agg["TotalVolume"] != 0,
        agg["OnStatVolume"] / agg["TotalVolume"] * 100.0,
        0.0,
    )
    agg["ActiveYear"] = active_year
    agg = agg.sort_values(group_keys).reset_index(drop=True)
    return agg[base_cols]


@st.cache_data(show_spinner=False)
def stat_adoption_material_table(
    filtered_df: pd.DataFrame,
    active_year: Optional[int] = None,
) -> pd.DataFrame:
    """Build the granular table (BL + Material + Ship To Sub Region + Active
    Status) of combinations that are **on** the Statistical Forecast from the
    active year onwards, with each combination's forecast volume per fiscal
    year and its share of the Ship To Sub Region's total forecast.

    Only combinations meeting the on-stat criteria (Review Req == No and
    Active/Sparse) are listed.
    """
    if active_year is None:
        active_year = active_year_from_history(filtered_df)
    if active_year is None:
        return pd.DataFrame()

    stat = filtered_df[filtered_df["Data"] == STAT_FORECAST_LABEL].copy()
    stat = stat.dropna(subset=["Value"])
    stat["Year"] = pd.to_datetime(stat["Date"]).dt.year
    stat = stat[stat["Year"] >= active_year]
    if stat.empty:
        return pd.DataFrame()

    # Region totals per year (denominator) — all forecast in the region.
    region_year_totals = (stat.groupby(["Ship To Sub Region", "Year"])["Value"]
                          .sum())

    on = stat[
        (stat[REVIEW_REQ_COL] == ON_STAT_REVIEW_VALUE)
        & (stat[ACTIVE_STATUS_COL].isin(ON_STAT_STATUSES))
    ].copy()
    if on.empty:
        return pd.DataFrame()

    id_cols = ["Business Line", "Material", "Material code",
               "Ship To Sub Region", "Arkieva ABC", "Arkieva Pattern",
               ACTIVE_STATUS_COL]
    id_cols = [c for c in id_cols if c in on.columns]

    # Per-year forecast volume for each granular combination.
    year_pivot = (on.groupby(id_cols + ["Year"], as_index=False)["Value"].sum()
                  .pivot_table(index=id_cols, columns="Year",
                               values="Value", aggfunc="sum", fill_value=0.0))
    year_pivot.columns = [f"FY {int(c)} (kg)" for c in year_pivot.columns]
    fy_cols = list(year_pivot.columns)
    year_pivot["Total forecast (kg)"] = year_pivot[fy_cols].sum(axis=1)
    out = year_pivot.reset_index()

    # Region total forecast across the active-year-onwards window (denominator)
    region_total = (region_year_totals.groupby("Ship To Sub Region").sum())
    out["Region total forecast (kg)"] = out["Ship To Sub Region"].map(region_total)
    out["% of region forecast"] = np.where(
        out["Region total forecast (kg)"] != 0,
        out["Total forecast (kg)"] / out["Region total forecast (kg)"] * 100.0,
        0.0,
    )
    out = out.sort_values("Total forecast (kg)", ascending=False).reset_index(drop=True)
    return out


# ---------------------------------------------------------------------------
# STF Variation & Exceptions engine
# ---------------------------------------------------------------------------
# The waterfall decomposes the net STF change by each row's **Arkieva Active
# Status** — the Demand Planner's own classification of where forecast is
# coming from (new items, sparse, obsolete, …). Every data row carries exactly
# one status, so the per-status deltas reconcile exactly to the net change.
# Display order for the buckets (statuses absent from the data are skipped):
STF_STATUS_ORDER = ["Active", "Active New", "New", "New Combination",
                    "Sparse", "Obsolete", "Inactive", "-"]


def stf_m4_month(filtered_df: pd.DataFrame) -> Optional[pd.Timestamp]:
    """Determine **Month 4 (M4)** — the start of the variation horizon.

    The "current month" is the last active month in Sales History (the
    history/forecast boundary). M4 is the 4th month after the current month
    (current + 4), e.g. current = June → M1 July, M2 Aug, M3 Sep, M4 Oct.
    """
    boundary = history_forecast_boundary(filtered_df)
    if boundary is None:
        return None
    return (pd.Timestamp(boundary) + pd.DateOffset(months=4)).normalize().replace(day=1)


def stf_horizon_months(
    all_months: List[pd.Timestamp],
    m4: pd.Timestamp,
    horizon: str = "fiscal_year",
    window: Optional[int] = None,
) -> List[pd.Timestamp]:
    """Return the list of months in the requested horizon, starting at M4.

    * ``horizon='fiscal_year'`` → M4 through December of M4's year.
    * ``horizon='next_12'``      → M4 plus the following 11 months (12 total).

    If ``window`` is given it caps the number of months returned (the sliding
    window from M4), so the user can shorten either horizon.
    """
    months_sorted = sorted(all_months)
    if m4 is None:
        return []
    if horizon == "fiscal_year":
        candidate = [m for m in months_sorted
                     if m >= m4 and m.year == m4.year and m.month <= 12]
    else:  # next_12
        future = [m for m in months_sorted if m >= m4]
        candidate = future[:12]
    if window is not None:
        candidate = candidate[:max(1, window)]
    return candidate


@st.cache_data(show_spinner=False)
def compute_stf_variance(
    filtered_df: pd.DataFrame,
    horizon_months: Tuple[pd.Timestamp, ...],
) -> pd.DataFrame:
    """Per-Key STF variance over the given horizon months.

    Variance = (Σ current committed − Σ lag-1 committed) / Σ lag-1 committed,
    summed across the horizon months. Returns one row per Key with the
    current/lag totals, signed STF Variance and Absolute STF Variance, plus
    descriptive columns, sorted by Absolute STF Variance (high → low).
    """
    cols = ["Key", "Business Line", "Material", "Ship To Sub Region",
            "Arkieva ABC", "Arkieva Pattern", "Arkieva Active Status",
            "CurrentSTF", "Lag1STF", "STF Variance", "Absolute STF Variance"]
    if filtered_df.empty or not horizon_months:
        return pd.DataFrame(columns=cols)

    hz = list(horizon_months)
    sub = filtered_df[
        filtered_df["Data"].isin([STF_CURRENT_LABEL, STF_LAG1_LABEL])
        & filtered_df["Date"].isin(hz)
    ].copy()
    if sub.empty:
        return pd.DataFrame(columns=cols)

    # Sum each series over the horizon, per Key.
    grp = (sub.groupby(["Key", "Data"], as_index=False)["Value"].sum()
           .pivot(index="Key", columns="Data", values="Value"))
    grp = grp.reindex(columns=[STF_CURRENT_LABEL, STF_LAG1_LABEL]).fillna(0.0)
    grp.columns = ["CurrentSTF", "Lag1STF"]
    grp = grp.reset_index()

    grp["STF Variance"] = np.where(
        grp["Lag1STF"] != 0,
        (grp["CurrentSTF"] - grp["Lag1STF"]) / grp["Lag1STF"],
        np.nan,
    )
    grp["Absolute STF Variance"] = grp["STF Variance"].abs()

    # Attach descriptors (first row per Key).
    desc_cols = ["Business Line", "Material", "Ship To Sub Region",
                 "Arkieva ABC", "Arkieva Pattern", "Arkieva Active Status"]
    desc_cols = [c for c in desc_cols if c in filtered_df.columns]
    desc = filtered_df.drop_duplicates(subset=["Key"]).set_index("Key")[desc_cols]
    out = grp.merge(desc, left_on="Key", right_index=True, how="left")

    out = out.dropna(subset=["STF Variance"])
    out = out.sort_values("Absolute STF Variance", ascending=False).reset_index(drop=True)
    return out[[c for c in cols if c in out.columns]]


@st.cache_data(show_spinner=False)
def compute_stf_month_variance(
    filtered_df: pd.DataFrame,
    keys: Tuple[str, ...],
    horizon_months: Tuple[pd.Timestamp, ...],
) -> pd.DataFrame:
    """Month-level STF variance for a set of Keys, for the heat-style table
    that highlights which months drive the variation.

    Returns long rows: [Key, Date, CurrentSTF, Lag1STF, STF Variance,
    Absolute STF Variance].
    """
    cols = ["Key", "Date", "CurrentSTF", "Lag1STF",
            "STF Variance", "Absolute STF Variance"]
    if filtered_df.empty or not horizon_months or not keys:
        return pd.DataFrame(columns=cols)

    hz = list(horizon_months)
    keyset = set(keys)
    sub = filtered_df[
        filtered_df["Key"].isin(keyset)
        & filtered_df["Data"].isin([STF_CURRENT_LABEL, STF_LAG1_LABEL])
        & filtered_df["Date"].isin(hz)
    ].copy()
    if sub.empty:
        return pd.DataFrame(columns=cols)

    piv = (sub.groupby(["Key", "Date", "Data"], as_index=False)["Value"].sum()
           .pivot_table(index=["Key", "Date"], columns="Data",
                        values="Value", fill_value=0.0))
    piv = piv.reindex(columns=[STF_CURRENT_LABEL, STF_LAG1_LABEL]).fillna(0.0)
    piv.columns = ["CurrentSTF", "Lag1STF"]
    piv = piv.reset_index()
    piv["STF Variance"] = np.where(
        piv["Lag1STF"] != 0,
        (piv["CurrentSTF"] - piv["Lag1STF"]) / piv["Lag1STF"],
        np.nan,
    )
    piv["Absolute STF Variance"] = piv["STF Variance"].abs()
    return piv[cols]


@st.cache_data(show_spinner=False)
def compute_stf_drivers(
    filtered_df: pd.DataFrame,
    horizon_months: Tuple[pd.Timestamp, ...],
    keys: Optional[Tuple[str, ...]] = None,
) -> pd.DataFrame:
    """Decompose the net STF change (current − lag1) over the horizon into
    **Arkieva Active Status** buckets for the waterfall chart.

    Every data row carries exactly one Arkieva Active Status (Active,
    Active New, New, New Combination, Sparse, Obsolete, Inactive, …), so
    Delta(status) = Σ current(status) − Σ lag1(status) and the bucket sums
    reconcile exactly to the net change. This lets Demand Planners read the
    waterfall in Arkieva's own vocabulary — e.g. forecast added by *Active
    New*/*New* items vs forecast lost from *Obsolete*/*Inactive* ones.

    Returns [Driver, Delta] where Driver is the status, ordered per
    STF_STATUS_ORDER (statuses with no data in the horizon are skipped).
    When ``keys`` is given the decomposition is restricted to those Keys
    (single-Key drill-down); otherwise it rolls up across ``filtered_df``.
    """
    cols = ["Driver", "Delta"]
    if filtered_df.empty or not horizon_months:
        return pd.DataFrame(columns=cols)

    hz = list(horizon_months)
    sub = filtered_df[
        filtered_df["Data"].isin([STF_CURRENT_LABEL, STF_LAG1_LABEL])
        & filtered_df["Date"].isin(hz)
    ]
    if keys:
        sub = sub[sub["Key"].isin(set(keys))]
    if sub.empty:
        return pd.DataFrame(columns=cols)

    # Σ per (status, series), then Delta = current − lag1 per status.
    grp = (sub.groupby([ACTIVE_STATUS_COL, "Data"], as_index=False)["Value"]
           .sum()
           .pivot(index=ACTIVE_STATUS_COL, columns="Data", values="Value"))
    grp = grp.reindex(columns=[STF_CURRENT_LABEL, STF_LAG1_LABEL]).fillna(0.0)
    grp["Delta"] = grp[STF_CURRENT_LABEL] - grp[STF_LAG1_LABEL]

    # Order buckets per STF_STATUS_ORDER; append any unexpected statuses at
    # the end so the decomposition always reconciles.
    present = list(grp.index)
    ordered = [s for s in STF_STATUS_ORDER if s in present] + \
              [s for s in present if s not in STF_STATUS_ORDER]
    rows = [{"Driver": s, "Delta": float(grp.loc[s, "Delta"])}
            for s in ordered]
    return pd.DataFrame(rows, columns=cols)


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
# UI – Tab: Anomaly Summary
# ---------------------------------------------------------------------------
def render_anomaly_tab(long_df: pd.DataFrame) -> None:
    st.subheader("🧭 Anomaly Summary")
    st.caption("Top anomalies per Ship To Sub Region, with reason codes. "
               "Drill down for root cause and next steps.")

    # ---- Filters (same set as the Forecast vs Actuals tab) -----------------
    with st.container(border=True):
        st.markdown("**🔎 Filters** — cascade like Excel slicers (each narrows the others). "
                    "*Arkieva Active Status* defaults to **Active + Sparse**; clear or change any filter as needed.")
        anom_filter_cols = [c for c in FILTER_COLUMNS if c != "Data"]
        selections = render_filter_strip(
            long_df, anom_filter_cols, key_prefix="anom",
            default_selections={ACTIVE_STATUS_COL: DEFAULT_ACTIVE_STATUSES})

    filtered = apply_filters(long_df, selections)
    if filtered.empty:
        st.warning("No data matches the current filter combination.")
        return

    top_n = st.slider(
        "Top anomalies per Ship To Sub Region", min_value=5, max_value=50,
        value=20, step=5, key="anom::top_n",
        help="How many of the most severe anomalous Keys to show for each "
             "Ship To Sub Region.",
    )

    summary = build_anomaly_summary(filtered, top_n=top_n)

    if summary.empty:
        st.success("No anomalies detected for the current filter selection. 🎉")
        with st.expander("How are anomalies detected?"):
            _render_reason_code_legend()
        return

    # ---- Headline metrics ---------------------------------------------------
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Anomalous Keys", f"{summary['Key'].nunique():,}")
    m2.metric("Sub Regions affected", f"{summary['Ship To Sub Region'].nunique():,}")
    m3.metric("Trend reversals",
              f"{(summary['Primary code'] == 'TREND_REVERSAL').sum():,}")
    m4.metric("Seasonality losses",
              f"{(summary['Primary code'] == 'SEASONALITY_LOSS').sum():,}")

    # ---- Reason-code breakdown ---------------------------------------------
    st.markdown("### 📊 Reason-code breakdown")
    breakdown = (summary["Primary reason"].value_counts()
                 .rename_axis("Reason").reset_index(name="Keys"))
    bcol, lcol = st.columns([1.3, 1])
    with bcol:
        fig = go.Figure(go.Bar(
            x=breakdown["Keys"], y=breakdown["Reason"], orientation="h",
            marker=dict(color="#4da3ff"),
            hovertemplate="%{y}<br>%{x} Key(s)<extra></extra>",
        ))
        dark_layout(
            fig, title="Primary reason code — Key count",
            xaxis_title="Number of Keys", yaxis_title="",
            height=260, margin=dict(t=50, l=10, r=20, b=40),
            legend=dict(orientation="h", y=-0.3, x=0.5, xanchor="center"),
        )
        fig.update_yaxes(tickformat="")
        st.plotly_chart(fig, use_container_width=True,
                        config={"displaylogo": False})
    with lcol:
        st.markdown("**Severity by Sub Region (max)**")
        region_sev = (summary.groupby("Ship To Sub Region")["Severity"].max()
                      .sort_values(ascending=False).reset_index())
        st.dataframe(
            region_sev.style.format({"Severity": "{:.1f}"}),
            use_container_width=True, hide_index=True,
        )

    # ---- Per-region table ---------------------------------------------------
    st.markdown(f"### 📋 Top {top_n} anomalies per Ship To Sub Region")
    regions = sorted(summary["Ship To Sub Region"].unique().tolist())
    region_pick = st.selectbox(
        "Ship To Sub Region", ["(All regions)"] + regions, index=0,
        key="anom::region_pick",
    )
    table = summary if region_pick == "(All regions)" \
        else summary[summary["Ship To Sub Region"] == region_pick]

    display_cols = ["Ship To Sub Region", "Key", "Primary reason",
                    "All reasons", "Severity", "Detail", "IQR outliers",
                    "Hist %/yr", "Fcst %/yr", "Seasonal corr"]
    st.dataframe(
        table[display_cols].style.format({
            "Severity": "{:.1f}",
            "Hist %/yr": lambda v: "–" if pd.isna(v) else f"{v:+.1f}%",
            "Fcst %/yr": lambda v: "–" if pd.isna(v) else f"{v:+.1f}%",
            "Seasonal corr": lambda v: "–" if pd.isna(v) else f"{v:+.2f}",
        }, na_rep="–"),
        use_container_width=True, hide_index=True, height=380,
    )

    csv = table[display_cols].to_csv(index=False)
    st.download_button(
        "⬇️ Download anomaly summary (CSV)", data=csv,
        file_name="anomaly_summary.csv", mime="text/csv",
        key="anom::dl",
    )

    # ---- Drill-down ---------------------------------------------------------
    st.markdown("### 🔬 Drill-down to root cause")
    st.caption("Pick a Key for root cause and next steps.")
    key_choices = table.sort_values("Severity", ascending=False)["Key"].tolist()
    if not key_choices:
        return
    chosen = st.selectbox("Key", key_choices, index=0, key="anom::drill_key")

    _render_anomaly_drilldown(chosen, filtered)

    with st.expander("ℹ️ How are anomalies detected and scored?"):
        _render_reason_code_legend()


def _render_reason_code_legend() -> None:
    st.markdown(
        "Each Key's Sales History and Statistical Forecast are scored; "
        "severity ranks the worst first. A Key may carry several codes:"
    )
    for code, info in REASON_CODES.items():
        st.markdown(f"- **{info['label']}** (`{code}`): {info['short']}")


def _render_anomaly_drilldown(key: str, filtered_df: pd.DataFrame) -> None:
    boundary = history_forecast_boundary(filtered_df)
    grp = filtered_df[filtered_df["Key"] == key]
    profile = analyze_key_anomalies(grp, boundary)
    if profile is None:
        st.info("No anomaly details available for this Key.")
        return

    # Severity + codes header
    chips = " ".join(
        f"<span style='background:#2d4b7a;color:#fff;padding:2px 8px;"
        f"border-radius:10px;margin-right:6px;font-size:0.82em;'>"
        f"{REASON_CODES[c]['label']}</span>"
        for c in profile["all_codes"]
    )
    st.markdown(
        f"**Key:** `{key}`  •  **Severity:** {profile['severity']:.1f}<br>{chips}",
        unsafe_allow_html=True,
    )

    # Mini chart: history (+cleansed) & forecast with the boundary
    sales_grp = profile["sales_grp"]
    fcst_grp = profile["fcst_grp"]
    iqr_res = profile["metrics"].get("iqr_result")

    fig = go.Figure()
    if boundary is not None:
        fig.add_vline(x=boundary, line=dict(color=DARK_MUTED, width=1, dash="dot"))
    fig.add_trace(go.Scatter(
        x=sales_grp["Date"], y=sales_grp["Value"], mode="lines+markers",
        name="Sales History",
        line=dict(color=SERIES_STYLE[SALES_HISTORY_LABEL]["color"], width=2.4),
        marker=dict(size=5),
        hovertemplate="Sales History<br>%{x|%b %Y}<br>%{y:,.0f}<extra></extra>",
    ))
    if not fcst_grp.empty:
        fig.add_trace(go.Scatter(
            x=fcst_grp["Date"], y=fcst_grp["Value"], mode="lines+markers",
            name="Statistical Forecast",
            line=dict(color=SERIES_STYLE[STAT_FORECAST_LABEL]["color"], width=2.4),
            marker=dict(size=5, symbol="diamond"),
            hovertemplate="Forecast<br>%{x|%b %Y}<br>%{y:,.0f}<extra></extra>",
        ))
    # Trend overlays
    ht = profile["metrics"].get("hist_trend")
    ft = profile["metrics"].get("fcst_trend")
    if ht is not None:
        fig.add_trace(go.Scatter(
            x=ht["dates"], y=ht["fit_line"], mode="lines", name="History trend",
            line=dict(color=HISTORY_TREND_COLOR, width=2, dash="dash"),
        ))
    if ft is not None:
        fig.add_trace(go.Scatter(
            x=ft["dates"], y=ft["fit_line"], mode="lines", name="Forecast trend",
            line=dict(color=FORECAST_TREND_COLOR, width=2, dash="dash"),
        ))
    # Outlier markers
    if iqr_res is not None:
        od = iqr_res[iqr_res["IsOutlier"]]
        if not od.empty:
            fig.add_trace(go.Scatter(
                x=od["Date"], y=od["Value"], mode="markers", name="IQR outlier",
                marker=dict(color="#ff6b6b", size=12, symbol="x-thin",
                            line=dict(color="#ff6b6b", width=3)),
                hovertemplate="Outlier<br>%{x|%b %Y}<br>%{y:,.0f}<extra></extra>",
            ))
    dark_layout(
        fig, title=f"{key} — history, forecast & trends",
        xaxis_title="Months", yaxis_title="Demand volume in kgs",
        height=420,
        legend=dict(orientation="h", y=-0.3, x=0.5, xanchor="center",
                    bgcolor="rgba(0,0,0,0)", font=dict(color=DARK_TEXT)),
        margin=dict(t=60, l=60, r=30, b=110),
    )
    st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False})

    # Supporting metrics
    mc1, mc2, mc3 = st.columns(3)
    h_pct = profile["metrics"].get("hist_pct_per_year")
    f_pct = profile["metrics"].get("fcst_pct_per_year")
    corr = profile["metrics"].get("seasonal_corr")
    mc1.metric("History trend", "–" if h_pct is None else f"{h_pct:+.1f}%/yr")
    mc2.metric("Forecast trend", "–" if f_pct is None else f"{f_pct:+.1f}%/yr")
    mc3.metric("Seasonal corr", "–" if corr is None else f"{corr:+.2f}")

    # Root cause + next steps for each applicable reason code
    st.markdown("#### 🩺 Root cause & recommended next steps")
    for r in profile["reasons"]:
        code = r["code"]
        info = REASON_CODES[code]
        with st.container(border=True):
            st.markdown(
                f"**{info['label']}**  ·  severity {r['severity']:.0f}  ·  "
                f"_{r['detail']}_"
            )
            st.markdown(f"**Root cause:** {info['root_cause']}")
            st.markdown(f"**Next steps:** {info['next_steps']}")


# ---------------------------------------------------------------------------
# UI – Tab: STF Variation & Exceptions
# ---------------------------------------------------------------------------
def render_stf_variation_tab(long_df: pd.DataFrame,
                             all_months: List[pd.Timestamp]) -> None:
    st.subheader("📊 STF Variation & Exceptions")
    st.caption(
        "Compares the current **Statistical Forecast Committed** against "
        "**Lag 1**. Variance = (Current − Lag 1) / Lag 1, per Key, over a "
        "horizon that starts at **Month 4 (M4)** — the 4th month after the "
        "last active Sales-History month."
    )

    # ---- Filters (same as Forecast vs Actuals, minus Data; NO status default)
    with st.container(border=True):
        st.markdown("**🔎 Filters** — cascade like Excel slicers (each narrows "
                    "the others). Leave a filter empty to include everything.")
        stf_filter_cols = [c for c in FILTER_COLUMNS if c != "Data"]
        selections = render_filter_strip(
            long_df, stf_filter_cols, key_prefix="stf")

    filtered = apply_filters(long_df, selections)
    if filtered.empty:
        st.warning("No data matches the current filter combination.")
        return

    # Need both current STF and Lag 1 to compute variance.
    if STF_LAG1_LABEL not in long_df["Data"].unique():
        st.info("This file has no **Statistical Forecast Committed Lag 1 (kg)** "
                "line, so STF variation cannot be computed. Please upload the "
                "raw Arkieva export that includes the Lag 1 series.")
        return

    m4 = stf_m4_month(filtered)
    if m4 is None:
        st.info("Couldn't determine Month 4 (no Sales History for the current "
                "filter selection).")
        return

    # ---- Horizon + sliding window controls ---------------------------------
    c1, c2, c3 = st.columns([1.2, 1.4, 1.2])
    with c1:
        horizon_choice = st.radio(
            "Horizon", ["M4 → fiscal year-end", "M4 → next 12 months"],
            key="stf::horizon",
            help="Fiscal year is Jan–Dec. Horizon 1 runs from M4 to December "
                 "of M4's year; Horizon 2 is M4 plus the next 11 months.",
        )
    horizon = "fiscal_year" if horizon_choice.startswith("M4 → fiscal") else "next_12"
    full_horizon = stf_horizon_months(all_months, m4, horizon=horizon)
    max_window = len(full_horizon)

    with c2:
        if max_window >= 2:
            window = st.slider(
                "Months from M4 (sliding window)",
                min_value=1, max_value=max_window, value=max_window, step=1,
                key="stf::window",
                help="Shorten the horizon by sliding in from the far end. "
                     "M4 is always the start.",
            )
        else:
            window = max_window
            st.caption("Horizon has a single month.")
    with c3:
        threshold_pct = st.slider(
            "Absolute STF variation threshold", min_value=1, max_value=100,
            value=5, step=1, format="%d%%", key="stf::threshold",
            help="Keys with Absolute STF Variance above this are flagged as "
                 "exceptions. Default 5%.",
        )
    threshold = threshold_pct / 100.0

    horizon_months = stf_horizon_months(all_months, m4, horizon=horizon, window=window)
    if not horizon_months:
        st.info("No forecast months fall in the selected horizon.")
        return

    st.caption(
        f"**M4 = {m4.strftime('%b %Y')}** · horizon "
        f"**{horizon_months[0].strftime('%b %Y')} – "
        f"{horizon_months[-1].strftime('%b %Y')}** "
        f"({len(horizon_months)} month(s)) · threshold **{threshold_pct}%**"
    )

    # ---- Per-Key variance ---------------------------------------------------
    var_df = compute_stf_variance(filtered, tuple(horizon_months))
    if var_df.empty:
        st.info("No Keys have a non-zero Lag 1 forecast in this horizon, so "
                "variance can't be computed.")
        return

    n_exceptions = int((var_df["Absolute STF Variance"] > threshold).sum())
    net_cur = var_df["CurrentSTF"].sum()
    net_lag = var_df["Lag1STF"].sum()
    rollup_var = ((net_cur - net_lag) / net_lag) if net_lag else np.nan

    m1, m2, m3, m4c = st.columns(4)
    m1.metric("Keys analysed", f"{len(var_df):,}")
    m2.metric(f"Exceptions (> {threshold_pct}%)", f"{n_exceptions:,}")
    m3.metric("Roll-up variance",
              "–" if pd.isna(rollup_var) else f"{rollup_var*100:+.1f}%")
    m4c.metric("Net STF change (kg)", f"{net_cur - net_lag:,.0f}")

    # ---- History (36m back) vs forecast vintages (24m forward) -------------
    st.markdown("### 📉 History & forecast vintages")
    st.caption(
        "Aggregated over the current filter selection: the **last 36 months** "
        "of **Sales History** and **History For Forecast** (up to the last "
        "active history month), and the **next 24 months** of the **current "
        "Statistical Forecast Committed** and **Lag 1**. Where the two "
        "committed lines separate is where this cycle's forecast changed "
        "from the prior cycle."
    )
    boundary = history_forecast_boundary(filtered)
    if boundary is None:
        st.info("No Sales History for this selection, so the comparison "
                "window can't be anchored.")
    else:
        hist_start = (pd.Timestamp(boundary) - pd.DateOffset(months=35)).replace(day=1)
        fcst_start = (pd.Timestamp(boundary) + pd.DateOffset(months=1)).replace(day=1)
        fcst_end = (pd.Timestamp(boundary) + pd.DateOffset(months=24)).replace(day=1)

        # (label, display name, colour, window start, window end)
        cmp_series = [
            (SALES_HISTORY_LABEL, "Sales History (last 36m)",
             SERIES_STYLE[SALES_HISTORY_LABEL]["color"], hist_start, boundary),
            (HISTORY_FOR_FORECAST_LABEL, "History For Forecast (last 36m)",
             SERIES_STYLE[HISTORY_FOR_FORECAST_LABEL]["color"], hist_start, boundary),
            (STF_CURRENT_LABEL, "Statistical Forecast Committed (next 24m)",
             SERIES_STYLE[STAT_FORECAST_LABEL]["color"], fcst_start, fcst_end),
            (STF_LAG1_LABEL, "Statistical Forecast Committed Lag 1 (next 24m)",
             "#c084fc", fcst_start, fcst_end),
        ]
        cmp_agg = (filtered[filtered["Data"].isin([s[0] for s in cmp_series])]
                   .dropna(subset=["Value"])
                   .groupby(["Date", "Data"], as_index=False)["Value"].sum())
        fig_cmp = go.Figure()
        fig_cmp.add_vline(x=boundary,
                          line=dict(color=DARK_MUTED, width=1, dash="dot"))
        n_traces = 0
        for label, name, color, w_start, w_end in cmp_series:
            s = cmp_agg[(cmp_agg["Data"] == label)
                        & (cmp_agg["Date"] >= w_start)
                        & (cmp_agg["Date"] <= w_end)]
            if s.empty:
                continue
            fig_cmp.add_trace(go.Scatter(
                x=s["Date"], y=s["Value"], name=name, mode="lines+markers",
                line=dict(color=color, width=2.2, shape="spline"),
                marker=dict(size=4),
                hovertemplate=f"{name}<br>%{{x|%b %Y}}<br>%{{y:,.0f}} kg"
                              "<extra></extra>",
            ))
            n_traces += 1
        if n_traces == 0:
            st.info("No data in the comparison windows for this selection.")
        else:
            dark_layout(
                fig_cmp, title="Sales History (36m) vs forecast vintages (24m)",
                xaxis_title="Months", yaxis_title="Demand volume in kgs",
                height=430,
                legend=dict(orientation="h", y=-0.28, x=0.5, xanchor="center",
                            bgcolor="rgba(0,0,0,0)", font=dict(color=DARK_TEXT)),
                margin=dict(t=60, l=60, r=30, b=110),
            )
            st.plotly_chart(fig_cmp, use_container_width=True,
                            config={"displaylogo": False})

    # ---- Top 25 exception Keys ---------------------------------------------
    st.markdown("### 🎯 Top 25 Keys by Absolute STF Variance")
    top25 = var_df.head(25)
    if n_exceptions == 0:
        st.success(f"No Keys exceed the {threshold_pct}% threshold in this "
                   "horizon — the table below shows the largest movers. 🎉")
    else:
        st.caption(
            f"**{n_exceptions}** Key(s) exceed the {threshold_pct}% threshold. "
            "Cells above the threshold are highlighted; the table is sorted "
            "worst-first, so the top rows are the Keys to review."
        )
    show = top25.copy()
    show["STF Variance"] = show["STF Variance"] * 100
    show["Absolute STF Variance"] = show["Absolute STF Variance"] * 100
    disp_cols = ["Key", "Business Line", "Ship To Sub Region",
                 "Arkieva Active Status", "CurrentSTF", "Lag1STF",
                 "STF Variance", "Absolute STF Variance"]
    disp_cols = [c for c in disp_cols if c in show.columns]

    # Highlight Absolute STF Variance cells above the threshold in light
    # red/pink (pure Python — no matplotlib dependency on Streamlit Cloud).
    thr_display = threshold * 100.0

    def _thr_highlight(col: pd.Series) -> list:
        return ["background-color: #ff8fa3; color: #0e1117; font-weight: 600;"
                if float(v) > thr_display else "" for v in col]

    st.dataframe(
        show[disp_cols].style.format({
            "CurrentSTF": "{:,.0f}", "Lag1STF": "{:,.0f}",
            "STF Variance": "{:+.1f}%", "Absolute STF Variance": "{:.1f}%",
        }).apply(_thr_highlight, subset=["Absolute STF Variance"]),
        use_container_width=True, hide_index=True, height=380,
    )
    csv = top25.to_csv(index=False)
    st.download_button("⬇️ Download top variation Keys (CSV)", data=csv,
                       file_name="stf_exception_keys.csv", mime="text/csv",
                       key="stf::dl_keys")

    # ---- Month highlight for the top Keys -----------------------------------
    exceptions = var_df[var_df["Absolute STF Variance"] > threshold].head(25)
    month_pool = exceptions if not exceptions.empty else top25
    if not month_pool.empty:
        st.markdown("### 🗓️ Which months drive the variation?")
        st.caption("Month-level STF Variance. Red = larger swing; the "
                   "strongest cells are the months to investigate.")
        month_options = ["(Top 10 exception Keys)"] + month_pool["Key"].tolist()
        month_scope = st.selectbox("Month-variation scope", month_options,
                                   index=0, key="stf::month_key")
        if month_scope.startswith("(Top 10"):
            mv_keys = tuple(month_pool["Key"].head(10).tolist())
        else:
            mv_keys = (month_scope,)
        mv = compute_stf_month_variance(filtered, mv_keys, tuple(horizon_months))
        if not mv.empty:
            heat = mv.pivot_table(index="Key", columns="Date",
                                  values="STF Variance", aggfunc="first")
            heat = heat.reindex(index=list(mv_keys))
            heat.columns = [pd.Timestamp(c).strftime("%b %Y") for c in heat.columns]
            z = (heat.values * 100)
            fig_hm = go.Figure(go.Heatmap(
                z=z, x=list(heat.columns), y=list(heat.index),
                colorscale="RdBu", zmid=0,
                hovertemplate="%{y}<br>%{x}<br>Variance: %{z:.1f}%<extra></extra>",
                colorbar=dict(title="%"),
            ))
            dark_layout(
                fig_hm, title="", xaxis_title="Month", yaxis_title="Key",
                height=max(240, 26 * len(mv_keys) + 140),
                margin=dict(t=20, l=200, r=30, b=60),
            )
            fig_hm.update_yaxes(autorange="reversed")
            st.plotly_chart(fig_hm, use_container_width=True,
                            config={"displaylogo": False})

    # ---- Waterfall: variation drivers by Arkieva Active Status -------------
    st.markdown("### 💧 STF variation drivers (waterfall)")
    st.caption(
        "Decomposes the net STF change (Current − Lag 1) into **Arkieva "
        "Active Status** buckets — Active, Active New, New, Sparse, Obsolete, "
        "Inactive, … — so you can see where the forecast changes are coming "
        "from (e.g. volume added by new items vs lost from obsolete ones). "
        "Rolls up across the current filter selection; pick a single Key to "
        "drill down."
    )
    drill_options = ["(All Keys — roll-up)"] + var_df["Key"].tolist()
    drill_key = st.selectbox("Driver scope", drill_options, index=0,
                             key="stf::driver_key")
    keys_arg = None if drill_key.startswith("(All") else (drill_key,)
    drivers = compute_stf_drivers(filtered, tuple(horizon_months), keys=keys_arg)

    if drivers.empty or drivers["Delta"].abs().sum() == 0:
        st.info("No driver decomposition available for this selection.")
    else:
        lag_base = net_lag if keys_arg is None else \
            var_df.loc[var_df["Key"] == drill_key, "Lag1STF"].sum()
        measures = ["absolute"] + ["relative"] * len(drivers) + ["total"]
        x_labels = ["Lag 1 base"] + drivers["Driver"].tolist() + ["Current"]
        y_vals = [lag_base] + drivers["Delta"].tolist() + [None]
        fig_wf = go.Figure(go.Waterfall(
            orientation="v", measure=measures, x=x_labels, y=y_vals,
            connector=dict(line=dict(color=DARK_MUTED)),
            increasing=dict(marker=dict(color="#51cf66")),
            decreasing=dict(marker=dict(color="#ff6b6b")),
            totals=dict(marker=dict(color=ACCENT)),
            hovertemplate="%{x}<br>%{y:,.0f} kg<extra></extra>",
        ))
        dark_layout(
            fig_wf,
            title=f"STF change decomposition — {'roll-up' if keys_arg is None else drill_key}",
            xaxis_title="", yaxis_title="Forecast volume (kg)",
            height=440, margin=dict(t=60, l=60, r=30, b=110),
            showlegend=False,
        )
        st.plotly_chart(fig_wf, use_container_width=True,
                        config={"displaylogo": False})

        dr_show = drivers.copy()
        total_abs = dr_show["Delta"].abs().sum()
        dr_show["% of gross change"] = np.where(
            total_abs != 0, dr_show["Delta"].abs() / total_abs * 100, 0.0)
        st.dataframe(
            dr_show.style.format({"Delta": "{:,.0f}",
                                  "% of gross change": "{:.1f}%"}),
            use_container_width=True, hide_index=True,
        )


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
        st.markdown("**🔎 Filters** — cascade like Excel slicers (each narrows the others). "
                    "*Arkieva Active Status* defaults to **Active + Sparse**; clear or change any filter as needed.")
        selections = render_filter_strip(
            long_df, FILTER_COLUMNS, key_prefix="dash",
            default_selections={ACTIVE_STATUS_COL: DEFAULT_ACTIVE_STATUSES})

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
        st.markdown("**🔎 Filters** — cascade like Excel slicers (each narrows the others). "
                    "*Arkieva Active Status* defaults to **Active + Sparse**; clear or change any filter as needed.")
        selections = render_filter_strip(
            sales, outlier_filter_cols, key_prefix="outlier",
            default_selections={ACTIVE_STATUS_COL: DEFAULT_ACTIVE_STATUSES},
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
        st.markdown("**🔎 Filters** — cascade like Excel slicers (each narrows the others). "
                    "*Arkieva Active Status* defaults to **Active + Sparse**; clear or change any filter as needed.")
        selections = render_filter_strip(
            long_df, season_filter_cols, key_prefix="season",
            default_selections={ACTIVE_STATUS_COL: DEFAULT_ACTIVE_STATUSES},
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
# UI – Tab: Statistical Forecast Adoption %
# ---------------------------------------------------------------------------
def render_stat_adoption_tab(long_df: pd.DataFrame) -> None:
    st.subheader("📦 Statistical Forecast Adoption %")
    st.caption(
        "Share of Statistical-Forecast volume that is **on** the statistical "
        "forecast — i.e. the granular Business Line + Material + Ship To Sub "
        "Region + Active Status combinations with **Arkieva Review Req = No** "
        "and **Active/Sparse** status — measured against each Ship To Sub "
        "Region's total forecast, by fiscal year (Jan–Dec)."
    )

    # ---- Filters (same set + Active/Sparse default) ------------------------
    with st.container(border=True):
        st.markdown("**🔎 Filters** — cascade like Excel slicers (each narrows the others). "
                    "*Arkieva Active Status* defaults to **Active + Sparse**; clear or change any filter as needed.")
        # The Data filter is dropped (the tab is scoped to the Statistical
        # Forecast internally) and so is Arkieva Review Req — the on-stat
        # definition already applies Review Req = No, so the filter is
        # redundant here.
        adopt_filter_cols = [c for c in FILTER_COLUMNS
                             if c not in ("Data", REVIEW_REQ_COL)]
        selections = render_filter_strip(
            long_df, adopt_filter_cols, key_prefix="adopt",
            default_selections={ACTIVE_STATUS_COL: DEFAULT_ACTIVE_STATUSES})

    filtered = apply_filters(long_df, selections)
    if filtered.empty:
        st.warning("No data matches the current filter combination.")
        return

    stat_rows = filtered[filtered["Data"] == STAT_FORECAST_LABEL]
    if stat_rows.dropna(subset=["Value"]).empty:
        st.info("No Statistical Forecast volume for the current filter selection.")
        return

    active_year = active_year_from_history(filtered)
    if active_year is None:
        st.info("Couldn't determine the active year (no Sales History found "
                "for the current filter selection).")
        return

    # ---- Overall metric (active year onwards) ------------------------------
    overall = compute_stat_adoption(filtered, group_col=None, active_year=active_year)
    if overall.empty:
        st.info("No Statistical Forecast volume from the active year onwards.")
        return
    grand_on = overall["OnStatVolume"].sum()
    grand_tot = overall["TotalVolume"].sum()
    grand_pct = (grand_on / grand_tot * 100) if grand_tot else 0.0
    years_covered = sorted(overall["Year"].astype(int).unique())

    st.caption(
        f"**Active year: {active_year}** (last active month in Sales History). "
        f"Adoption is computed for **{active_year}–{years_covered[-1]}** "
        "(active year onwards); earlier years are excluded."
    )

    m1, m2, m3 = st.columns(3)
    m1.metric(f"Overall adoption ({active_year}+)", f"{grand_pct:.1f}%")
    m2.metric("On-stat forecast (kg)", f"{grand_on:,.0f}")
    m3.metric("Total forecast (kg)", f"{grand_tot:,.0f}")

    # ---- Adoption % by fiscal year (line + bar combo) ----------------------
    st.markdown("### 📈 Adoption % by fiscal year")
    year_df = overall.copy()
    if not year_df.empty:
        year_df["YearLabel"] = year_df["Year"].astype(int).astype(str)
        fig_year = go.Figure()
        fig_year.add_trace(go.Bar(
            x=year_df["YearLabel"], y=year_df["AdoptionPct"],
            name="Adoption %", marker=dict(color="#4da3ff"),
            text=[f"{v:.1f}%" for v in year_df["AdoptionPct"]],
            textposition="outside",
            hovertemplate="FY %{x}<br>Adoption: %{y:.1f}%<extra></extra>",
        ))
        fig_year.add_trace(go.Scatter(
            x=year_df["YearLabel"], y=year_df["AdoptionPct"],
            name="Trend", mode="lines+markers",
            line=dict(color="#ffa94d", width=2.5),
            marker=dict(size=7),
            hovertemplate="FY %{x}<br>%{y:.1f}%<extra></extra>",
        ))
        dark_layout(
            fig_year, title="Statistical Forecast adoption by fiscal year",
            xaxis_title="Fiscal year (Jan–Dec)", yaxis_title="Adoption %",
            height=380,
            legend=dict(orientation="h", y=-0.25, x=0.5, xanchor="center",
                        bgcolor="rgba(0,0,0,0)", font=dict(color=DARK_TEXT)),
            margin=dict(t=60, l=60, r=30, b=80),
        )
        fig_year.update_yaxes(ticksuffix="%", rangemode="tozero")
        st.plotly_chart(fig_year, use_container_width=True,
                        config={"displaylogo": False})

    # ---- By Business Line and by Ship To Sub Region (grouped bars) ----------
    bcol, rcol = st.columns(2)

    with bcol:
        st.markdown("### 🏭 By Business Line")
        bl_df = compute_stat_adoption(filtered, group_col="Business Line",
                                      active_year=active_year)
        if bl_df.empty:
            st.info("No data.")
        else:
            bl_df["YearLabel"] = bl_df["Year"].astype(int).astype(str)
            fig_bl = go.Figure()
            for bl in sorted(bl_df["Business Line"].unique()):
                sub = bl_df[bl_df["Business Line"] == bl]
                fig_bl.add_trace(go.Bar(
                    x=sub["YearLabel"], y=sub["AdoptionPct"], name=str(bl),
                    hovertemplate=f"<b>{bl}</b><br>FY %{{x}}<br>"
                                  "%{y:.1f}%<extra></extra>",
                ))
            dark_layout(
                fig_bl, title="Adoption % by Business Line",
                xaxis_title="Fiscal year", yaxis_title="Adoption %",
                height=400, barmode="group",
                legend=dict(orientation="h", y=-0.35, x=0.5, xanchor="center",
                            bgcolor="rgba(0,0,0,0)", font=dict(color=DARK_TEXT)),
                margin=dict(t=60, l=50, r=20, b=110),
            )
            fig_bl.update_yaxes(ticksuffix="%", rangemode="tozero")
            st.plotly_chart(fig_bl, use_container_width=True,
                            config={"displaylogo": False})

    with rcol:
        st.markdown("### 🌍 By Ship To Sub Region")
        rg_df = compute_stat_adoption(filtered, group_col="Ship To Sub Region",
                                      active_year=active_year)
        if rg_df.empty:
            st.info("No data.")
        else:
            rg_df["YearLabel"] = rg_df["Year"].astype(int).astype(str)
            fig_rg = go.Figure()
            for rg in sorted(rg_df["Ship To Sub Region"].unique()):
                sub = rg_df[rg_df["Ship To Sub Region"] == rg]
                fig_rg.add_trace(go.Bar(
                    x=sub["YearLabel"], y=sub["AdoptionPct"], name=str(rg),
                    hovertemplate=f"<b>{rg}</b><br>FY %{{x}}<br>"
                                  "%{y:.1f}%<extra></extra>",
                ))
            dark_layout(
                fig_rg, title="Adoption % by Ship To Sub Region",
                xaxis_title="Fiscal year", yaxis_title="Adoption %",
                height=400, barmode="group",
                legend=dict(orientation="h", y=-0.35, x=0.5, xanchor="center",
                            bgcolor="rgba(0,0,0,0)", font=dict(color=DARK_TEXT)),
                margin=dict(t=60, l=50, r=20, b=110),
            )
            fig_rg.update_yaxes(ticksuffix="%", rangemode="tozero")
            st.plotly_chart(fig_rg, use_container_width=True,
                            config={"displaylogo": False})

    # ---- Heatmap: Business Line × Year ------------------------------------
    bl_df_h = compute_stat_adoption(filtered, group_col="Business Line",
                                    active_year=active_year)
    if not bl_df_h.empty and bl_df_h["Business Line"].nunique() > 1:
        st.markdown("### 🔥 Adoption % heatmap — Business Line × fiscal year")
        pivot = bl_df_h.pivot(index="Business Line", columns="Year",
                              values="AdoptionPct").sort_index()
        fig_hm = go.Figure(go.Heatmap(
            z=pivot.values,
            x=[str(int(c)) for c in pivot.columns],
            y=pivot.index.tolist(),
            colorscale="Blues", zmin=0,
            text=[[f"{v:.0f}%" if pd.notna(v) else "" for v in row]
                  for row in pivot.values],
            texttemplate="%{text}", textfont=dict(color="#0e1117"),
            hovertemplate="%{y}<br>FY %{x}<br>%{z:.1f}%<extra></extra>",
            colorbar=dict(title="%"),
        ))
        dark_layout(
            fig_hm, title="", xaxis_title="Fiscal year",
            yaxis_title="Business Line", height=300,
            margin=dict(t=20, l=120, r=30, b=50),
        )
        st.plotly_chart(fig_hm, use_container_width=True,
                        config={"displaylogo": False})

    # ---- Materials on Statistical Forecast (interactive table) -------------
    st.markdown("### 📋 Materials on Statistical Forecast")
    st.caption(
        "Granular Business Line + Material + Ship To Sub Region + Active "
        "Status combinations on the Statistical Forecast (Review Req = No and "
        "Active/Sparse), with forecast volume per fiscal year (active year "
        "onwards) and each combination's share of its Ship To Sub Region's "
        "total forecast. Reacts to the filters above."
    )
    mat_table = stat_adoption_material_table(filtered, active_year=active_year)
    if mat_table.empty:
        st.info("No material combinations are on the Statistical Forecast for "
                "the current filter selection.")
    else:
        st.caption(f"**{len(mat_table):,}** combination(s) on Statistical "
                   f"Forecast (active year {active_year} onwards).")
        fmt_cols = [c for c in mat_table.columns
                    if c.startswith("FY ") or c in
                    ("Total forecast (kg)", "Region total forecast (kg)")]
        styler = mat_table.style.format(
            {c: "{:,.0f}" for c in fmt_cols}, na_rep="–")
        if "% of region forecast" in mat_table.columns:
            styler = styler.format({"% of region forecast": "{:.1f}%"})
        st.dataframe(styler, use_container_width=True, hide_index=True, height=420)
        csv = mat_table.to_csv(index=False)
        st.download_button(
            "⬇️ Download materials on Statistical Forecast (CSV)",
            data=csv, file_name="materials_on_stat_forecast.csv",
            mime="text/csv", key="adopt::dl",
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
            "1. An **Anomaly Summary** that ranks the top anomalies per Ship "
            "To Sub Region, assigns each a reason code (trend reversal, trend "
            "mismatch, seasonality loss, IQR outliers), and lets you drill "
            "down to the root cause and recommended next steps.\n"
            "2. An **STF Variation & Exceptions** view comparing the current "
            "Statistical Forecast Committed against Lag 1 over M4-based "
            "horizons, flagging top exception Keys, the months driving them, "
            "and a waterfall of variation drivers.\n"
            "3. An interactive replica of the *Forecast vs Actuals* pivot "
            "chart with all original slicer filters on top — clearly "
            "splitting **History** vs **Forecast** series, with optional "
            "**trend lines** and an automatic trend interpretation.\n"
            "4. An **outlier detection & correction** view (choose **IQR** or "
            "**Sigma**) for the **Sales History** of every Key, producing a "
            "cleansed history alongside the original.\n"
            "5. A **year-over-year seasonality** view comparing the forecast's "
            "seasonal shape against the last 3 historical years, with an "
            "automatic interpretation of forecast quality.\n"
            "6. A **Statistical Forecast Adoption %** view showing the share "
            "of volume on statistical forecast by fiscal year, Business Line "
            "and Ship To Sub Region, with the underlying materials table."
        )
        st.stop()

    file_id = f"{uploaded.name}::{uploaded.size}"
    if st.session_state["current_file_id"] != file_id:
        reset_app_state()
        st.session_state["current_file_id"] = file_id

    file_bytes = uploaded.getvalue()
    try:
        long_df, all_months = load_excel(file_bytes, uploaded.name)
    except ValueError as exc:
        st.error(f"❌ {exc}")
        st.stop()
    except Exception as exc:  # noqa: BLE001
        st.error(f"❌ Unexpected error while loading the file: {exc}")
        st.stop()

    if long_df.empty:
        st.error("The uploaded file contains no usable rows.")
        st.stop()

    tab0, tab_stf, tab1, tab2, tab3, tab4 = st.tabs([
        "🧭 Anomaly Summary",
        "📊 STF Variation & Exceptions",
        "📈 Forecast vs Actuals",
        "🚨 Outlier Detection & Correction",
        "📅 Seasonality (YoY)",
        "📦 Statistical Forecast Adoption %",
    ])
    with tab0:
        render_anomaly_tab(long_df)
    with tab_stf:
        render_stf_variation_tab(long_df, all_months)
    with tab1:
        render_dashboard_tab(long_df, uploaded.name)
    with tab2:
        render_outlier_tab(long_df)
    with tab3:
        render_seasonality_tab(long_df)
    with tab4:
        render_stat_adoption_tab(long_df)


if __name__ == "__main__":
    main()
