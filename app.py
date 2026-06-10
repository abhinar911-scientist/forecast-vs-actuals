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

# Visual style: cool/solid for History, warm/dashed for Forecast
SERIES_STYLE = {
    "Sales History (kg)":           {"color": "#1f77b4", "dash": "solid", "category": "History"},
    "History For Forecast (kg)":    {"color": "#17a2b8", "dash": "solid", "category": "History"},
    "Statistical Forecast (kg)":    {"color": "#ff7f0e", "dash": "dash",  "category": "Forecast"},
    "Final Demand Plan Lag 1 (kg)": {"color": "#d62728", "dash": "dot",   "category": "Forecast"},
}

# Light tints for the History / Forecast background bands
HISTORY_BAND_COLOR = "rgba(31, 119, 180, 0.07)"
FORECAST_BAND_COLOR = "rgba(255, 127, 14, 0.07)"


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
        "<div style='background:rgba(31,119,180,0.10); padding:10px 14px;"
        " border-left:4px solid #1f77b4; border-radius:4px;'>"
        "<b>📘 History</b><br>"
        "<span style='font-size:0.85em; color:#444;'>"
        "Solid lines · Sales History, History For Forecast"
        "</span></div>",
        unsafe_allow_html=True,
    )
    bcol2.markdown(
        "<div style='background:rgba(255,127,14,0.10); padding:10px 14px;"
        " border-left:4px solid #ff7f0e; border-radius:4px;'>"
        "<b>📕 Forecast</b><br>"
        "<span style='font-size:0.85em; color:#444;'>"
        "Dashed lines · Statistical Forecast, Final Demand Plan Lag 1"
        "</span></div>",
        unsafe_allow_html=True,
    )
    if boundary is not None:
        st.caption(
            f"History / Forecast boundary detected at "
            f"**{boundary.strftime('%b %Y')}** "
            "(last month with Sales History data)."
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
                annotation_font=dict(color="#1f77b4", size=12),
            )
        if x_max >= boundary:
            fig.add_vrect(
                x0=boundary, x1=x_max,
                fillcolor=FORECAST_BAND_COLOR, line_width=0, layer="below",
                annotation_text="Forecast", annotation_position="top right",
                annotation_font=dict(color="#ff7f0e", size=12),
            )
        fig.add_vline(x=boundary, line=dict(color="#888", width=1, dash="dot"))

    # One trace per Data series, grouped in legend by category
    ordered = [s for s in HISTORY_SERIES + FORECAST_SERIES if s in pivot.columns]
    for col in ordered:
        style = SERIES_STYLE.get(
            col, {"color": "#888", "dash": "solid", "category": "Other"}
        )
        fig.add_trace(
            go.Scatter(
                x=pivot.index, y=pivot[col],
                name=col, mode="lines+markers",
                line=dict(color=style["color"], width=2.4, dash=style["dash"]),
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

    fig.update_layout(
        title="Forecast v/s Actuals",
        xaxis_title="Months", yaxis_title="Demand volume in kgs",
        hovermode="x unified",
        legend=dict(
            orientation="h",
            yanchor="top", y=-0.18,
            xanchor="center", x=0.5,
            groupclick="toggleitem",
        ),
        height=560, margin=dict(t=70, l=60, r=30, b=140),
        template="plotly_white",
    )
    fig.update_xaxes(tickformat="%b %Y", showgrid=True, gridcolor="#eee")
    fig.update_yaxes(showgrid=True, gridcolor="#eee", tickformat=",.0f")

    st.plotly_chart(fig, use_container_width=True)


# ---------------------------------------------------------------------------
# UI – Tab 2: Outlier Detection & Correction (Hampel filter)
# ---------------------------------------------------------------------------
def render_outlier_tab(long_df: pd.DataFrame) -> None:
    st.subheader("🚨 Outlier Detection & Correction — Hampel Filter")
    st.caption(
        "Outliers are detected and **corrected** on **Sales History (kg)** "
        "using the Hampel filter: a centred rolling window flags points more "
        "than *n·σ* (robust, MAD-based) from the local median, and replaces "
        "each flagged point with that median to produce the **Hampel filter "
        "cleansed history**. *History For Forecast (kg)* is shown unchanged "
        "for comparison."
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

    cfg1, cfg2 = st.columns(2)
    with cfg1:
        window_size = st.slider(
            "Window half-width", min_value=1, max_value=12, value=5, step=1,
            help="Half-width of the centred rolling window. The full window "
                 "spans 2×(this)+1 points. Larger windows smooth more.",
            key="outlier::window",
        )
    with cfg2:
        n_sigma = st.slider(
            "n_sigma (threshold)", min_value=1.0, max_value=5.0,
            value=3.0, step=0.5,
            help="Tolerance in robust (MAD-based) standard deviations. "
                 "Larger values flag fewer points. The Hampel default is 3.0.",
            key="outlier::nsigma",
        )

    @st.cache_data(show_spinner="Running Hampel filter across all Keys…")
    def compute_all(sales_in: pd.DataFrame, window_in: int, nsigma_in: float) -> pd.DataFrame:
        frames = []
        for key, grp in sales_in.groupby("Key"):
            res = hampel_filter(grp, window_size=window_in, n_sigma=nsigma_in)
            if res.empty:
                continue
            res.insert(0, "Key", key)
            frames.append(res)
        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

    all_results = compute_all(sales_f, window_size, n_sigma)
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
        _render_key_chart(chosen_key, key_df, hff_key, window_size, n_sigma)

        # Per-Key cleansed-history download
        cleansed = key_df[["Date", "Value", "Filtered", "IsOutlier"]].copy()
        cleansed = cleansed.rename(columns={
            "Value": "Sales History (kg)",
            "Filtered": "Hampel Cleansed History (kg)",
        })
        cleansed["Date"] = pd.to_datetime(cleansed["Date"]).dt.strftime("%Y-%m-%d")
        cbuf = io.StringIO()
        cleansed.to_csv(cbuf, index=False)
        st.download_button(
            "⬇️ Download this Key's cleansed history (CSV)",
            data=cbuf.getvalue(),
            file_name=f"hampel_cleansed_{chosen_key}.csv",
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
                "Corrected", "Median", "Lower", "Upper"]
        st.dataframe(
            outlier_df[cols].style.format(
                {c: "{:,.2f}" for c in
                 ["Original", "Corrected", "Median", "Lower", "Upper"]},
                na_rep="–",
            ),
            use_container_width=True,
        )
        buf = io.StringIO()
        outlier_df[cols].to_csv(buf, index=False)
        st.download_button(
            "⬇️ Download all corrected outliers (CSV)",
            data=buf.getvalue(),
            file_name="hampel_corrected_outliers.csv",
            mime="text/csv",
            key="outlier::dl_all",
        )

    # ---- Full cleansed-history export (all Keys) ----------------------------
    st.markdown("### 🧼 Cleansed history export")
    st.caption(
        "The full Hampel-cleansed Sales History for every Key in the current "
        "filter selection — outliers replaced by the local median, all other "
        "points unchanged."
    )
    full_clean = all_results[["Key", "Date", "Value", "Filtered", "IsOutlier"]].copy()
    full_clean = full_clean.rename(columns={
        "Value": "Sales History (kg)",
        "Filtered": "Hampel Cleansed History (kg)",
    })
    full_clean["Date"] = pd.to_datetime(full_clean["Date"]).dt.strftime("%Y-%m-%d")
    fbuf = io.StringIO()
    full_clean.to_csv(fbuf, index=False)
    st.download_button(
        "⬇️ Download full cleansed history — all Keys (CSV)",
        data=fbuf.getvalue(),
        file_name="hampel_cleansed_history_all_keys.csv",
        mime="text/csv",
        key="outlier::dl_full_clean",
    )


def _render_key_chart(
    key: str,
    key_df: pd.DataFrame,
    hff_df: pd.DataFrame,
    window_size: int,
    n_sigma: float,
) -> None:
    fig = go.Figure()

    # Threshold band (median ± threshold)
    bounds_df = key_df.dropna(subset=["Lower", "Upper"]).sort_values("Date")
    if not bounds_df.empty:
        fig.add_trace(go.Scatter(
            x=bounds_df["Date"], y=bounds_df["Upper"],
            mode="lines", name="Median + threshold",
            line=dict(color="rgba(150,150,150,0.5)", width=1, dash="dot"),
            hovertemplate="Upper: %{y:,.2f}<extra></extra>",
        ))
        fig.add_trace(go.Scatter(
            x=bounds_df["Date"], y=bounds_df["Lower"],
            mode="lines", name="Median − threshold",
            line=dict(color="rgba(150,150,150,0.5)", width=1, dash="dot"),
            fill="tonexty", fillcolor="rgba(150,150,150,0.12)",
            hovertemplate="Lower: %{y:,.2f}<extra></extra>",
        ))

    # Original Sales History
    fig.add_trace(go.Scatter(
        x=key_df["Date"], y=key_df["Value"],
        mode="lines+markers", name="Sales History (original)",
        line=dict(color=SERIES_STYLE[SALES_HISTORY_LABEL]["color"], width=2),
        marker=dict(size=6),
        hovertemplate="<b>Sales History</b><br>%{x|%b %Y}"
                      "<br>Value: %{y:,.2f} kg<extra></extra>",
    ))

    # Hampel cleansed history
    fig.add_trace(go.Scatter(
        x=key_df["Date"], y=key_df["Filtered"],
        mode="lines+markers", name="Hampel cleansed history",
        line=dict(color="#2ca02c", width=2, dash="dash"),
        marker=dict(size=5, symbol="square"),
        hovertemplate="<b>Hampel cleansed</b><br>%{x|%b %Y}"
                      "<br>Value: %{y:,.2f} kg<extra></extra>",
    ))

    # History For Forecast (kept as-is) — only if available
    if hff_df is not None and not hff_df.empty:
        fig.add_trace(go.Scatter(
            x=hff_df["Date"], y=hff_df["Value"],
            mode="lines", name="History For Forecast (unchanged)",
            line=dict(color=SERIES_STYLE[HISTORY_FOR_FORECAST_LABEL]["color"],
                      width=1.6),
            hovertemplate="<b>History For Forecast</b><br>%{x|%b %Y}"
                          "<br>Value: %{y:,.2f} kg<extra></extra>",
        ))

    # Outlier markers (on the original series)
    out_df = key_df[key_df["IsOutlier"]]
    if not out_df.empty:
        fig.add_trace(go.Scatter(
            x=out_df["Date"], y=out_df["Value"],
            mode="markers", name="Outlier (corrected)",
            marker=dict(color="#d62728", size=13, symbol="x-thin",
                        line=dict(color="#d62728", width=3)),
            hovertemplate="<b>OUTLIER</b><br>%{x|%b %Y}<br>"
                          "Value: %{y:,.2f}<extra></extra>",
        ))

    fig.update_layout(
        title=f"{key} — Hampel filter (window ±{window_size}, n_sigma = {n_sigma})",
        xaxis_title="Months", yaxis_title="Demand volume in kgs",
        hovermode="x unified", template="plotly_white",
        height=520, margin=dict(t=70, l=60, r=30, b=140),
        legend=dict(
            orientation="h",
            yanchor="top", y=-0.18,
            xanchor="center", x=0.5,
        ),
    )
    fig.update_xaxes(tickformat="%b %Y", showgrid=True, gridcolor="#eee")
    fig.update_yaxes(showgrid=True, gridcolor="#eee", tickformat=",.0f")
    st.plotly_chart(fig, use_container_width=True)

    n_out = int(key_df["IsOutlier"].sum())
    if n_out:
        st.warning(
            f"⚠️ {n_out} outlier observation(s) detected and corrected for this Key."
        )
    else:
        st.success("✅ No outliers detected for this Key.")


# ---------------------------------------------------------------------------
# Main entry-point
# ---------------------------------------------------------------------------
def main() -> None:
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
            "splitting **History** vs **Forecast** series.\n"
            "2. A **Hampel-filter** outlier detection & correction view for "
            "the **Sales History** time series of every Key, producing a "
            "cleansed history alongside the original."
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

    tab1, tab2 = st.tabs(["📈 Forecast vs Actuals", "🚨 Outlier Detection & Correction"])
    with tab1:
        render_dashboard_tab(long_df, uploaded.name)
    with tab2:
        render_outlier_tab(long_df)


if __name__ == "__main__":
    main()
