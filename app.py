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
import inspect
import io
import re
from datetime import datetime
from typing import List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio
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
    """Page styling that the theme in .streamlit/config.toml can't express.

    Colours, fonts, radii and widget borders come from the config theme; this
    only adds the app header, panel titles, tab bar, metric cards, button
    hierarchy and callout tones. Selectors target Streamlit's stable
    data-testid hooks, never generic tags.
    """
    st.markdown(
        """
        <style>
        :root {
          --fva-plane: #0f141b; --fva-surface: #161c25; --fva-raised: #1a212b;
          --fva-line: #2a3441; --fva-ink: #e6ebf1; --fva-ink-2: #a7b1bd;
          --fva-ink-3: #7c8794; --fva-accent: #4c9be8;
          --fva-forecast: #d95926;
        }
        [data-testid="stMainBlockContainer"] {
          padding-top: 2.25rem; padding-bottom: 4rem; max-width: 1480px;
        }
        [data-testid="stHeader"] { background: transparent; }

        /* Numbers line up in tables, metrics and captions. */
        [data-testid="stMetricValue"], [data-testid="stDataFrame"],
        .fva-cycle { font-variant-numeric: tabular-nums; }

        /* ---- App header ---------------------------------------------- */
        .fva-header-title {
          font-size: 1.6rem; font-weight: 600; letter-spacing: -0.01em;
          color: var(--fva-ink); line-height: 1.2; margin: 0;
        }
        .fva-header-sub { color: var(--fva-ink-2); font-size: 0.875rem;
          margin-top: 0.25rem; }
        .fva-cycle { display: flex; flex-wrap: wrap; gap: 0; margin-top: 0.9rem;
          border: 1px solid var(--fva-line); border-radius: 0.375rem;
          width: fit-content; overflow: hidden; }
        .fva-cycle > div { padding: 0.45rem 0.9rem; border-right: 1px solid var(--fva-line);
          background: var(--fva-surface); }
        .fva-cycle > div:last-child { border-right: none; }
        .fva-cycle .lbl { display: block; color: var(--fva-ink-3); font-size: 0.72rem; }
        .fva-cycle .val { display: block; color: var(--fva-ink); font-size: 0.95rem;
          font-weight: 500; }
        .fva-cycle .fcst .val { color: #f0875a; }

        /* ---- Tabs ---------------------------------------------------- */
        .stTabs [data-baseweb="tab-list"] { gap: 1.75rem;
          border-bottom: 1px solid var(--fva-line); }
        .stTabs [data-baseweb="tab"] { padding: 0.65rem 0 0.7rem 0;
          color: var(--fva-ink-3); font-weight: 500; font-size: 0.95rem; }
        .stTabs [data-baseweb="tab"]:hover { color: var(--fva-ink); }
        .stTabs [aria-selected="true"] { color: var(--fva-ink) !important; }
        .stTabs [data-baseweb="tab-highlight"] { background-color: var(--fva-accent);
          height: 2px; }
        .stTabs [data-baseweb="tab-border"] { display: none; }
        .stTabs [data-baseweb="tab-panel"] { padding-top: 1.5rem; }

        /* ---- Section rhythm ------------------------------------------- */
        [data-testid="stHeading"] h3 { margin-top: 1.25rem; }
        [data-testid="stHeading"] h4 { margin-top: 0.5rem; color: var(--fva-ink); }
        .fva-tab-intro { color: var(--fva-ink-2); font-size: 0.925rem;
          max-width: 78ch; margin: -0.35rem 0 1rem 0; line-height: 1.5; }

        /* ---- Filters panel --------------------------------------------- */
        [class*="__filter_panel"] { background: var(--fva-surface); }
        .fva-panel-title { color: var(--fva-ink); font-weight: 600;
          font-size: 0.95rem; }
        .fva-panel-note { color: var(--fva-ink-3); font-size: 0.8rem;
          margin-bottom: 0.25rem; }
        [data-testid="stMultiSelect"] label p { color: var(--fva-ink-2);
          font-size: 0.8rem; }
        [data-baseweb="select"] [data-baseweb="tag"] {
          background-color: #22364f !important; border: 1px solid #2f4d70; }
        [data-baseweb="select"] [data-baseweb="tag"] span { color: #dce9f7; }
        /* Placeholder ("All 12", "4 of 12 available") stays readable. */
        [data-baseweb="select"] [aria-hidden="true"] + div input::placeholder,
        [data-baseweb="select"] div[class*="placeholder"] {
          color: var(--fva-ink-3) !important; opacity: 1; }

        /* ---- Key line + reason chips (anomaly drill-down) --------------- */
        .fva-keyline { display: flex; flex-wrap: wrap; align-items: center;
          gap: 0.5rem; margin: 0.25rem 0 0.5rem 0; }
        .fva-key { font-weight: 600; color: var(--fva-ink); margin-right: 0.25rem; }
        .fva-sev { color: var(--fva-ink-2); font-size: 0.85rem;
          padding-right: 0.5rem; border-right: 1px solid var(--fva-line); }
        .fva-chip { font-size: 0.78rem; color: #dce9f7; background: #22364f;
          border: 1px solid #2f4d70; border-radius: 999px; padding: 0.1rem 0.6rem; }

        /* ---- Metric cards ---------------------------------------------- */
        [data-testid="stMetric"] {
          background: var(--fva-surface); border: 1px solid var(--fva-line);
          border-radius: 0.375rem; padding: 0.8rem 1rem 0.85rem 1rem; }
        [data-testid="stMetricLabel"] p { color: var(--fva-ink-2);
          font-size: 0.8rem; }
        [data-testid="stMetricValue"] { font-size: 1.6rem; font-weight: 500;
          color: var(--fva-ink); }

        /* ---- Buttons: secondary actions are quiet ----------------------- */
        .stButton button, .stDownloadButton button {
          background: var(--fva-raised); border: 1px solid var(--fva-line);
          color: var(--fva-ink); font-weight: 500; }
        .stButton button:hover, .stDownloadButton button:hover {
          border-color: var(--fva-accent); color: #ffffff; }
        .stButton button:focus-visible, .stDownloadButton button:focus-visible {
          outline: 2px solid var(--fva-accent); outline-offset: 2px; }
        .stFormSubmitButton button { background: var(--fva-accent);
          border: 1px solid var(--fva-accent); color: #08111c; font-weight: 600; }

        /* ---- Callouts (interpretations) --------------------------------- */
        [data-testid="stAlert"] { border-radius: 0.375rem; }
        [data-testid="stAlertContainer"] { border: 1px solid var(--fva-line); }

        /* ---- Expanders and uploader ------------------------------------- */
        [data-testid="stExpander"] details { background: var(--fva-surface);
          border-color: var(--fva-line); }
        [data-testid="stFileUploaderDropzone"] { background: var(--fva-surface);
          border: 1px dashed #3a4656; }

        /* ---- Small screens --------------------------------------------- */
        @media (max-width: 760px) {
          .stTabs [data-baseweb="tab-list"] { gap: 1rem; overflow-x: auto; }
          .fva-header-title { font-size: 1.3rem; }
        }

        /* ---- Fullscreen charts keep the dark canvas --------------------- */
        [data-testid="stFullScreenFrame"], :fullscreen, ::backdrop {
          background-color: var(--fva-plane) !important; }
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
    _, mid, _ = st.columns([1, 1.1, 1])
    with mid:
        st.markdown(
            '<div style="height:12vh"></div>'
            '<div class="fva-header-title">Forecast vs Actuals</div>'
            '<div class="fva-header-sub">Demand planning review. '
            'Sign in with your team account.</div>'
            '<div style="height:1rem"></div>',
            unsafe_allow_html=True)
        with st.form("login_form", clear_on_submit=False):
            username = st.text_input("Username", key="login::username",
                                     autocomplete="username")
            password = st.text_input("Password", type="password",
                                     key="login::password",
                                     autocomplete="current-password")
            submitted = st.form_submit_button("Sign in",
                                              width="stretch")

        if submitted:
            if verify_credentials(username, password):
                st.session_state["authenticated"] = True
                st.session_state["auth_user"] = username.strip().lower()
                # Stamp the login time. Streamlit keeps a session alive for as
                # long as the browser tab holds the websocket, so an active
                # user stays signed in well beyond a few hours; this timestamp
                # lets the app show/verify session age if ever needed.
                st.session_state["login_time"] = datetime.now().isoformat()
                # Don't keep the raw password in session state.
                st.session_state.pop("login::password", None)
                st.rerun()
            else:
                st.error("That username and password don't match. Check both and try again.")

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
    "Arkieva Active Status Lag 1",
    "Arkieva Pattern Lag 1",
    "Arkieva Review Req",
    "Data",
]

ACTIVE_STATUS_COL = "Arkieva Active Status"
PATTERN_COL = "Arkieva Pattern"
# Prior-cycle (Lag 1) attribute columns, used by the STF driver waterfall to
# detect status / pattern changes between the current and prior cycle. These
# are optional — older files without them fall back to "no change".
ACTIVE_STATUS_LAG_COL = "Arkieva Active Status Lag 1"
PATTERN_LAG_COL = "Arkieva Pattern Lag 1"
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

# STF (Statistical Forecast Committed) current vs prior-cycle committed —
# used by the "STF Variation & Exceptions" tab. The current committed line is
# normalised to STAT_FORECAST_LABEL on load, so that label IS the current STF.
# The prior-cycle committed line may be labelled "... Lag 1", "... Lag 2", etc.
# depending on the export cycle; every such variant is normalised on load to
# the single canonical STF_LAG1_LABEL so the rest of the app is label-agnostic.
STF_CURRENT_LABEL = STAT_FORECAST_LABEL
STF_LAG1_LABEL = "Statistical Forecast Committed Lag (kg)"
# Regex matching any committed-lag variant, e.g. "Statistical Forecast
# Committed Lag 1 (kg)", "... Lag 2 (kg)", "... Lag 3 (kg)".
STF_LAG_PATTERN = r"^Statistical Forecast Committed Lag\s*\d+\s*\(kg\)$"
# Minimum prior-cycle volume (kg, summed over the horizon or month) for a
# % variance to be meaningful.
STF_MIN_BASE_KG = 1.0

# Some input files label the statistical-forecast line as
# "Statistical Forecast Committed (kg)". Any of these aliases is normalised
# to the canonical STAT_FORECAST_LABEL when the file is loaded.
STAT_FORECAST_ALIASES = [
    "Statistical Forecast Committed (kg)",
    "Statistical Forecast (kg)",
]

# Categorisation of the "Data" series. History lines span the full timeline;
# Forecast lines are drawn ONLY in the Forecast period (after the last active
# Sales-History month), so planners can read historical sales/seasonality
# cleanly in the History band.
# Order = drawing order: actual Sales History is drawn last of the two,
# so it sits on top of the (usually identical) History For Forecast line.
HISTORY_SERIES = ["History For Forecast (kg)", "Sales History (kg)"]
FORECAST_SERIES = ["Statistical Forecast (kg)", STF_LAG1_LABEL,
                   "Final Demand Plan Lag 1 (kg)"]

# ---------------------------------------------------------------------------
# Design tokens (mirror .streamlit/config.toml)
# ---------------------------------------------------------------------------
# A calm "ink-slate" dark theme built for long review sessions and screen
# sharing: one steel-blue accent for interactive state, recessive chart
# chrome, and series colours that mean the same thing on every tab.
INK_PLANE = "#0f141b"       # page + chart background
INK_SURFACE = "#161c25"     # panels, hover labels
INK_RAISED = "#1a212b"      # inputs, buttons
INK_LINE = "#2a3441"        # hairlines and borders
INK_GRID = "#1e2631"        # chart gridlines (deliberately quiet)
TEXT_PRIMARY = "#e6ebf1"
TEXT_SECONDARY = "#a7b1bd"
TEXT_MUTED = "#7c8794"
ACCENT = "#4c9be8"
FONT_FAMILY = '"IBM Plex Sans", system-ui, -apple-system, "Segoe UI", sans-serif'

# Names kept for the tab code that predates the token names.
DARK_BG = INK_PLANE
DARK_PANEL = INK_SURFACE
DARK_GRID = INK_GRID
DARK_TEXT = TEXT_PRIMARY
DARK_MUTED = TEXT_MUTED

# Categorical order, validated against INK_PLANE with the dataviz palette
# validator: every adjacent pair clears the colour-blind (ΔE ≥ 8) and
# normal-vision (ΔE ≥ 15) separation floors, and every colour clears 3:1
# contrast. Colours are assigned in this fixed order, never cycled.
CATEGORICAL = ["#3987e5", "#d95926", "#199e70", "#c98500",
               "#d55181", "#008300", "#9085e9", "#e66767"]
# Sequential (magnitude): one blue hue, dark → light on the dark plane.
SEQUENTIAL_BLUE = [[0.0, "#17283d"], [0.25, "#184f95"], [0.5, "#2a78d6"],
                   [0.75, "#5598e7"], [1.0, "#9ec5f4"]]
# Diverging (signed change): blue = lower, red = higher, neutral midpoint.
DIVERGING_BLUE_RED = [[0.0, "#2a78d6"], [0.5, "#273140"], [1.0, "#e66767"]]

# Series colours follow the series on every chart. History vs Forecast is
# also told apart by marker shape (circles vs diamonds), never colour alone.
SERIES_STYLE = {
    "Sales History (kg)":           {"color": "#3987e5", "dash": "solid", "category": "History"},
    "History For Forecast (kg)":    {"color": "#199e70", "dash": "solid", "category": "History"},
    "Statistical Forecast (kg)":    {"color": "#d95926", "dash": "solid", "category": "Forecast"},
    STF_LAG1_LABEL:                 {"color": "#9085e9", "dash": "solid", "category": "Forecast"},
    "Final Demand Plan Lag 1 (kg)": {"color": "#d55181", "dash": "solid", "category": "Forecast"},
}

# Only the forecast period gets a wash; history sits on the plain plane.
HISTORY_BAND_COLOR = "rgba(0, 0, 0, 0)"
FORECAST_BAND_COLOR = "rgba(217, 89, 38, 0.06)"

# A trend line wears the colour of the series it is fitted to (dashed).
HISTORY_TREND_COLOR = SERIES_STYLE["Sales History (kg)"]["color"]
FORECAST_TREND_COLOR = SERIES_STYLE["Statistical Forecast (kg)"]["color"]

# STF change buckets: a neutral grey for "no change" so real changes stand
# out; the three change colours validate on every pair (stacked bars).
# Cleansed (corrected) history in the Outliers tab: categorical slot 4.
CLEANSED_COLOR = "#c98500"

CHANGE_BUCKET_COLORS = {
    "No Changes": "#6b6c70",
    "Status Change": "#3987e5",
    "Pattern Change": "#199e70",
    "Status & Pattern Change": "#d95926",
}

# Plotly template registered once at import (cheaper than restyling every
# figure on every rerun, and keeps all charts consistent).
pio.templates["fva_dark"] = go.layout.Template(
    layout=dict(
        font=dict(family=FONT_FAMILY, size=12, color=TEXT_SECONDARY),
        paper_bgcolor=INK_PLANE,
        plot_bgcolor=INK_PLANE,
        colorway=CATEGORICAL,
        hovermode="x unified",
        hoverlabel=dict(bgcolor=INK_SURFACE, bordercolor=INK_LINE,
                        font=dict(family=FONT_FAMILY, size=12,
                                  color=TEXT_PRIMARY)),
        legend=dict(orientation="h", yanchor="bottom", y=1.02,
                    xanchor="left", x=0, bgcolor="rgba(0,0,0,0)",
                    font=dict(size=12, color=TEXT_SECONDARY),
                    title=dict(font=dict(size=12, color=TEXT_MUTED)),
                    itemclick="toggle", itemdoubleclick="toggleothers"),
        xaxis=dict(showgrid=False, zeroline=False, showline=True,
                   linecolor=INK_LINE, ticks="outside", ticklen=4,
                   tickcolor=INK_LINE, tickfont=dict(color=TEXT_MUTED),
                   title=dict(font=dict(size=12, color=TEXT_MUTED),
                              standoff=10),
                   showspikes=True, spikemode="across", spikesnap="cursor",
                   spikethickness=1, spikecolor=TEXT_MUTED,
                   spikedash="solid"),
        yaxis=dict(showgrid=True, gridcolor=INK_GRID, gridwidth=1,
                   zeroline=False, showline=False, ticks="",
                   tickfont=dict(color=TEXT_MUTED),
                   title=dict(font=dict(size=12, color=TEXT_MUTED),
                              standoff=8)),
        bargap=0.3,
        bargroupgap=0.08,
        margin=dict(t=56, l=64, r=24, b=48),
        colorscale=dict(sequential=SEQUENTIAL_BLUE,
                        diverging=DIVERGING_BLUE_RED),
        annotationdefaults=dict(font=dict(color=TEXT_SECONDARY, size=11)),
    ),
    data=dict(
        scatter=[go.Scatter(line=dict(width=2), marker=dict(size=6))],
        bar=[go.Bar(marker=dict(line=dict(width=0)))],
        heatmap=[go.Heatmap(colorscale=SEQUENTIAL_BLUE, xgap=2, ygap=2,
                            colorbar=dict(outlinewidth=0, thickness=10,
                                          tickfont=dict(color=TEXT_MUTED)))],
    ),
)

PLOTLY_CONFIG = {
    "displaylogo": False,
    "modeBarButtonsToRemove": ["lasso2d", "select2d", "autoScale2d"],
    "toImageButtonOptions": {"format": "png", "scale": 2},
}
_PLOTLY_ACCEPTS_WIDTH = "width" in inspect.signature(st.plotly_chart).parameters


# ---------------------------------------------------------------------------
# Safe table rendering
# ---------------------------------------------------------------------------
# pandas' Styler can only render `pd.options.styler.render.max_elements` cells
# (262,144 by default). Streamlit raises a StreamlitAPIException past that, and
# Styler is punishingly slow long before the ceiling. Tables whose size is
# data-dependent (e.g. every flagged outlier) can blow through it, so those are
# rendered WITHOUT a Styler: values stay numeric (still sortable) and the
# formatting is applied by Streamlit's column_config, which is Arrow-based and
# virtualised, so it handles any row count quickly.
STYLER_CELL_LIMIT = 150_000  # safety margin below pandas' 262,144 default


def _fmt_decimals(spec: str) -> int:
    """Number of decimals in a format spec like '{:,.2f}' → 2."""
    m = re.search(r"\.(\d+)f", spec)
    return int(m.group(1)) if m else 0


def render_numeric_table(df: pd.DataFrame, fmt_map: dict, na_rep: str = "–",
                         **kwargs) -> None:
    """Render ``df`` with per-column number formatting.

    Small tables use a pandas Styler (identical look to before). Tables above
    ``STYLER_CELL_LIMIT`` cells skip the Styler and use column_config instead,
    which avoids Streamlit's Styler cell-limit exception entirely.
    """
    if df.size <= STYLER_CELL_LIMIT:
        st.dataframe(df.style.format(fmt_map, na_rep=na_rep), **kwargs)
        return

    disp = df.copy()
    col_cfg = {}
    for col, spec in fmt_map.items():
        if col not in disp.columns:
            continue
        dec = _fmt_decimals(spec)
        disp[col] = pd.to_numeric(disp[col], errors="coerce").round(dec)
        if spec.rstrip().endswith("%"):
            col_cfg[col] = st.column_config.NumberColumn(col, format=f"%.{dec}f%%")
        else:
            # "localized" adds thousands separators, matching '{:,.Nf}'.
            col_cfg[col] = st.column_config.NumberColumn(col, format="localized")
    st.dataframe(disp, column_config=col_cfg, **kwargs)


_LEGEND_STYLE_KEYS = {"traceorder", "groupclick", "title", "tracegroupgap",
                      "itemsizing"}


def dark_layout(fig: "go.Figure", **overrides) -> "go.Figure":
    """Apply the shared chart style (the ``fva_dark`` template).

    Section headers in the page already name each chart, so the in-chart
    title is not drawn; it is kept as the PNG export file name instead.
    Legends always sit top-left above the plot, so callers may only adjust
    legend *grouping*, not its position or colours.
    """
    title = overrides.pop("title", None)
    if isinstance(title, dict):
        title = title.get("text")
    legend = overrides.pop("legend", None) or {}
    margin = dict(overrides.pop("margin", None) or {})
    # Bottom margins were sized for the old under-the-chart legend.
    margin["t"] = max(margin.get("t", 56), 56)
    margin["b"] = min(margin.get("b", 48), 64)
    has_cycle_marks = any(str(a.text or "").startswith("Last actual")
                          for a in (fig.layout.annotations or ()))
    if has_cycle_marks:
        # Boundary labels sit above the plot: lift the legend over them.
        margin["t"] = max(margin["t"], 84)
        legend = dict(legend, y=1.09)
    # Backgrounds are set explicitly as well as in the template: Streamlit's
    # frontend fills in its own plot background when only a template sets it.
    base = dict(template="fva_dark", margin=margin,
                paper_bgcolor=INK_PLANE, plot_bgcolor=INK_PLANE,
                title=dict(text=None),
                meta={"export_name": str(title or "chart")})
    if legend:
        base["legend"] = {k: v for k, v in legend.items()
                          if k in _LEGEND_STYLE_KEYS or k == "y"}
    base.update(overrides)
    fig.update_layout(**base)
    return fig


def add_time_controls(fig: "go.Figure") -> "go.Figure":
    """Add a slim range slider and quick range buttons to a time chart."""
    fig.update_xaxes(
        rangeslider=dict(visible=True, thickness=0.05, bgcolor=INK_SURFACE,
                         bordercolor=INK_LINE, borderwidth=1),
        rangeselector=dict(
            x=1, xanchor="right", y=1.02, yanchor="bottom",
            bgcolor=INK_RAISED, activecolor="#24476e",
            bordercolor=INK_LINE, borderwidth=1,
            font=dict(color=TEXT_SECONDARY, size=11),
            buttons=[
                dict(count=6, label="6M", step="month", stepmode="backward"),
                dict(count=1, label="1Y", step="year", stepmode="backward"),
                dict(count=2, label="2Y", step="year", stepmode="backward"),
                dict(step="all", label="All"),
            ],
        ),
        tickformat="%b %Y",
    )
    return fig


def show_chart(fig: "go.Figure", key: Optional[str] = None) -> None:
    """Render a Plotly figure with the app's own template.

    ``theme=None`` stops Streamlit from overriding the template (that
    override is what made chart titles render as faint grey), and the
    sizing argument is chosen for the installed Streamlit version, so no
    deprecation banner is shown in the app.
    """
    config = dict(PLOTLY_CONFIG)
    name = (fig.layout.meta or {}).get("export_name", "chart") \
        if isinstance(fig.layout.meta, dict) else "chart"
    safe = re.sub(r"[^A-Za-z0-9]+", "_", str(name)).strip("_")[:60] or "chart"
    config["toImageButtonOptions"] = dict(config["toImageButtonOptions"],
                                          filename=safe)
    kwargs = dict(theme=None, config=config, key=key)
    if _PLOTLY_ACCEPTS_WIDTH:
        kwargs["width"] = "stretch"
    st.plotly_chart(fig, **kwargs)


# Display names for the Data lines (the canonical prior-cycle label is an
# internal normalisation, so planners see what it actually is).
SERIES_DISPLAY_NAME = {
    STF_LAG1_LABEL: "Statistical Forecast, prior cycle (kg)",
}
# Legend order (independent of drawing order).
SERIES_RANK = {SALES_HISTORY_LABEL: 1, HISTORY_FOR_FORECAST_LABEL: 2,
               STAT_FORECAST_LABEL: 3, STF_LAG1_LABEL: 4,
               "Final Demand Plan Lag 1 (kg)": 5}


def format_kg(value: float) -> str:
    """Compact kg figure for cards: 245.6 M kg, 5.64 bn kg, 12.3 k kg."""
    if value is None or pd.isna(value):
        return "–"
    v = float(value)
    a = abs(v)
    if a >= 1e9:
        return f"{v / 1e9:,.2f} bn kg"
    if a >= 1e6:
        return f"{v / 1e6:,.1f} M kg"
    if a >= 1e4:
        return f"{v / 1e3:,.1f} k kg"
    return f"{v:,.0f} kg"


def entity_colors(df: pd.DataFrame, column: str) -> dict:
    """Fixed colour per value of ``column`` across the whole file, so a
    Business Line or region keeps its colour when filters remove others.
    Colours are never cycled: values past the eighth are drawn in neutral
    grey (each bar still names its value on hover)."""
    values = sorted(filter_dims(df, [column])[column].dropna().astype(str).unique())
    return {v: (CATEGORICAL[i] if i < len(CATEGORICAL) else "#6b6c70")
            for i, v in enumerate(values)}


def material_label(material: str) -> str:
    """'Dissolvine E-39, bulk__5401493' -> 'Dissolvine E-39, bulk (5401493)'.
    The code stays visible: it keeps same-named materials distinct."""
    text = str(material)
    if "__" in text:
        name, code = text.rsplit("__", 1)
        return f"{name.strip()} ({code.strip()})"
    return text


def mask_trailing_zeros(series: pd.Series) -> pd.Series:
    """Blank the zero run at the end of a forecast line.

    A forecast cycle that stops short of the file's last month is exported
    as zeros, which drew a plunge to 0 that read as 'forecast is zero'.
    Interior zeros are real values and stay.
    """
    nz = series[(series.notna()) & (series != 0)]
    if nz.empty:
        return series.where(series != 0)
    return series.where(series.index <= nz.index.max())


def add_cycle_marks(fig: "go.Figure", boundary, x_max) -> None:
    """The shared history/forecast marker: a faint wash over the forecast
    period and a labelled line at the last actual month (labels sit just
    above the plot area, so they never cover data)."""
    boundary = pd.Timestamp(boundary)
    if pd.Timestamp(x_max) > boundary:
        fig.add_vrect(x0=boundary, x1=x_max, fillcolor=FORECAST_BAND_COLOR,
                      line_width=0, layer="below")
        fig.add_annotation(x=boundary, xref="x", y=1, yref="paper",
                           text="Forecast", showarrow=False, xanchor="left",
                           yanchor="bottom", xshift=6, yshift=2,
                           font=dict(size=11, color="#f0875a"))
    fig.add_vline(x=boundary, line=dict(color="#3d4a5c", width=1))
    fig.add_annotation(x=boundary, xref="x", y=1, yref="paper",
                       text=f"Last actual {boundary.strftime('%b %Y')}",
                       showarrow=False, xanchor="right", yanchor="bottom",
                       xshift=-6, yshift=2,
                       font=dict(size=11, color=TEXT_MUTED))


def time_series_trace(x, y, data_label: str, style: dict,
                      name: Optional[str] = None, legend_group: bool = True,
                      **overrides) -> "go.Scatter":
    """A monthly line drawn with straight segments (no spline overshoot);
    history uses circles, forecast diamonds, so identity is never colour
    alone."""
    category = style.get("category", "Other")
    kwargs = dict(
        x=x, y=y, mode="lines+markers",
        name=name or SERIES_DISPLAY_NAME.get(data_label, data_label),
        line=dict(color=style["color"], width=2),
        marker=dict(size=5, color=style["color"],
                    symbol="circle" if category == "History" else "diamond"),
        hovertemplate="%{y:,.0f} kg",
        connectgaps=False,
        legendrank=SERIES_RANK.get(data_label, 50),
    )
    if legend_group:
        kwargs.update(legendgroup=category, legendgrouptitle_text=category)
    kwargs.update(overrides)
    return go.Scatter(**kwargs)


def trend_trace(trend: dict, name: str, color: str) -> "go.Scatter":
    """A fitted trend wears its series' colour, dashed and lighter."""
    return go.Scatter(
        x=trend["dates"], y=trend["fit_line"], mode="lines", name=name,
        line=dict(color=color, width=1.5, dash="dash"), opacity=0.85,
        legendgroup="Trend", legendgrouptitle_text="Trend",
        hovertemplate="%{y:,.0f} kg",
    )


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

    # Optional prior-cycle attribute columns (used by the STF driver
    # waterfall). When absent, default them to the current-cycle value so the
    # change detection reports "no change".
    if ACTIVE_STATUS_LAG_COL not in raw.columns:
        raw[ACTIVE_STATUS_LAG_COL] = raw[ACTIVE_STATUS_COL]
    if PATTERN_LAG_COL not in raw.columns:
        raw[PATTERN_LAG_COL] = raw[PATTERN_COL]

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
    # Normalise the prior-cycle committed line ("... Lag 1/2/3 (kg)") to the
    # single canonical STF_LAG1_LABEL so the STF variance engine is agnostic
    # to which lag cycle the export carries.
    long_df["Data"] = long_df["Data"].str.replace(
        STF_LAG_PATTERN, STF_LAG1_LABEL, regex=True
    )

    # Clean the prior-cycle attribute columns to readable strings.
    for c in (ACTIVE_STATUS_LAG_COL, PATTERN_LAG_COL):
        if c in long_df.columns:
            long_df[c] = long_df[c].astype(str).str.strip()

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

    # File fingerprint, carried by pandas through row filtering. Lets the
    # filter panels cache their small option table per file (see filter_dims).
    long_df.attrs["fva_token"] = hashlib.sha1(file_bytes).hexdigest()[:16]

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
# Why filters used to "vanish":
#   Streamlit identifies a widget by its key *plus* some of its parameters.
#   For st.multiselect / st.selectbox the option list is part of that
#   identity, and for st.select_slider / st.radio every parameter is. A
#   cascading filter changes the option lists of the OTHER filters whenever
#   one filter changes, so Streamlit treated those filters as brand-new
#   widgets and discarded their selections (and a plain rerun then wiped the
#   rest).
#
# The fix: every data-dependent widget keeps its value in a separate
# "store" key that Streamlit never discards. An on_change callback copies
# the user's choice into the store, and on every run the widget is
# re-seeded from the store *before* it is created (the only time Streamlit
# lets code set a widget's value). Re-creating a widget can therefore no
# longer lose the selection.
def unique_sorted(series: pd.Series) -> List[str]:
    return sorted(series.dropna().astype(str).unique().tolist())


def _filter_store_key(key_prefix: str, col: str) -> str:
    return f"{key_prefix}::{col}"


def _filter_widget_key(key_prefix: str, col: str) -> str:
    return f"{key_prefix}::{col}::w"


def _sync_widget_to_store(widget_key: str, store_key: str) -> None:
    """on_change callback: copy a widget's new value into its store key."""
    value = st.session_state.get(widget_key)
    if isinstance(value, tuple):
        value = tuple(value)
    elif isinstance(value, list):
        value = list(value)
    st.session_state[store_key] = value


def apply_filters(df: pd.DataFrame, selections: dict) -> pd.DataFrame:
    """Rows of ``df`` matching every non-empty selection.

    Builds one combined boolean mask and slices once (instead of slicing the
    1M-row table once per filter). Filter columns are already strings after
    ``load_excel``; non-string columns are compared as strings.
    """
    mask = None
    for col, sel in selections.items():
        if not sel or col not in df.columns:
            continue
        series = df[col]
        if series.dtype != object:
            series = series.astype(str)
        m = series.isin([str(v) for v in sel]).to_numpy()
        mask = m if mask is None else (mask & m)
    if mask is None:
        return df
    return df[mask]


def _df_token(df: pd.DataFrame) -> Optional[str]:
    """Cheap identity for a frame derived from the loaded file.

    ``load_excel`` stamps ``attrs["fva_token"]`` (the file fingerprint);
    pandas carries attrs through row filtering, so the Sales-History subset
    used by the outlier tab gets its own token via its row count.
    """
    token = df.attrs.get("fva_token")
    return None if token is None else f"{token}|{len(df)}"


@st.cache_data(show_spinner=False, max_entries=64)
def _cached_filter_dims(token: str, columns: Tuple[str, ...],
                        _df: pd.DataFrame) -> pd.DataFrame:
    return _df[list(columns)].drop_duplicates().reset_index(drop=True)


def filter_dims(df: pd.DataFrame, columns: List[str]) -> pd.DataFrame:
    """Unique combinations of the filter columns (~3,400 rows instead of
    ~1M). The cascade is computed on this small table, which is what keeps
    filter clicks fast."""
    token = _df_token(df)
    if token is None:
        return df[list(columns)].drop_duplicates().reset_index(drop=True)
    return _cached_filter_dims(token, tuple(columns), df)


def cascading_options(
    df: pd.DataFrame,
    columns: List[str],
    current_selections: dict,
) -> dict:
    """For each filter column, the values consistent with the selections in
    **all other** filter columns (Excel-slicer behaviour: a filter never
    narrows its own list). ``df`` may be the full table or its dims table.
    """
    options: dict = {}
    for col in columns:
        others = {c: v for c, v in current_selections.items()
                  if c != col and v}
        options[col] = unique_sorted(apply_filters(df, others)[col])
    return options


def _reset_filter_strip(key_prefix: str, columns: Tuple[str, ...],
                        defaults: dict) -> None:
    """on_click callback for the Reset button (runs before the rerun, so it
    may freely change widget keys)."""
    for col in columns:
        value = list(defaults.get(col, []))
        st.session_state[_filter_store_key(key_prefix, col)] = value
        st.session_state[_filter_widget_key(key_prefix, col)] = value


def active_filter_summary(selections: dict) -> str:
    """Plain-language summary of the active filters, e.g.
    'Business Line: Cleaning__BL02 · Arkieva Active Status: 2 selected'."""
    parts = []
    for col, sel in selections.items():
        if not sel:
            continue
        parts.append(f"{col}: {sel[0]}" if len(sel) == 1
                     else f"{col}: {len(sel)} selected")
    return " · ".join(parts)


def render_filter_strip(
    df: pd.DataFrame,
    columns: List[str],
    key_prefix: str,
    n_per_row: int = 4,
    default_selections: Optional[dict] = None,
) -> dict:
    """Render a cascading multi-select filter strip and return the
    selections as ``{column: [values]}``.

    Behaviour matches Excel slicers:
    * Each filter lists only values consistent with the *other* filters.
    * Selecting values in one filter narrows the options in every other.
    * Selections persist across reruns, other widgets and tab switches.
      A value you picked is always kept in its list, so the cascade never
      silently drops it; clearing it is always your decision.

    ``default_selections`` (e.g. Arkieva Active Status → Active + Sparse)
    are applied the first time a tab's strip renders; afterwards the user's
    choices, including clearing a filter, are respected. Reset restores
    them. Each tab passes its own ``key_prefix``, so tabs filter
    independently.
    """
    default_selections = default_selections or {}
    dims = filter_dims(df, columns)
    all_values = {col: set(dims[col].astype(str)) for col in columns}

    # ---- One-time defaults (seed the stores, never the widgets) -----------
    init_flag = f"{key_prefix}::__initialized__"
    if not st.session_state.get(init_flag, False):
        for col, default_vals in default_selections.items():
            if col not in all_values:
                continue
            store_key = _filter_store_key(key_prefix, col)
            if store_key not in st.session_state:
                st.session_state[store_key] = [
                    v for v in default_vals if v in all_values[col]]
        st.session_state[init_flag] = True

    # ---- Current selections come from the stores ---------------------------
    current: dict = {}
    for col in columns:
        raw = st.session_state.get(_filter_store_key(key_prefix, col), [])
        if not isinstance(raw, (list, tuple)):
            raw = []
        current[col] = [str(v) for v in raw if str(v) in all_values[col]]

    options_per_col = cascading_options(dims, columns, current)

    # ---- Widgets -------------------------------------------------------------
    rows = [columns[i:i + n_per_row] for i in range(0, len(columns), n_per_row)]
    for row in rows:
        slots = st.columns(n_per_row)
        for slot, col_name in zip(slots, row):
            cascaded = options_per_col[col_name]
            chosen = current[col_name]
            # Keep chosen values in the list even if the other filters now
            # exclude them, so they stay visible and removable.
            opts = list(cascaded) + [v for v in chosen if v not in cascaded]
            full_count = len(all_values[col_name])
            placeholder = (f"All {full_count}" if len(cascaded) == full_count
                           else f"{len(cascaded)} of {full_count} available")
            widget_key = _filter_widget_key(key_prefix, col_name)
            store_key = _filter_store_key(key_prefix, col_name)
            # Seed the widget from its store BEFORE creating it.
            st.session_state[widget_key] = list(chosen)
            with slot:
                st.multiselect(
                    col_name,
                    options=opts,
                    key=widget_key,
                    placeholder=placeholder,
                    on_change=_sync_widget_to_store,
                    args=(widget_key, store_key),
                )

    # ---- Footer: active-filter summary + Reset --------------------------------
    valid_defaults = {
        col: [v for v in vals if v in all_values.get(col, set())]
        for col, vals in default_selections.items()
    }
    summary = active_filter_summary(current)
    left, right = st.columns([5, 1], vertical_alignment="center")
    with left:
        st.caption(summary if summary else "No filters applied — showing everything.")
    with right:
        st.button(
            "Reset filters", key=f"{key_prefix}::reset", width="stretch",
            on_click=_reset_filter_strip,
            args=(key_prefix, tuple(columns), valid_defaults),
            help="Restore this tab's default filters.",
        )

    return current


def filter_panel(
    df: pd.DataFrame,
    columns: List[str],
    key_prefix: str,
    default_selections: Optional[dict] = None,
) -> dict:
    """The bordered Filters panel every tab shows above its content."""
    with st.container(border=True, key=f"{key_prefix}__filter_panel"):
        note = ("Each filter narrows the others, like Excel slicers. "
                "Arkieva Active Status starts at Active + Sparse."
                if default_selections else
                "Each filter narrows the others, like Excel slicers. "
                "Leave a filter empty to include everything.")
        st.markdown(f'<div class="fva-panel-title">Filters</div>'
                    f'<div class="fva-panel-note">{note}</div>',
                    unsafe_allow_html=True)
        return render_filter_strip(df, columns, key_prefix,
                                   default_selections=default_selections)


# ---- Sticky single-value widgets ---------------------------------------------
# Same store-and-reseed pattern for pickers whose options depend on the data
# (e.g. the Key picker), so a filter change doesn't snap them back to the
# first item while the chosen item is still available.
def _store_key(widget_key: str) -> str:
    return f"{widget_key}::store"


def sticky_selectbox(label: str, options: list, key: str,
                     default_index: int = 0, **kwargs):
    if not options:
        return None
    store = _store_key(key)
    value = st.session_state.get(store)
    if value not in options:
        value = options[min(default_index, len(options) - 1)]
    st.session_state[key] = value
    return st.selectbox(label, options, key=key,
                        on_change=_sync_widget_to_store, args=(key, store),
                        **kwargs)


def sticky_radio(label: str, options: list, key: str,
                 default_index: int = 0, **kwargs):
    store = _store_key(key)
    value = st.session_state.get(store)
    if value not in options:
        value = options[min(default_index, len(options) - 1)]
    st.session_state[key] = value
    return st.radio(label, options, key=key,
                    on_change=_sync_widget_to_store, args=(key, store),
                    **kwargs)


def sticky_range_select_slider(label: str, options: list, key: str,
                               default: tuple, **kwargs) -> tuple:
    """Range select_slider that keeps the user's range.

    Unlike the other sticky widgets this one is NOT seeded through session
    state: a select_slider created without ``value`` starts as a single-thumb
    slider, and seeding a (start, end) range into it crashes the browser
    component (TypeError: reading 'current'). Passing the remembered range as
    ``value`` gives the same guarantee: when the options change (new filter
    selection) Streamlit re-creates the widget at the remembered range.
    """
    store = _store_key(key)
    value = st.session_state.get(store)
    if not (isinstance(value, (tuple, list)) and len(value) == 2
            and value[0] in options and value[1] in options):
        value = default
    return st.select_slider(label, options=options, value=tuple(value),
                            key=key, on_change=_sync_widget_to_store,
                            args=(key, store), **kwargs)


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
# The waterfall decomposes the net STF change (current − prior committed) into
# **change buckets** derived from whether each Key's Arkieva Active Status
# and/or Arkieva Pattern changed between the current and prior cycle:
#   • No Changes             — status same AND pattern same
#   • Pattern Change         — pattern differs, status same
#   • Status Change          — status differs, pattern same
#   • Status & Pattern Change— both differ
# Every Key falls into exactly one bucket, so the per-bucket deltas reconcile
# exactly to the net change. Drill-down then shows the specific status or
# pattern transitions (from → to) driving each change bucket.
STF_BUCKET_NONE = "No Changes"
STF_BUCKET_PATTERN = "Pattern Change"
STF_BUCKET_STATUS = "Status Change"
STF_BUCKET_BOTH = "Status & Pattern Change"
STF_BUCKET_ORDER = [STF_BUCKET_NONE, STF_BUCKET_STATUS,
                    STF_BUCKET_PATTERN, STF_BUCKET_BOTH]


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


def stf_forecast_months(
    all_months: List[pd.Timestamp],
    boundary: Optional[pd.Timestamp],
) -> List[pd.Timestamp]:
    """All forecast months available for STF variation — every month strictly
    after the history/forecast boundary (the last active Sales-History
    month). These are the months the planner can pick a range within."""
    if boundary is None:
        return []
    return [m for m in sorted(all_months) if m > pd.Timestamp(boundary)]


def stf_next_fiscal_years(boundary: Optional[pd.Timestamp], n: int = 2) -> List[int]:
    """The next ``n`` full fiscal years (Jan–Dec) after the current year.

    The current year is the year of the history/forecast boundary. e.g.
    boundary in 2026 → [2027, 2028]."""
    if boundary is None:
        return []
    cur_year = pd.Timestamp(boundary).year
    return [cur_year + i for i in range(1, n + 1)]


def stf_fiscal_year_months(
    all_months: List[pd.Timestamp],
    year: int,
) -> List[pd.Timestamp]:
    """The available months of a fiscal year (Jan–Dec of that year)."""
    return [m for m in sorted(all_months) if m.year == year]


def stf_months_between(
    all_months: List[pd.Timestamp],
    start: pd.Timestamp,
    end: pd.Timestamp,
) -> List[pd.Timestamp]:
    """Inclusive list of available months in [start, end]."""
    if start is None or end is None:
        return []
    lo, hi = min(start, end), max(start, end)
    return [m for m in sorted(all_months)
            if pd.Timestamp(lo) <= m <= pd.Timestamp(hi)]


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

    # A prior-cycle total under STF_MIN_BASE_KG has no meaningful base:
    # Arkieva exports float residues (~1e-22 kg) that previously produced
    # variances like 4e26 %. Those Keys get no % variance.
    base_ok = grp["Lag1STF"].abs() >= STF_MIN_BASE_KG
    grp["STF Variance"] = np.where(
        base_ok,
        (grp["CurrentSTF"] - grp["Lag1STF"]) / grp["Lag1STF"].where(base_ok, 1.0),
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
    base_ok = piv["Lag1STF"].abs() >= STF_MIN_BASE_KG
    piv["STF Variance"] = np.where(
        base_ok,
        (piv["CurrentSTF"] - piv["Lag1STF"]) / piv["Lag1STF"].where(base_ok, 1.0),
        np.nan,
    )
    piv["Absolute STF Variance"] = piv["STF Variance"].abs()
    return piv[cols]


# The change-bucket classification operates at the granular grain — a single
# Key can span several rows with different Arkieva Active Status / Pattern, so
# classifying at Key level would collapse (and misattribute) those. The grain
# is Key + current status + current pattern + prior-cycle status + pattern.
STF_GRAIN_COLS = ["Key", ACTIVE_STATUS_COL, PATTERN_COL,
                  ACTIVE_STATUS_LAG_COL, PATTERN_LAG_COL]

# Attribute values that all mean "no meaningful value" and must compare equal
# to each other, so a placeholder-to-placeholder move (e.g. "" ↔ "nan") is not
# mistaken for a real status/pattern change. Note: "-" and "Unknown" are NOT
# folded together — those are distinct Arkieva states and a move between them
# is a genuine change.
_STF_ATTR_BLANKS = {"", "nan", "none", "null", "na", "n/a"}


def _stf_norm_attr(series: pd.Series) -> pd.Series:
    """Normalise an attribute column for robust change comparison: cast to
    string, strip surrounding whitespace, and map blank/placeholder tokens to
    a single canonical empty string."""
    s = series.astype(str).str.strip()
    return s.mask(s.str.lower().isin(_STF_ATTR_BLANKS), "")


def _stf_grain_classified(sub: pd.DataFrame) -> pd.DataFrame:
    """Classify each granular row (Key + current & prior status/pattern) into
    a change bucket. ``sub`` must already be filtered to the STF rows of
    interest. Returns one row per granular combination with the current/prior
    status & pattern and the change ``Bucket``.

    Operating at this grain (not Key level) is essential: a Key legitimately
    carries multiple statuses, and each must be classified on its own so the
    per-bucket deltas are correct and reconcile to the net change.
    """
    has_lag_status = ACTIVE_STATUS_LAG_COL in sub.columns
    has_lag_pattern = PATTERN_LAG_COL in sub.columns

    grp_cols = ["Key", ACTIVE_STATUS_COL, PATTERN_COL]
    if has_lag_status:
        grp_cols.append(ACTIVE_STATUS_LAG_COL)
    if has_lag_pattern:
        grp_cols.append(PATTERN_LAG_COL)
    attrs = sub[grp_cols].drop_duplicates().reset_index(drop=True)

    cur_status = attrs[ACTIVE_STATUS_COL].astype(str)
    lag_status = (attrs[ACTIVE_STATUS_LAG_COL].astype(str)
                  if has_lag_status else cur_status)
    cur_pat = attrs[PATTERN_COL].astype(str)
    lag_pat = (attrs[PATTERN_LAG_COL].astype(str)
               if has_lag_pattern else cur_pat)

    # Compare on normalised values so blank/placeholder tokens don't create
    # spurious changes, while keeping the display values as-is.
    status_changed = (_stf_norm_attr(cur_status).values
                      != _stf_norm_attr(lag_status).values)
    pattern_changed = (_stf_norm_attr(cur_pat).values
                       != _stf_norm_attr(lag_pat).values)
    bucket = np.where(
        status_changed & pattern_changed, STF_BUCKET_BOTH,
        np.where(status_changed, STF_BUCKET_STATUS,
                 np.where(pattern_changed, STF_BUCKET_PATTERN,
                          STF_BUCKET_NONE)))

    out = pd.DataFrame({
        "Key": attrs["Key"].values,
        "CurStatus": cur_status.values, "LagStatus": lag_status.values,
        "CurPattern": cur_pat.values, "LagPattern": lag_pat.values,
        "Bucket": bucket,
    })
    return out


def _stf_grain_delta(sub: pd.DataFrame) -> pd.DataFrame:
    """Per granular-group net STF delta over the horizon (Σ current − Σ prior),
    keyed by Key + current & prior status/pattern, joined to the change
    bucket. ``sub`` is already restricted to the STF rows of the horizon."""
    has_lag_status = ACTIVE_STATUS_LAG_COL in sub.columns
    has_lag_pattern = PATTERN_LAG_COL in sub.columns
    grp_cols = ["Key", ACTIVE_STATUS_COL, PATTERN_COL]
    if has_lag_status:
        grp_cols.append(ACTIVE_STATUS_LAG_COL)
    if has_lag_pattern:
        grp_cols.append(PATTERN_LAG_COL)

    signed = sub.assign(
        _signed=np.where(sub["Data"].values == STF_CURRENT_LABEL,
                         sub["Value"].values, -sub["Value"].values))
    gd = signed.groupby(grp_cols, as_index=False)["_signed"].sum() \
        .rename(columns={"_signed": "Delta"})

    cur_status = gd[ACTIVE_STATUS_COL].astype(str)
    lag_status = (gd[ACTIVE_STATUS_LAG_COL].astype(str)
                  if has_lag_status else cur_status)
    cur_pat = gd[PATTERN_COL].astype(str)
    lag_pat = (gd[PATTERN_LAG_COL].astype(str)
               if has_lag_pattern else cur_pat)
    status_changed = (_stf_norm_attr(cur_status).values
                      != _stf_norm_attr(lag_status).values)
    pattern_changed = (_stf_norm_attr(cur_pat).values
                       != _stf_norm_attr(lag_pat).values)
    gd["Bucket"] = np.where(
        status_changed & pattern_changed, STF_BUCKET_BOTH,
        np.where(status_changed, STF_BUCKET_STATUS,
                 np.where(pattern_changed, STF_BUCKET_PATTERN,
                          STF_BUCKET_NONE)))
    # Explicit per-unique-key change flags (Yes/No), as required for the
    # material change-split visual and downstream classification.
    gd["Status Change"] = np.where(status_changed, "Yes", "No")
    gd["Pattern Change"] = np.where(pattern_changed, "Yes", "No")
    gd["Status and Pattern Change"] = np.where(
        status_changed & pattern_changed, "Yes", "No")
    gd["CurStatus"] = cur_status.values
    gd["LagStatus"] = lag_status.values
    gd["CurPattern"] = cur_pat.values
    gd["LagPattern"] = lag_pat.values
    return gd


@st.cache_data(show_spinner=False)
def compute_stf_material_changes(
    filtered_df: pd.DataFrame,
    horizon_months: Tuple[pd.Timestamp, ...],
    top_n: int = 25,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Top-N Materials by absolute STF deviation, split by change bucket.

    The classification grain — the **unique key** that tracks a forecast line
    across cycles — is Business Line + Material code + Ship To Sub Region +
    Arkieva Active Status + Arkieva Pattern (this is the app's ``Key`` plus
    status + pattern, which is collision-free). For each unique key the net
    STF deviation over the horizon (Σ current − Σ prior) is computed and the
    key is flagged/bucketed via ``Status Change`` / ``Pattern Change`` /
    ``Status and Pattern Change`` → No Changes / Status Change / Pattern
    Change / Status & Pattern Change.

    Those per-key deviations are then rolled up to **Material × Bucket** so
    each Material can be drawn as a stacked bar of its change split.

    Returns ``(long_df, wide_df)``:
      * ``long_df`` — [Material, Bucket, Delta, AbsDelta] for the top-N
        materials (tidy, for a stacked bar).
      * ``wide_df`` — one row per Material with a column per bucket's signed
        delta, the total net delta, total absolute deviation, and each
        bucket's **% share of the material's gross (absolute) change**.
    Materials are ranked by total absolute deviation (desc).
    """
    long_cols = ["Material", "Bucket", "Delta", "AbsDelta"]
    if filtered_df.empty or not horizon_months:
        return pd.DataFrame(columns=long_cols), pd.DataFrame()

    hz = list(horizon_months)
    sub = filtered_df[
        filtered_df["Data"].isin([STF_CURRENT_LABEL, STF_LAG1_LABEL])
        & filtered_df["Date"].isin(hz)
    ]
    if sub.empty:
        return pd.DataFrame(columns=long_cols), pd.DataFrame()

    # Per unique-key delta + bucket (grain already includes status+pattern).
    gd = _stf_grain_delta(sub)
    # Attach the Material label for each Key (one Material per Key).
    mat_map = (sub.drop_duplicates(subset=["Key"])
               .set_index("Key")["Material"])
    gd["Material"] = gd["Key"].map(mat_map)

    # Roll up to Material × Bucket.
    mb = (gd.groupby(["Material", "Bucket"], as_index=False)["Delta"].sum())
    mb["AbsDelta"] = mb["Delta"].abs()

    # Rank materials by total absolute deviation and keep the top N.
    mat_abs = (mb.groupby("Material")["AbsDelta"].sum()
               .sort_values(ascending=False))
    top_materials = list(mat_abs.head(top_n).index)
    long_df = mb[mb["Material"].isin(top_materials)].copy()
    # Order the Material categories by rank for the chart.
    long_df["Material"] = pd.Categorical(long_df["Material"],
                                         categories=top_materials, ordered=True)
    long_df = long_df.sort_values("Material").reset_index(drop=True)

    # Wide view: one row per material, a signed-delta column per bucket.
    wide = (long_df.pivot_table(index="Material", columns="Bucket",
                                values="Delta", aggfunc="sum", fill_value=0.0,
                                observed=False))
    for b in STF_BUCKET_ORDER:
        if b not in wide.columns:
            wide[b] = 0.0
    wide = wide[[b for b in STF_BUCKET_ORDER]]
    wide["Net delta"] = wide.sum(axis=1)
    gross = long_df.groupby("Material", observed=False)["AbsDelta"].sum()
    wide["Total abs deviation"] = gross
    for b in STF_BUCKET_ORDER:
        wide[f"{b} %"] = np.where(
            wide["Total abs deviation"] != 0,
            wide[b].abs() / wide["Total abs deviation"] * 100.0, 0.0)
    wide = wide.reindex(top_materials).reset_index()
    return long_df, wide


@st.cache_data(show_spinner=False)
def compute_stf_drivers(
    filtered_df: pd.DataFrame,
    horizon_months: Tuple[pd.Timestamp, ...],
    keys: Optional[Tuple[str, ...]] = None,
) -> pd.DataFrame:
    """Decompose the net STF change (current − prior committed) over the
    horizon into **change buckets** based on whether each granular unit's
    Arkieva Active Status and/or Arkieva Pattern changed vs the prior cycle:
    No Changes, Status Change, Pattern Change, Status & Pattern Change.

    Classification is at the granular grain (Key + status + pattern), because a
    single Key can carry several statuses; each granular unit's net delta
    (Σ current − Σ prior over the horizon) is assigned to exactly one bucket,
    so the bucket sums reconcile exactly to the overall net change. When
    ``keys`` is given the decomposition is restricted to those Keys (drill-
    down); otherwise it rolls up across ``filtered_df``.

    Returns [Driver, Delta] ordered per STF_BUCKET_ORDER (empty buckets are
    skipped).
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

    gd = _stf_grain_delta(sub)
    bucket_delta = gd.groupby("Bucket", as_index=False)["Delta"].sum()
    by_bucket = dict(zip(bucket_delta["Bucket"], bucket_delta["Delta"]))

    rows = [{"Driver": b, "Delta": float(by_bucket[b])}
            for b in STF_BUCKET_ORDER if b in by_bucket]
    return pd.DataFrame(rows, columns=cols)


@st.cache_data(show_spinner=False)
def compute_stf_driver_breakdown(
    filtered_df: pd.DataFrame,
    horizon_months: Tuple[pd.Timestamp, ...],
    bucket: str,
    keys: Optional[Tuple[str, ...]] = None,
    axis: Optional[str] = None,
) -> pd.DataFrame:
    """Drill-down for a change bucket: what specific transitions drive it.

    * For **Status Change** → group by the status transition
      ``LagStatus → CurStatus``.
    * For **Pattern Change** → group by the pattern transition
      ``LagPattern → CurPattern``.
    * For **Status & Pattern Change** → group by ``axis`` (defaults to status;
      pass ``axis='pattern'`` to view the pattern transitions instead).

    Returns [Transition, Delta, Keys] sorted by |Delta| desc, so a Demand
    Planner sees which from→to moves account for the bucket's STF change.
    """
    cols = ["Transition", "Delta", "Keys"]
    if filtered_df.empty or not horizon_months or not bucket:
        return pd.DataFrame(columns=cols)

    if axis is None:
        axis = "pattern" if bucket == STF_BUCKET_PATTERN else "status"

    hz = list(horizon_months)
    sub = filtered_df[
        filtered_df["Data"].isin([STF_CURRENT_LABEL, STF_LAG1_LABEL])
        & filtered_df["Date"].isin(hz)
    ]
    if keys:
        sub = sub[sub["Key"].isin(set(keys))]
    if sub.empty:
        return pd.DataFrame(columns=cols)

    # Granular-grain deltas already carry the change bucket and transitions.
    gd = _stf_grain_delta(sub)
    gd = gd[gd["Bucket"] == bucket]
    if gd.empty:
        return pd.DataFrame(columns=cols)

    if axis == "pattern":
        gd = gd.assign(Transition=gd["LagPattern"] + " → " + gd["CurPattern"])
    else:
        gd = gd.assign(Transition=gd["LagStatus"] + " → " + gd["CurStatus"])

    grp = (gd.groupby("Transition", as_index=False)
           .agg(Delta=("Delta", "sum"), Keys=("Key", "nunique")))
    grp["_abs"] = grp["Delta"].abs()
    grp = grp.sort_values("_abs", ascending=False).drop(columns="_abs")
    return grp[cols].reset_index(drop=True)


# ---------------------------------------------------------------------------
# Reset on file removal
# ---------------------------------------------------------------------------
def reset_app_state() -> None:
    """Clear THIS session's filter state and date range on file change.

    Only per-session widget keys are cleared (filter selections, date ranges,
    and the per-prefix init flags). We intentionally do NOT call
    ``load_excel.clear()``: that cache is global and keyed by file content, so
    clearing it would needlessly evict other concurrent users' data. Different
    files already map to different cache entries.

    Uses ``pop(k, None)`` rather than ``del`` so a key Streamlit has already
    reaped can't raise ``KeyError``. Auth keys (``authenticated``,
    ``auth_user``) and ``current_file_id`` are never touched here.
    """
    protected = {"authenticated", "auth_user", "current_file_id", "login_time"}
    keys_to_clear = [
        k for k in list(st.session_state.keys())
        if k not in protected and ("::" in k or k.startswith("flt_"))
    ]
    for k in keys_to_clear:
        st.session_state.pop(k, None)


# ---------------------------------------------------------------------------
# UI – Tab: Anomaly Summary
# ---------------------------------------------------------------------------
@st.fragment
def render_anomaly_tab(long_df: pd.DataFrame) -> None:
    st.markdown(
        '<div class="fva-tab-intro">'
        'Keys whose forecast breaks from their sales history, ranked by severity within each Ship To Sub Region. Pick a Key further down to see the root cause and suggested next steps.'
        "</div>", unsafe_allow_html=True)

    # ---- Filters (same set as the Forecast vs Actuals tab) -----------------
    anom_filter_cols = [c for c in FILTER_COLUMNS if c != "Data"]
    selections = filter_panel(
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
        st.success("No anomalies detected for the current filter selection.")
        with st.expander("How anomalies are detected and scored"):
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
    st.markdown("### Reason-code breakdown")
    breakdown = (summary["Primary reason"].value_counts()
                 .rename_axis("Reason").reset_index(name="Keys"))
    bcol, lcol = st.columns([1.3, 1])
    with bcol:
        breakdown = breakdown.sort_values("Keys")
        fig = go.Figure(go.Bar(
            x=breakdown["Keys"], y=breakdown["Reason"], orientation="h",
            marker=dict(color=CATEGORICAL[0], cornerradius=3),
            text=breakdown["Keys"], textposition="outside", cliponaxis=False,
            textfont=dict(color=TEXT_PRIMARY, size=12),
            hovertemplate="%{y}: %{x} Key(s)<extra></extra>",
        ))
        dark_layout(
            fig, title="Primary reason code by Key count",
            xaxis_title="Keys (by primary reason)", yaxis_title=None,
            height=120 + 46 * len(breakdown), showlegend=False,
            hovermode="closest", margin=dict(t=16, l=8, r=40, b=40),
        )
        fig.update_xaxes(showgrid=True, gridcolor=INK_GRID, showline=False,
                         showspikes=False)
        fig.update_yaxes(showgrid=False, ticksuffix="  ", automargin=True,
                         showspikes=False, tickfont=dict(color=TEXT_SECONDARY))
        show_chart(fig)
    with lcol:
        st.markdown("**Highest severity per Ship To Sub Region**")
        region_sev = (summary.groupby("Ship To Sub Region")
                      .agg(Severity=("Severity", "max"), Keys=("Key", "nunique"))
                      .sort_values("Severity", ascending=False).reset_index())
        st.dataframe(
            region_sev, width="stretch", hide_index=True,
            column_config={
                "Severity": st.column_config.ProgressColumn(
                    "Max severity", min_value=0,
                    max_value=float(max(100.0, region_sev["Severity"].max())),
                    format="%.0f"),
                "Keys": st.column_config.NumberColumn("Anomalous Keys"),
            },
        )

    # ---- Per-region table ---------------------------------------------------
    st.markdown(f"### Top {top_n} anomalies per Ship To Sub Region")
    regions = sorted(summary["Ship To Sub Region"].unique().tolist())
    region_pick = sticky_selectbox(
        "Ship To Sub Region", ["(All regions)"] + regions,
        key="anom::region_pick",
    )
    table = summary if region_pick == "(All regions)" \
        else summary[summary["Ship To Sub Region"] == region_pick]

    display_cols = ["Ship To Sub Region", "Key", "Primary reason",
                    "All reasons", "Severity", "Detail", "IQR outliers",
                    "Hist %/yr", "Fcst %/yr", "Seasonal corr"]
    st.dataframe(
        table[display_cols], width="stretch", hide_index=True, height=400,
        column_config={
            "Severity": st.column_config.ProgressColumn(
                "Severity", min_value=0,
                max_value=float(max(100.0, summary["Severity"].max())),
                format="%.0f"),
            "Key": st.column_config.TextColumn("Key", width="medium"),
            "All reasons": st.column_config.TextColumn("All reasons", width="medium"),
            "Detail": st.column_config.TextColumn("Detail", width="large"),
            "IQR outliers": st.column_config.NumberColumn("IQR outliers", format="%d"),
            "Hist %/yr": st.column_config.NumberColumn("History trend", format="%+.1f%%",
                                                       help="Fitted history trend, % per year"),
            "Fcst %/yr": st.column_config.NumberColumn("Forecast trend", format="%+.1f%%",
                                                       help="Fitted forecast trend, % per year"),
            "Seasonal corr": st.column_config.NumberColumn(
                "Seasonal match", format="%+.2f",
                help="Correlation of forecast vs historical seasonal shape (1 = same shape)"),
        },
    )

    csv = table[display_cols].to_csv(index=False)
    st.download_button(
        "Download anomaly summary (CSV)", data=csv,
        file_name="anomaly_summary.csv", mime="text/csv",
        key="anom::dl",
    )

    # ---- Drill-down ---------------------------------------------------------
    st.markdown("### Root cause for a Key")
    st.caption("Keys are listed from most to least severe.")
    key_choices = table.sort_values("Severity", ascending=False)["Key"].tolist()
    if not key_choices:
        return
    chosen = sticky_selectbox("Key", key_choices, key="anom::drill_key")

    _render_anomaly_drilldown(chosen, filtered)

    with st.expander("How anomalies are detected and scored"):
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

    # Severity + reason codes
    chips = "".join(
        f'<span class="fva-chip">{REASON_CODES[c]["label"]}</span>'
        for c in profile["all_codes"]
    )
    st.markdown(
        f'<div class="fva-keyline"><span class="fva-key">{key}</span>'
        f'<span class="fva-sev">Severity {profile["severity"]:.0f}</span>'
        f'{chips}</div>',
        unsafe_allow_html=True,
    )

    # Mini chart: history & forecast with trends and outliers
    sales_grp = profile["sales_grp"]
    fcst_grp = profile["fcst_grp"]
    iqr_res = profile["metrics"].get("iqr_result")

    fig = go.Figure()
    x_max = max([d.max() for d in (sales_grp["Date"], fcst_grp["Date"])
                 if len(d)] or [boundary])
    if boundary is not None:
        add_cycle_marks(fig, boundary, x_max)
    fig.add_trace(time_series_trace(
        sales_grp["Date"], sales_grp["Value"], SALES_HISTORY_LABEL,
        SERIES_STYLE[SALES_HISTORY_LABEL], name="Sales history",
        legend_group=False))
    if not fcst_grp.empty:
        fig.add_trace(time_series_trace(
            fcst_grp["Date"], mask_trailing_zeros(
                fcst_grp.set_index("Date")["Value"]).values,
            STAT_FORECAST_LABEL, SERIES_STYLE[STAT_FORECAST_LABEL],
            name="Statistical forecast", legend_group=False))
    ht = profile["metrics"].get("hist_trend")
    ft = profile["metrics"].get("fcst_trend")
    if ht is not None:
        tr = trend_trace(ht, "History trend", HISTORY_TREND_COLOR)
        tr.update(legendgroup=None, legendgrouptitle_text=None)
        fig.add_trace(tr)
    if ft is not None:
        tr = trend_trace(ft, "Forecast trend", FORECAST_TREND_COLOR)
        tr.update(legendgroup=None, legendgrouptitle_text=None)
        fig.add_trace(tr)
    if iqr_res is not None:
        od = iqr_res[iqr_res["IsOutlier"]]
        if not od.empty:
            fig.add_trace(go.Scatter(
                x=od["Date"], y=od["Value"], mode="markers", name="IQR outlier",
                marker=dict(size=13, symbol="circle-open", color=TEXT_PRIMARY,
                            line=dict(width=2)),
                hovertemplate="%{y:,.0f} kg (outlier)",
            ))
    dark_layout(
        fig, title=f"{key} history forecast trends",
        xaxis_title=None, yaxis_title="Demand (kg)", height=400,
        margin=dict(t=56, l=64, r=24, b=40),
    )
    fig.update_yaxes(tickformat="~s")
    show_chart(fig)

    # Supporting metrics
    mc1, mc2, mc3 = st.columns(3)
    h_pct = profile["metrics"].get("hist_pct_per_year")
    f_pct = profile["metrics"].get("fcst_pct_per_year")
    corr = profile["metrics"].get("seasonal_corr")
    mc1.metric("History trend", "–" if h_pct is None else f"{h_pct:+.1f}% / yr")
    mc2.metric("Forecast trend", "–" if f_pct is None else f"{f_pct:+.1f}% / yr")
    mc3.metric("Seasonal match", "–" if corr is None else f"{corr:+.2f}",
               help="Correlation between the forecast's and history's monthly "
                    "seasonal shape; 1 means the same shape.")

    # Root cause + next steps for each applicable reason code
    st.markdown("#### Root cause and next steps")
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
@st.fragment
def render_stf_variation_tab(long_df: pd.DataFrame,
                             all_months: List[pd.Timestamp]) -> None:
    st.markdown(
        '<div class="fva-tab-intro">'
        "How this cycle's committed statistical forecast moved against the prior cycle's, per Key. Variance = (current − prior) ÷ prior, summed over the horizon you choose. The horizon starts at M4, the fourth month after the last actual month."
        "</div>", unsafe_allow_html=True)

    # ---- Filters (same as Forecast vs Actuals, minus Data; NO status default)
    stf_filter_cols = [c for c in FILTER_COLUMNS if c != "Data"]
    selections = filter_panel(long_df, stf_filter_cols, key_prefix="stf")

    filtered = apply_filters(long_df, selections)
    if filtered.empty:
        st.warning("No data matches the current filter combination.")
        return

    # Need both current STF and Lag 1 to compute variance.
    if STF_LAG1_LABEL not in long_df["Data"].unique():
        st.info("This file has no **Statistical Forecast Committed** prior-"
                "cycle line (e.g. *… Lag 1/Lag 2 (kg)*), so STF variation "
                "cannot be computed. Please upload the raw Arkieva export that "
                "includes the prior-cycle committed series.")
        return

    boundary = history_forecast_boundary(filtered)
    m4 = stf_m4_month(filtered)
    if boundary is None or m4 is None:
        st.info("Couldn't determine the forecast horizon (no Sales History "
                "for the current filter selection).")
        return

    fcst_months = stf_forecast_months(all_months, boundary)
    if not fcst_months:
        st.info("No forecast months are available after the last active "
                "Sales-History month for this selection.")
        return

    # Auto-detected default start of the forecasting horizon = M4 (falls back
    # to the first forecast month if M4 is beyond the available range).
    default_start = m4 if m4 in fcst_months else fcst_months[0]
    default_end = fcst_months[-1]
    next_fys = stf_next_fiscal_years(boundary, n=2)

    # ---- Horizon controls: fiscal-year presets + custom month range --------
    c1, c2 = st.columns([1.1, 1.9])
    with c1:
        fy_labels = [f"FY {fy} (Jan–Dec)" for fy in next_fys]
        mode_options = ["Custom range"] + fy_labels
        mode = sticky_radio(
            "Forecast horizon", mode_options, key="stf::mode",
            help="Pick a full fiscal year (Jan–Dec) or set a custom start/end "
                 "month with the slider. The horizon start is auto-detected "
                 "at M4 (last active Sales-History month + 4).",
        )

    # Resolve the horizon months from the chosen mode.
    fy_selected = None
    if mode != "Custom range":
        fy_selected = next_fys[fy_labels.index(mode)]
        horizon_months = stf_fiscal_year_months(all_months, fy_selected)

    with c2:
        if len(fcst_months) >= 2:
            sel_start, sel_end = sticky_range_select_slider(
                "Forecasting horizon months (start → end)",
                options=fcst_months,
                key="stf::range",
                default=(default_start, default_end),
                format_func=lambda d: pd.Timestamp(d).strftime("%b %Y"),
                disabled=(mode != "Custom range"),
                help="Auto-detected start (M4) through the last forecast "
                     "month. Drag either end to tweak the horizon. Disabled "
                     "when a fiscal-year preset is selected.",
            )
        else:
            sel_start = sel_end = fcst_months[0]
            st.caption("Only one forecast month is available.")

        # Threshold sits directly under the horizon range, in a narrower
        # column so it doesn't span the full width. Capped at 200%.
        tcol, _tspacer = st.columns([2, 1])
        with tcol:
            threshold_pct = st.slider(
                "Absolute STF variation threshold", min_value=1,
                max_value=200, value=5, step=1, format="%d%%",
                key="stf::threshold",
                help="Keys with Absolute STF Variance above this are flagged "
                     "as exceptions. Default 5%, max 200%.",
            )
    threshold = threshold_pct / 100.0

    if mode == "Custom range":
        horizon_months = stf_months_between(all_months, sel_start, sel_end)

    if not horizon_months:
        if fy_selected is not None:
            st.info(f"No forecast months are available in FY {fy_selected} "
                    "for this selection.")
        else:
            st.info("No forecast months fall in the selected range.")
        return

    horizon_note = (f"fiscal year {fy_selected}" if fy_selected is not None
                    else "custom range")
    st.caption(
        f"Horizon **{horizon_months[0].strftime('%b %Y')} to "
        f"{horizon_months[-1].strftime('%b %Y')}** ({len(horizon_months)} "
        f"months, {horizon_note}; M4 is {m4.strftime('%b %Y')}). Keys above "
        f"**{threshold_pct}%** absolute variance count as exceptions."
    )

    # ---- Per-Key variance ---------------------------------------------------
    var_df = compute_stf_variance(filtered, tuple(horizon_months))
    if var_df.empty:
        st.info("No Keys have a prior-cycle forecast in this horizon, so "
                "there is nothing to compare against. Widen the horizon or "
                "the filters.")
        return

    n_exceptions = int((var_df["Absolute STF Variance"] > threshold).sum())
    net_cur = var_df["CurrentSTF"].sum()
    net_lag = var_df["Lag1STF"].sum()
    rollup_var = ((net_cur - net_lag) / net_lag) if net_lag else np.nan

    m1, m2, m3, m4c = st.columns(4)
    m1.metric("Keys compared", f"{len(var_df):,}",
              help=f"Keys with at least {STF_MIN_BASE_KG:,.0f} kg of prior-cycle "
                   "forecast in the horizon (a smaller base gives no "
                   "meaningful % variance).")
    m2.metric(f"Exceptions above {threshold_pct}%", f"{n_exceptions:,}")
    m3.metric("Total variance",
              "–" if pd.isna(rollup_var) else f"{rollup_var*100:+.1f}%",
              help="(Σ current − Σ prior) ÷ Σ prior across the Keys compared.")
    net = net_cur - net_lag
    m4c.metric("Net change", ("+" if net > 0 else "") + format_kg(net))

    # ---- History (36m back) vs forecast vintages (24m forward) -------------
    st.markdown("### This cycle's forecast against the prior cycle")
    st.caption(
        "The last 36 months of sales history and the next 24 months of the "
        "current and prior-cycle committed statistical forecast, summed over "
        "the filters. Where the two forecast lines part is where this "
        "cycle's forecast changed."
    )
    boundary = history_forecast_boundary(filtered)
    if boundary is None:
        st.info("No Sales History for this selection, so the comparison "
                "window can't be anchored.")
    else:
        hist_start = (pd.Timestamp(boundary) - pd.DateOffset(months=35)).replace(day=1)
        fcst_start = (pd.Timestamp(boundary) + pd.DateOffset(months=1)).replace(day=1)
        fcst_end = (pd.Timestamp(boundary) + pd.DateOffset(months=24)).replace(day=1)

        # (Data label, display name, window start, window end)
        cmp_series = [
            (HISTORY_FOR_FORECAST_LABEL, "History for forecast", hist_start, boundary),
            (SALES_HISTORY_LABEL, "Sales history", hist_start, boundary),
            (STF_CURRENT_LABEL, "Statistical forecast, this cycle", fcst_start, fcst_end),
            (STF_LAG1_LABEL, "Statistical forecast, prior cycle", fcst_start, fcst_end),
        ]
        cmp_agg = (filtered[filtered["Data"].isin([c[0] for c in cmp_series])]
                   .dropna(subset=["Value"])
                   .groupby(["Date", "Data"], as_index=False)["Value"].sum())
        fig_cmp = go.Figure()
        add_cycle_marks(fig_cmp, boundary, fcst_end)
        n_traces = 0
        for label, name, w_start, w_end in cmp_series:
            ser = cmp_agg[(cmp_agg["Data"] == label)
                          & (cmp_agg["Date"] >= w_start)
                          & (cmp_agg["Date"] <= w_end)].set_index("Date")["Value"]
            if ser.empty:
                continue
            style = SERIES_STYLE[STAT_FORECAST_LABEL if label == STF_CURRENT_LABEL
                                 else label]
            if style["category"] == "Forecast":
                ser = mask_trailing_zeros(ser)
            fig_cmp.add_trace(time_series_trace(ser.index, ser.values, label,
                                                style, name=name))
            n_traces += 1
        if n_traces == 0:
            st.info("No data in the comparison windows for this selection.")
        else:
            dark_layout(
                fig_cmp, title="Sales history and forecast vintages",
                xaxis_title=None, yaxis_title="Demand (kg)", height=430,
                margin=dict(t=72, l=64, r=24, b=40),
            )
            fig_cmp.update_yaxes(tickformat="~s")
            show_chart(fig_cmp)

    # ---- Top 25 exception Keys ---------------------------------------------
    st.markdown("### Top 25 Keys by absolute variance")
    top25 = var_df.head(25)
    if n_exceptions == 0:
        st.success(f"No Keys exceed the {threshold_pct}% threshold in this "
                   "horizon — the table below shows the largest movers.")
    else:
        st.caption(
            f"{n_exceptions:,} Keys exceed the {threshold_pct}% threshold; "
            "variances above it are highlighted. Sorted largest first."
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
        return ["background-color: rgba(230, 103, 103, 0.22); color: #ffd6d6;"
                if float(v) > thr_display else "" for v in col]

    show_disp = show[disp_cols].rename(columns={
        "CurrentSTF": "Current STF (kg)", "Lag1STF": "Prior STF (kg)",
        "STF Variance": "Variance", "Absolute STF Variance": "Abs. variance",
        "Arkieva Active Status": "Active Status"})
    st.dataframe(
        show_disp.style.format({
            "Current STF (kg)": "{:,.0f}", "Prior STF (kg)": "{:,.0f}",
            "Variance": "{:+.1f}%", "Abs. variance": "{:.1f}%",
        }).apply(_thr_highlight, subset=["Abs. variance"]),
        width="stretch", hide_index=True, height=400,
    )
    csv = top25.to_csv(index=False)
    st.download_button("Download top variation Keys (CSV)", data=csv,
                       file_name="stf_exception_keys.csv", mime="text/csv",
                       key="stf::dl_keys")

    # ---- Month highlight for the top Keys -----------------------------------
    exceptions = var_df[var_df["Absolute STF Variance"] > threshold].head(25)
    month_pool = exceptions if not exceptions.empty else top25
    if not month_pool.empty:
        st.markdown("### Which months drive the variation")
        st.caption("Variance by month. Red is a higher forecast than the "
                   "prior cycle, blue lower; the colour scale stops at ±100%. "
                   "Empty cells had no prior-cycle forecast that month.")
        month_options = ["(Top 10 exception Keys)"] + month_pool["Key"].tolist()
        month_scope = sticky_selectbox("Month-variation scope", month_options,
                                       key="stf::month_key")
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
                colorscale=DIVERGING_BLUE_RED, zmin=-100, zmax=100, zmid=0,
                hovertemplate="%{y}<br>%{x}: %{z:+.1f}%<extra></extra>",
                colorbar=dict(ticksuffix="%", tickvals=[-100, -50, 0, 50, 100],
                              ticktext=["≤−100", "−50", "0", "+50", "≥+100"],
                              len=0.9),
            ))
            dark_layout(
                fig_hm, title="Month level variance", xaxis_title=None,
                yaxis_title=None, height=max(240, 30 * len(mv_keys) + 110),
                hovermode="closest", margin=dict(t=16, l=8, r=24, b=56),
            )
            fig_hm.update_xaxes(showspikes=False, showline=False, ticks="",
                                tickangle=-45)
            fig_hm.update_yaxes(autorange="reversed", showgrid=False,
                                automargin=True, showspikes=False)
            show_chart(fig_hm)

    # ---- Top 25 Materials by change split ----------------------------------
    st.markdown("### Top 25 Materials: what kind of change moved the forecast")
    st.caption(
        "Materials ranked by absolute forecast change (current − prior cycle) "
        "over the horizon. Each bar is split by whether the forecast line's "
        "Arkieva Active Status and/or Pattern changed since the prior cycle; "
        "grey is change with no status or pattern change."
    )
    mat_long, mat_wide = compute_stf_material_changes(
        filtered, tuple(horizon_months), top_n=25)

    if mat_long.empty:
        st.info("No material STF deviation available for this selection.")
    else:
        bucket_colors = CHANGE_BUCKET_COLORS
        # mat_wide is ranked largest first. Plotly draws the first category
        # at the BOTTOM, so the category array is reversed to put the largest
        # material at the top (no extra autorange reversal).
        mat_order = mat_wide["Material"].tolist()
        labels = {m: material_label(m) for m in mat_order}
        totals = mat_wide.set_index("Material")["Total abs deviation"]
        fig_mat = go.Figure()
        for bucket in STF_BUCKET_ORDER:
            seg = mat_long[mat_long["Bucket"] == bucket]
            if seg.empty:
                continue
            seg = seg.set_index("Material").reindex(mat_order).reset_index()
            fig_mat.add_trace(go.Bar(
                y=[labels[m] for m in seg["Material"]],
                x=seg["AbsDelta"].fillna(0),
                name=bucket, orientation="h",
                marker=dict(color=bucket_colors[bucket],
                            line=dict(color=INK_PLANE, width=1.5)),
                customdata=np.stack([seg["Delta"].fillna(0)], axis=-1),
                hovertemplate=(bucket + ": %{x:,.0f} kg"
                               " (net %{customdata[0]:+,.0f} kg)<extra></extra>"),
            ))
        # Direct label: each material's total at the end of its bar.
        fig_mat.add_trace(go.Scatter(
            y=[labels[m] for m in mat_order],
            x=[float(totals[m]) for m in mat_order],
            mode="text", text=[format_kg(totals[m]).replace(" kg", "")
                               for m in mat_order],
            textposition="middle right", textfont=dict(size=11, color=TEXT_SECONDARY),
            showlegend=False, hoverinfo="skip", cliponaxis=False,
        ))
        dark_layout(
            fig_mat, title="Top 25 materials by absolute STF change",
            xaxis_title="Absolute change in committed statistical forecast (kg)",
            yaxis_title=None, height=40 + 28 * len(mat_order) + 90,
            hovermode="y unified", margin=dict(t=56, l=8, r=56, b=48),
        )
        fig_mat.update_layout(barmode="stack", bargap=0.35)
        fig_mat.update_yaxes(categoryorder="array",
                             categoryarray=[labels[m] for m in reversed(mat_order)],
                             showgrid=False, automargin=True, showspikes=False,
                             tickfont=dict(color=TEXT_SECONDARY))
        fig_mat.update_xaxes(tickformat="~s", showgrid=True, gridcolor=INK_GRID,
                             showline=False, showspikes=False)
        show_chart(fig_mat)

        # ---- Table: absolute (kg) and % split per material -----------------
        st.markdown("#### Change split by material")
        disp = mat_wide.copy()
        kg_cols = list(STF_BUCKET_ORDER) + ["Net delta", "Total abs deviation"]
        disp[[c for c in kg_cols if c in disp.columns]] = (
            disp[[c for c in kg_cols if c in disp.columns]].round(0))
        disp.insert(0, "Material ", disp["Material"].map(material_label))
        disp = disp.drop(columns=["Material"])
        col_cfg = {"Material ": st.column_config.TextColumn("Material", width="large")}
        for b in STF_BUCKET_ORDER:
            col_cfg[b] = st.column_config.NumberColumn(f"{b} (kg)", format="localized")
            col_cfg[f"{b} %"] = st.column_config.NumberColumn(
                f"{b} %", format="%.1f%%",
                help="Share of the material's total absolute change")
        col_cfg["Net delta"] = st.column_config.NumberColumn(
            "Net change (kg)", format="localized")
        col_cfg["Total abs deviation"] = st.column_config.NumberColumn(
            "Absolute change (kg)", format="localized")
        order = (["Material "] + list(STF_BUCKET_ORDER)
                 + ["Net delta", "Total abs deviation"]
                 + [f"{b} %" for b in STF_BUCKET_ORDER])
        st.dataframe(disp[[c for c in order if c in disp.columns]],
                     column_config=col_cfg, width="stretch",
                     hide_index=True, height=420)
        st.download_button(
            "Download material change split (CSV)",
            data=mat_wide.to_csv(index=False),
            file_name="stf_material_change_split.csv", mime="text/csv",
            key="stf::dl_materials")


# ---------------------------------------------------------------------------
# UI – Tab 1: Forecast vs Actuals
# ---------------------------------------------------------------------------
@st.fragment
def render_dashboard_tab(long_df: pd.DataFrame, file_name: str) -> None:
    st.markdown(
        '<div class="fva-tab-intro">'
        'Sales history and history for forecast across the whole timeline; forecast lines from the last actual month onward. Use the legend to show or hide lines.'
        "</div>", unsafe_allow_html=True)

    # ---- Filter strip on top -----------------------------------------------
    selections = filter_panel(
        long_df, FILTER_COLUMNS, key_prefix="dash",
        default_selections={ACTIVE_STATUS_COL: DEFAULT_ACTIVE_STATUSES})

    min_d = long_df["Date"].min().to_pydatetime()
    max_d = long_df["Date"].max().to_pydatetime()
    range_col, _ = st.columns([2, 3])
    with range_col:
        st.slider(
            "Months shown",
            min_value=min_d, max_value=max_d, value=(min_d, max_d),
            format="MMM YYYY", key="dash::date_range",
        )

    d_start, d_end = st.session_state.get("dash::date_range", (min_d, max_d))
    d_start, d_end = pd.Timestamp(d_start), pd.Timestamp(d_end)

    filtered = apply_filters(long_df, selections)
    filtered = filtered[(filtered["Date"] >= d_start) & (filtered["Date"] <= d_end)]

    if filtered.empty:
        st.warning("No data matches the current filter combination.")
        return

    # ---- KPI cards ----------------------------------------------------------
    boundary = history_forecast_boundary(filtered)
    sales_last12 = fcst_next12 = None
    if boundary is not None:
        b0 = pd.Timestamp(boundary)
        sh = filtered[(filtered["Data"] == SALES_HISTORY_LABEL)
                      & (filtered["Date"] > b0 - pd.DateOffset(months=12))
                      & (filtered["Date"] <= b0)]
        sales_last12 = float(sh["Value"].sum(skipna=True))
        fc = filtered[(filtered["Data"] == STAT_FORECAST_LABEL)
                      & (filtered["Date"] > b0)
                      & (filtered["Date"] <= b0 + pd.DateOffset(months=12))]
        fcst_next12 = float(fc["Value"].sum(skipna=True))
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Keys", f"{filtered['Key'].nunique():,}")
    c2.metric("Materials", f"{filtered['Material'].nunique():,}")
    c3.metric("Sales history, last 12 months",
              "–" if sales_last12 is None else format_kg(sales_last12),
              help="Sum of Sales History (kg) over the 12 months up to the "
                   "last actual month.")
    if fcst_next12 is not None and sales_last12:
        delta = (fcst_next12 - sales_last12) / sales_last12 * 100
        c4.metric("Statistical forecast, next 12 months", format_kg(fcst_next12),
                  delta=f"{delta:+.1f}% vs last 12 months", delta_color="off",
                  help="Sum of the Statistical Forecast (kg) over the 12 "
                       "months after the last actual month.")
    else:
        c4.metric("Statistical forecast, next 12 months",
                  "–" if fcst_next12 is None else format_kg(fcst_next12))

    # ---- Trend toggle -------------------------------------------------------
    show_trend = st.toggle(
        "Show trend lines", value=True,
        key="dash::show_trend",
        help="Fits a straight-line trend separately to the history and "
             "forecast periods, to show whether the forecast follows the "
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
    if boundary is not None and len(pivot.index) > 0:
        add_cycle_marks(fig, boundary, pivot.index.max())

    # History series run across BOTH periods; forecast series (Statistical
    # Forecast Committed, its prior-cycle line, Final Demand Plan Lag 1) are
    # drawn only after the last actual month, so the history period stays
    # clean for reading actual sales, seasonality and outliers.
    ordered = [s for s in HISTORY_SERIES + FORECAST_SERIES if s in pivot.columns]
    for col in ordered:
        style = SERIES_STYLE.get(
            col, {"color": TEXT_MUTED, "dash": "solid", "category": "Other"}
        )
        y_vals = pivot[col]
        if style["category"] == "Forecast":
            if boundary is not None:
                y_vals = y_vals.where(pivot.index > boundary)
            y_vals = mask_trailing_zeros(y_vals)
        fig.add_trace(time_series_trace(pivot.index, y_vals, col, style))

    # ---- Trend lines -------------------------------------------------------
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
            fig.add_trace(trend_trace(hist_trend, "Sales history trend",
                                      HISTORY_TREND_COLOR))
        if fcst_trend is not None:
            fig.add_trace(trend_trace(fcst_trend, "Statistical forecast trend",
                                      FORECAST_TREND_COLOR))

    dark_layout(
        fig,
        title="Forecast vs actuals",
        xaxis_title=None, yaxis_title="Demand (kg)",
        height=540,
        legend=dict(groupclick="toggleitem"),
        margin=dict(t=88, l=64, r=24, b=40),
    )
    fig.update_layout(legend=dict(y=1.14), margin=dict(t=104))
    add_time_controls(fig)
    fig.update_yaxes(tickformat="~s")

    show_chart(fig)

    # ---- Trend interpretation ----------------------------------------------
    if show_trend:
        status, message = interpret_trend(hist_trend, fcst_trend)
        st.markdown("#### What the trend says")
        if status == "ok":
            st.success(message)
        elif status == "warn":
            st.warning(message)
        else:
            st.info(message)


# ---------------------------------------------------------------------------
# UI – Tab 2: Outlier Detection & Correction (Hampel filter)
# ---------------------------------------------------------------------------
@st.fragment
def render_outlier_tab(long_df: pd.DataFrame) -> None:
    st.markdown(
        '<div class="fva-tab-intro">'
        "Find and correct outliers in Sales History with the IQR or Sigma method. Each Key's chart shows the original and the cleansed history, with History For Forecast unchanged for comparison."
        "</div>", unsafe_allow_html=True)

    sales = long_df[long_df["Data"] == SALES_HISTORY_LABEL].copy()
    if sales.empty:
        st.warning(f"No rows with Data == '{SALES_HISTORY_LABEL}' were found.")
        return

    hist_ff = long_df[long_df["Data"] == HISTORY_FOR_FORECAST_LABEL].copy()

    # 'Data' filter is omitted: this view is intrinsically Sales History.
    outlier_filter_cols = [c for c in FILTER_COLUMNS if c != "Data"]

    selections = filter_panel(
        sales, outlier_filter_cols, key_prefix="outlier",
        default_selections={ACTIVE_STATUS_COL: DEFAULT_ACTIVE_STATUSES})

    sales_f = apply_filters(sales, selections)
    if sales_f.empty:
        st.warning("No Sales History rows match the current filter combination.")
        return

    # Detect on actual history only. Sales History is exported as zeros for
    # future months; including them dragged the IQR quartiles and the median
    # correction toward zero and flagged ordinary sales as outliers.
    last_actual = history_forecast_boundary(sales_f)
    if last_actual is not None:
        sales_f = sales_f[sales_f["Date"] <= last_actual]
        hist_ff = hist_ff[hist_ff["Date"] <= last_actual]
        st.caption(
            f"Detecting on {sales_f['Key'].nunique():,} Keys, using actual "
            f"history up to {pd.Timestamp(last_actual):%b %Y} (future months, "
            "exported as zero, are excluded)."
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
    m1.metric("Keys checked", f"{total_keys:,}")
    m2.metric("Keys with an outlier", f"{keys_with_outliers:,}")
    m3.metric("Months checked", f"{total_pts:,}")
    rate = (total_outliers / total_pts * 100) if total_pts else 0
    m4.metric("Outliers corrected", f"{total_outliers:,}",
              delta=f"{rate:.1f}% of months", delta_color="off")

    # ---- Per-Key chooser ----------------------------------------------------
    st.markdown("### Inspect a single Key")
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
        chosen_key = sticky_selectbox("Select a Key", key_options,
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
            "Download this Key's cleansed history (CSV)",
            data=cbuf.getvalue(),
            file_name=f"cleansed_{method.lower()}_{chosen_key}.csv",
            mime="text/csv",
            key="outlier::dl_key",
        )

    # ---- Full outlier listing -----------------------------------------------
    st.markdown("### All corrected outliers")
    outlier_df = all_results[all_results["IsOutlier"]].copy()
    if outlier_df.empty:
        st.success("No outliers detected with the current settings.")
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
        st.caption(f"**{len(outlier_df):,}** flagged outlier(s) in the current "
                   "selection. Use the filters above to narrow the list, or "
                   "download the full set as CSV below.")
        render_numeric_table(
            outlier_df[cols],
            {c: "{:,.2f}" for c in
             ["Original", "Corrected", "Center", "Lower", "Upper"]},
            na_rep="–",
            width="stretch", hide_index=True,
        )
        buf = io.StringIO()
        outlier_df[cols].to_csv(buf, index=False)
        st.download_button(
            "Download all corrected outliers (CSV)",
            data=buf.getvalue(),
            file_name=f"corrected_outliers_{method.lower()}.csv",
            mime="text/csv",
            key="outlier::dl_all",
        )

    # ---- Full cleansed-history export (all Keys) ----------------------------
    st.markdown("### Cleansed history export")
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
        "Download cleansed history for all Keys (CSV)",
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

    # Normal range between the method's lower and upper bounds (one legend item).
    bounds_df = key_df.dropna(subset=["Lower", "Upper"]).sort_values("Date")
    if not bounds_df.empty:
        fig.add_trace(go.Scatter(
            x=bounds_df["Date"], y=bounds_df["Upper"], mode="lines",
            line=dict(width=0), showlegend=False, legendgroup="band",
            hovertemplate="Upper bound %{y:,.0f} kg<extra></extra>",
        ))
        fig.add_trace(go.Scatter(
            x=bounds_df["Date"], y=bounds_df["Lower"], mode="lines",
            line=dict(width=0), fill="tonexty",
            fillcolor="rgba(167, 177, 189, 0.10)",
            name="Normal range", legendgroup="band",
            hovertemplate="Lower bound %{y:,.0f} kg<extra></extra>",
        ))

    if hff_df is not None and not hff_df.empty:
        fig.add_trace(time_series_trace(
            hff_df["Date"], hff_df["Value"], HISTORY_FOR_FORECAST_LABEL,
            SERIES_STYLE[HISTORY_FOR_FORECAST_LABEL],
            name="History for forecast (unchanged)", legend_group=False,
            line=dict(color=SERIES_STYLE[HISTORY_FOR_FORECAST_LABEL]["color"],
                      width=1.5), mode="lines"))
    fig.add_trace(time_series_trace(
        key_df["Date"], key_df["Value"], SALES_HISTORY_LABEL,
        SERIES_STYLE[SALES_HISTORY_LABEL], name="Sales history (original)",
        legend_group=False))
    fig.add_trace(time_series_trace(
        key_df["Date"], key_df["Filtered"], "Cleansed",
        {"color": CLEANSED_COLOR, "category": "History"},
        name="Cleansed history", legend_group=False,
        line=dict(color=CLEANSED_COLOR, width=2, dash="dot"),
        marker=dict(size=5, color=CLEANSED_COLOR, symbol="square")))

    out_df = key_df[key_df["IsOutlier"]]
    if not out_df.empty:
        fig.add_trace(go.Scatter(
            x=out_df["Date"], y=out_df["Value"], mode="markers",
            name="Outlier (corrected)",
            marker=dict(size=14, symbol="circle-open", color=TEXT_PRIMARY,
                        line=dict(width=2)),
            hovertemplate="%{y:,.0f} kg (outlier)",
        ))

    param_label = f"k {param}" if method == "IQR" else f"n {param}"
    dark_layout(
        fig, title=f"{key} {method} {param_label}",
        xaxis_title=None, yaxis_title="Sales (kg)", height=460,
        margin=dict(t=56, l=64, r=24, b=40),
    )
    fig.update_yaxes(tickformat="~s")
    show_chart(fig)

    n_out = int(key_df["IsOutlier"].sum())
    st.caption(f"{n_out} month(s) flagged and corrected for this Key."
               if n_out else "No months flagged for this Key.")


# ---------------------------------------------------------------------------
# UI – Tab 3: Year-over-Year Seasonality (Change 4)
# ---------------------------------------------------------------------------
@st.fragment
def render_seasonality_tab(long_df: pd.DataFrame) -> None:
    st.markdown(
        '<div class="fva-tab-intro">'
        "Each line is a seasonal index: 1.10 means that month runs about 10% above the period's average. Compare the forecast's monthly shape with the last three years of sales."
        "</div>", unsafe_allow_html=True)

    # ---- Filter strip (same six as the outlier tab) ------------------------
    season_filter_cols = [c for c in FILTER_COLUMNS if c != "Data"]
    selections = filter_panel(
        long_df, season_filter_cols, key_prefix="season",
        default_selections={ACTIVE_STATUS_COL: DEFAULT_ACTIVE_STATUSES})

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

    # Individual history years: thin tints of the Sales History blue, older
    # years fainter. The 3-year average is the solid blue line and the
    # forecast is the Statistical Forecast orange, as on every other tab.
    hist_color = SERIES_STYLE[SALES_HISTORY_LABEL]["color"]
    fcst_color = SERIES_STYLE[STAT_FORECAST_LABEL]["color"]
    years = [(yr, prof) for yr, prof in sorted(analysis["per_year"].items())
             if prof is not None]
    for i, (yr, prof) in enumerate(years):
        opacity = 0.35 + 0.25 * (i / max(1, len(years) - 1))
        fig.add_trace(go.Scatter(
            x=month_labels, y=prof.reindex(months).values,
            mode="lines", name=f"{yr}", opacity=opacity,
            line=dict(color=hist_color, width=1.25),
            legendgroup="History", legendgrouptitle_text="History",
            hovertemplate="%{y:.2f}",
        ))
    fig.add_trace(go.Scatter(
        x=month_labels, y=analysis["hist_profile"].reindex(months).values,
        mode="lines+markers", name="3-year average",
        line=dict(color=hist_color, width=2.5), marker=dict(size=6),
        legendgroup="History", legendgrouptitle_text="History",
        hovertemplate="%{y:.2f}",
    ))
    if analysis["fcst_profile"] is not None:
        fig.add_trace(go.Scatter(
            x=month_labels, y=analysis["fcst_profile"].reindex(months).values,
            mode="lines+markers", name="Statistical forecast",
            line=dict(color=fcst_color, width=2.5),
            marker=dict(size=7, symbol="diamond"),
            legendgroup="Forecast", legendgrouptitle_text="Forecast",
            hovertemplate="%{y:.2f}",
        ))
    fig.add_hline(y=1.0, line=dict(color="#3d4a5c", width=1),
                  annotation_text="Average month = 1.00",
                  annotation_position="top right",
                  annotation_font=dict(color=TEXT_MUTED, size=11))

    dark_layout(
        fig, title="Monthly seasonal index history vs forecast",
        xaxis_title=None, yaxis_title="Seasonal index",
        height=460, margin=dict(t=80, l=64, r=24, b=40),
    )
    fig.update_xaxes(showspikes=True)
    fig.update_yaxes(tickformat=".2f")
    show_chart(fig)

    # ---- Interpretation -----------------------------------------------------
    status, message = interpret_seasonality(analysis)
    st.markdown("#### What the seasonality says")
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
@st.fragment
def render_stat_adoption_tab(long_df: pd.DataFrame) -> None:
    st.markdown(
        '<div class="fva-tab-intro">'
        "Share of forecast volume running on the statistical forecast: combinations with Arkieva Review Req = No and status Active or Sparse, measured against each Ship To Sub Region's total forecast, per fiscal year (Jan–Dec)."
        "</div>", unsafe_allow_html=True)

    # ---- Filters (same set + Active/Sparse default) ------------------------
    # The Data filter is dropped (the tab is scoped to the Statistical
    # Forecast internally) and so is Arkieva Review Req — the on-stat
    # definition already applies Review Req = No, so the filter is
    # redundant here.
    adopt_filter_cols = [c for c in FILTER_COLUMNS
                         if c not in ("Data", REVIEW_REQ_COL)]
    selections = filter_panel(
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
        f"Fiscal years {active_year} to {years_covered[-1]}: {active_year} is "
        "the year of the last actual month, and earlier years are excluded."
    )

    m1, m2, m3 = st.columns(3)
    m1.metric(f"Adoption, {active_year} onward", f"{grand_pct:.1f}%")
    m2.metric("Forecast on statistical forecast", format_kg(grand_on))
    m3.metric("Total forecast", format_kg(grand_tot))

    # ---- Adoption % by fiscal year (line + bar combo) ----------------------
    st.markdown("### Adoption by fiscal year")
    year_df = overall.copy()
    if not year_df.empty:
        year_df["YearLabel"] = year_df["Year"].astype(int).astype(str)
        fig_year = go.Figure()
        fig_year.add_trace(go.Bar(
            x=year_df["YearLabel"], y=year_df["AdoptionPct"],
            name="Adoption", marker=dict(color=CATEGORICAL[0], cornerradius=3),
            text=[f"{v:.1f}%" for v in year_df["AdoptionPct"]],
            textposition="outside", cliponaxis=False,
            textfont=dict(color=TEXT_PRIMARY, size=13),
            hovertemplate="FY %{x}: %{y:.1f}%<extra></extra>",
        ))
        dark_layout(
            fig_year, title="Statistical forecast adoption by fiscal year",
            xaxis_title="Fiscal year (Jan–Dec)", yaxis_title="Adoption",
            height=340, showlegend=False, hovermode="closest",
            margin=dict(t=36, l=64, r=24, b=48), bargap=0.55,
        )
        fig_year.update_xaxes(showspikes=False, type="category")
        fig_year.update_yaxes(ticksuffix="%", rangemode="tozero")
        show_chart(fig_year)

    # ---- By Business Line and by Ship To Sub Region (grouped bars) ----------
    bcol, rcol = st.columns(2)

    with bcol:
        st.markdown("### By Business Line")
        bl_colors = entity_colors(long_df, "Business Line")
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
                    marker=dict(color=bl_colors.get(str(bl), TEXT_MUTED),
                                cornerradius=2),
                    hovertemplate="%{y:.1f}%",
                ))
            dark_layout(
                fig_bl, title="Adoption by Business Line",
                xaxis_title=None, yaxis_title="Adoption",
                height=380, barmode="group",
                margin=dict(t=72, l=56, r=16, b=40),
            )
            fig_bl.update_xaxes(type="category")
            fig_bl.update_yaxes(ticksuffix="%", rangemode="tozero")
            show_chart(fig_bl)

    with rcol:
        st.markdown("### By Ship To Sub Region")
        rg_colors = entity_colors(long_df, "Ship To Sub Region")
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
                    marker=dict(color=rg_colors.get(str(rg), TEXT_MUTED),
                                cornerradius=2),
                    hovertemplate="%{y:.1f}%",
                ))
            dark_layout(
                fig_rg, title="Adoption by Ship To Sub Region",
                xaxis_title=None, yaxis_title="Adoption",
                height=380, barmode="group",
                margin=dict(t=72, l=56, r=16, b=40),
            )
            fig_rg.update_xaxes(type="category")
            fig_rg.update_yaxes(ticksuffix="%", rangemode="tozero")
            show_chart(fig_rg)

    # ---- Heatmap: Business Line × Year ------------------------------------
    bl_df_h = compute_stat_adoption(filtered, group_col="Business Line",
                                    active_year=active_year)
    if not bl_df_h.empty and bl_df_h["Business Line"].nunique() > 1:
        st.markdown("### Adoption by Business Line and fiscal year")
        pivot = bl_df_h.pivot(index="Business Line", columns="Year",
                              values="AdoptionPct").sort_index()
        fig_hm = go.Figure(go.Heatmap(
            z=pivot.values,
            x=[str(int(c)) for c in pivot.columns],
            y=pivot.index.tolist(),
            colorscale=SEQUENTIAL_BLUE, zmin=0, zmax=100,
            text=[[f"{v:.0f}%" if pd.notna(v) else "" for v in row]
                  for row in pivot.values],
            texttemplate="%{text}", textfont=dict(size=13, color="#ffffff"),
            hovertemplate="%{y}, FY %{x}: %{z:.1f}%<extra></extra>",
            colorbar=dict(ticksuffix="%", len=0.9),
        ))
        dark_layout(
            fig_hm, title="Adoption heatmap", xaxis_title=None,
            yaxis_title=None, height=110 + 52 * len(pivot.index),
            hovermode="closest", margin=dict(t=16, l=8, r=24, b=40),
        )
        fig_hm.update_xaxes(showspikes=False, showline=False, ticks="",
                            type="category")
        fig_hm.update_yaxes(showgrid=False, automargin=True, showspikes=False)
        show_chart(fig_hm)

    # ---- Materials on Statistical Forecast (interactive table) -------------
    st.markdown("### Materials on the statistical forecast")
    st.caption(
        "Each Business Line + Material + Ship To Sub Region + Active Status "
        "combination on the statistical forecast, with its forecast volume per "
        "fiscal year and its share of its region's total forecast."
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
        fmt_map = {c: "{:,.0f}" for c in fmt_cols}
        if "% of region forecast" in mat_table.columns:
            fmt_map["% of region forecast"] = "{:.1f}%"
        render_numeric_table(mat_table, fmt_map, na_rep="–",
                             width="stretch", hide_index=True, height=420)
        csv = mat_table.to_csv(index=False)
        st.download_button(
            "Download materials on Statistical Forecast (CSV)",
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

    # ---- Header -----------------------------------------------------------------
    auth_user = st.session_state.get("auth_user", "")
    display_name = USER_DISPLAY_NAMES.get(auth_user, auth_user.capitalize())
    head_l, head_r = st.columns([6, 1], vertical_alignment="top")
    with head_l:
        st.markdown('<div class="fva-header-title">Forecast vs Actuals</div>',
                    unsafe_allow_html=True)
        sub_slot = st.empty()
        cycle_slot = st.empty()
    with head_r:
        if st.button("Log out", key="auth::logout", width="stretch"):
            logout()

    if "current_file_id" not in st.session_state:
        st.session_state["current_file_id"] = None

    has_file = st.session_state["current_file_id"] is not None
    file_label = ("Data file" if not has_file else
                  f"Data file: {st.session_state['current_file_id'].split('::')[0]}")
    with st.expander(file_label, expanded=not has_file):
        uploaded = st.file_uploader(
            "Upload the Arkieva Forecast vs Actuals export (.xlsx)",
            type=["xlsx"], accept_multiple_files=False,
            help="Upload the raw export as downloaded from Arkieva; the data "
                 "is read from the first sheet. Replace it here at any time.",
        )

    if uploaded is None:
        if st.session_state["current_file_id"] is not None:
            st.session_state["current_file_id"] = None
            reset_app_state()
        sub_slot.markdown(
            f'<div class="fva-header-sub">Signed in as {display_name}</div>',
            unsafe_allow_html=True)
        st.markdown(
            '<div class="fva-tab-intro" style="margin-top:1rem">'
            "Upload this month's Arkieva <b>Forecast vs Actuals</b> export to "
            "start the review. The dashboard then shows:</div>",
            unsafe_allow_html=True)
        st.markdown(
            "- **Anomaly summary** — Keys whose forecast breaks from history "
            "(trend, seasonality, outliers), ranked per Ship To Sub Region, "
            "with root cause and next steps.\n"
            "- **STF variation & exceptions** — where this cycle's committed "
            "statistical forecast moved against the prior cycle, and why.\n"
            "- **Forecast vs actuals** — history and forecast lines for any "
            "slice, with trend lines.\n"
            "- **Outliers** — IQR or Sigma detection on Sales History, with a "
            "cleansed history to export.\n"
            "- **Seasonality (YoY)** — the forecast's seasonal shape against "
            "the last three years.\n"
            "- **Stat forecast adoption** — share of forecast volume running "
            "on the statistical forecast, by year, Business Line and region."
        )
        st.stop()

    file_id = f"{uploaded.name}::{uploaded.size}"
    if st.session_state["current_file_id"] != file_id:
        reset_app_state()
        st.session_state["current_file_id"] = file_id
        st.rerun()  # redraw the header with the file collapsed and named

    file_bytes = uploaded.getvalue()
    try:
        with st.spinner("Reading the export…"):
            long_df, all_months = load_excel(file_bytes, uploaded.name)
    except ValueError as exc:
        st.error(str(exc))
        st.stop()
    except Exception as exc:  # noqa: BLE001
        st.error(f"The file couldn't be read: {exc}. Check that it is the raw "
                 "Arkieva Forecast vs Actuals export and upload it again.")
        st.stop()

    if long_df.empty:
        st.error("The uploaded file has no data rows. Export it again from "
                 "Arkieva and upload the new file.")
        st.stop()

    # ---- Planning-cycle readout (what every tab is anchored to) ---------------
    boundary = history_forecast_boundary(long_df)
    m4 = stf_m4_month(long_df)
    reporting = extract_month_from_filename(uploaded.name)
    sub_bits = [f"Signed in as {display_name}",
                f"{long_df['Key'].nunique():,} Keys"]
    if reporting:
        sub_bits.append(f"reporting month {reporting}")
    sub_slot.markdown(f'<div class="fva-header-sub">{" | ".join(sub_bits)}</div>',
                      unsafe_allow_html=True)
    cells = []
    if boundary is not None:
        cells.append(("Last actual month", pd.Timestamp(boundary).strftime("%b %Y"), ""))
    if m4 is not None:
        cells.append(("Variation horizon starts (M4)", pd.Timestamp(m4).strftime("%b %Y"), "fcst"))
    if all_months:
        cells.append(("Forecast runs to", pd.Timestamp(all_months[-1]).strftime("%b %Y"), "fcst"))
        cells.append(("History shown from", pd.Timestamp(all_months[0]).strftime("%b %Y"), ""))
    if cells:
        cycle_slot.markdown(
            '<div class="fva-cycle">' + "".join(
                f'<div class="{cls}"><span class="lbl">{lbl}</span>'
                f'<span class="val">{val}</span></div>'
                for lbl, val, cls in cells) + "</div>",
            unsafe_allow_html=True)

    tab0, tab_stf, tab1, tab2, tab3, tab4 = st.tabs([
        "Anomaly summary",
        "STF variation & exceptions",
        "Forecast vs actuals",
        "Outliers",
        "Seasonality (YoY)",
        "Stat forecast adoption",
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
