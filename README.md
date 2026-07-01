# Forecast vs Actuals — Interactive Dashboard

A production-grade, **dark-themed** Streamlit application that replicates
the Excel **"Forecast vs Actuals" pivot dashboard** and adds an
**Anomaly Summary** triage view, an **STF Variation & Exceptions** view, a
**Statistical Forecast Adoption %** view, interactive **outlier detection
& correction** (IQR or Sigma), **trend analysis**, and **year-over-year
seasonality** views — each with automatic, plain-language interpretations
of forecast quality and recommended next steps.

The app is designed for Demand Planners to use directly: **download the
raw "Forecast vs Actuals" export from Arkieva and upload it as-is** (the
required data is always on the first sheet). The app derives the Key,
reshapes the data, and renders every chart dynamically from the most
recently uploaded file — no manual preparation needed.

---

## Features

### Sign-in screen (authentication)
- The app is gated behind a **login screen**: nothing (including the file
  uploader) renders until a valid username and password are entered.
- Seven authorised users are configured. Usernames are case-insensitive;
  passwords are case-sensitive. Credentials are stored in the code as
  **SHA-256 hashes only — never plaintext** — and verified with a
  constant-time comparison (`hmac.compare_digest`).
- After signing in, the header shows who is signed in and provides a
  **🚪 Log out** button, which clears all session state (filters, data
  cache, upload) and returns to the sign-in screen.
- The raw password is purged from Streamlit session state immediately
  after a successful login.

> **Security note for public repositories:** Streamlit Community Cloud's
> free tier requires a public GitHub repo, which means the hashed
> credential list in `app.py` is publicly visible. SHA-256 hashes cannot
> be reversed directly, but short/patterned passwords are vulnerable to
> guessing. For stronger protection, move the hash table into **Streamlit
> secrets** (App → Settings → Secrets) and read it via `st.secrets`, or
> upgrade to a private repo on a paid workspace. Share the actual
> passwords with your users through a private channel — never commit
> plaintext passwords or document them in this README.

### Dark theme
- The entire UI uses a **dark theme** — background, filters, dropdown
  menus, the file uploader, buttons, tabs, metric cards and tables — with
  high-contrast light text chosen for legibility. The theme is set both
  via `.streamlit/config.toml` (so it applies from first paint) and via
  in-app CSS for the widgets the config alone doesn't fully cover. All
  Plotly charts use a matching dark canvas with light, legible labels.
- **Fullscreen charts stay dark**: every chart uses a solid dark canvas
  (equal to the page background), and the fullscreen overlay is styled
  dark via CSS, so expanding any chart with the *Fullscreen* button keeps
  the same dark theme instead of flashing a white background.

### Tab 1 — Anomaly Summary
The **first tab** — a triage view for Demand Planners.

- Uses the **same cascading filters** as the *Forecast vs Actuals*
  tab (Business Line, Arkieva ABC, Ship To Sub Region, Material, Arkieva
  Review Req, Arkieva Active Status, Arkieva Pattern, Data), with Arkieva
  Active Status defaulting to Active + Sparse.
- Scans every Key and lists the **top N anomalies (default 20) per Ship
  To Sub Region**, ranked by a severity score, each tagged with a
  **reason code**:
  - **Trend reversal** — the forecast trends opposite to recent history.
  - **Trend magnitude mismatch** — same direction, but the forecast slope
    is disproportionately steeper/flatter than history.
  - **Seasonality not captured** — history is clearly seasonal but the
    forecast is (near) flat / diverges from the historical shape.
  - **Outliers in sales history (IQR)** — multiple history points violate
    the IQR fences (biggest offenders by count).
  A Key may carry several codes; the most severe is the *primary* reason.
- A reason-code breakdown chart and a per-region severity table give an
  at-a-glance picture, and the table is exportable to CSV.
- **Drill-down**: pick any Key to see a focused chart (history, forecast,
  both trend lines, and IQR outlier markers), the supporting metrics, and
  — for every applicable reason code — a **root-cause explanation** and
  **recommended next steps** written from a demand-planning standpoint.

#### How anomalies are detected and scored
For each Key, the engine compares the Sales History (up to the
history/forecast boundary) against the Statistical Forecast:
1. **Trend**: linear trends are fitted to history and forecast. Opposite
   directions ⇒ *trend reversal*; same direction but the forecast rate is
   ≥ 3× or ≤ ⅓× the historical rate ⇒ *trend mismatch*.
2. **Seasonality**: the 3-year historical monthly seasonal index is
   correlated with the forecast's. A meaningfully seasonal history
   (amplitude ≥ 0.10) paired with a flat or poorly-correlated forecast
   (corr < 0.3) ⇒ *seasonality loss*.
3. **Outliers**: the IQR method counts history points outside Tukey's
   fences ⇒ *IQR outliers*.
Each detected issue gets a severity contribution (scaled by magnitude /
count); the Key's overall score is the strongest issue plus a small bonus
for multiple concurrent issues. Keys with no issues are omitted.

### Tab 2 — STF Variation & Exceptions
Compares the current **Statistical Forecast Committed (kg)** against the
**Statistical Forecast Committed Lag 1 (kg)** line (the prior cycle's
committed forecast), to surface where the forecast has moved.

- **Variance** = (Current − Lag 1) / Lag 1, computed per Key over a horizon,
  reported as both signed **STF Variance** and **Absolute STF Variance**,
  sorted high → low by the absolute value.
- **Month 4 (M4) auto-detection:** the "current month" is the last active
  month in Sales History; **M4 = current + 4**. (e.g. current = June →
  M1 July, M2 Aug, M3 Sep, **M4 October**.) M4 is re-detected every time a
  file is loaded.
- **Two horizons**, both starting at M4:
  - **M4 → fiscal year-end** — M4 through December of M4's year
    (fiscal year is Jan–Dec). e.g. Oct–Dec.
  - **M4 → next 12 months** — M4 plus the following 11 months (12 total).
- A **sliding window** slider shortens either horizon from the far end
  (M4 always stays the start).
- An **Absolute STF variation threshold** slider (default **5%**) flags
  exception Keys; the tab lists the **top 25** Keys above the threshold,
  with a red-gradient table, a variance bar chart, and a **month-level
  heatmap** that highlights *which months* drive each Key's variation.
- A **waterfall chart** decomposes the net STF change (Current − Lag 1)
  into driver buckets — **new item / distribution added**, **trend up**,
  **trend down**, and **disco / distribution loss** — rolled up across the
  current filter selection, with a drill-down to any single Key. The bucket
  deltas reconcile exactly to the net change.
- Uses the **same cascading filters** as *Forecast vs Actuals*, with
  Arkieva Active Status defaulting to Active + Sparse. Every visual reacts
  to the filters.

  If the uploaded file has no Lag 1 line, the tab explains that variance
  can't be computed and points to the raw Arkieva export that includes it.

### Tab 3 — Forecast vs Actuals
A faithful reproduction of the original Excel pivot dashboard:

- **All eight multi-select cascading filters at the top of the page**
  (no sidebar): Business Line, Arkieva ABC, Ship To Sub Region, Material,
  Arkieva Review Req, Arkieva Active Status, Arkieva Pattern, Data. Picking
  a value in one filter narrows the choices shown by every other filter
  (Excel-slicer style). Arkieva Active Status defaults to Active + Sparse.
- **Date-range slider** plus an **interactive chart range slider and quick
  range buttons** (6m / 1y / 2y / All) right on the plot.
- **Highly interactive, presentable line chart**: smooth (spline) **solid**
  lines for *every* series — both History and Forecast (no more dotted
  forecast lines); unified hover, spike lines, zoom/pan, and a legend below
  the chart for maximum plot area.
- **Clear History vs Forecast highlighting**: a tinted *History* band and
  *Forecast* band separated by a vertical boundary at the last month of
  Sales History; the two groups are distinguished by colour family and
  marker shape.
- **Trend lines** (toggle): a linear trend is fitted separately to the
  history and forecast periods and overlaid, with an **automatic
  interpretation** telling you whether the forecast follows the historical
  trend direction and flagging implausibly steep forecast growth.
- KPI cards (unique Keys, Materials, Months, Total volume).

### Tab 4 — Outlier Detection & Correction (IQR or Sigma)
- The **same six cascading filters at the top** (the *Data* filter is
  omitted because this view is intrinsically scoped to *Sales History
  (kg)*). Filter state is independent per tab.
- **Choose one of two methods**:
  - **IQR (Tukey's fences)** — flags points outside
    `[Q1 − k·IQR, Q3 + k·IQR]` and corrects each to the series **median**.
  - **Sigma (z-score)** — flags points beyond `mean ± n·σ` (mean/σ
    estimated **robustly**, after trimming gross outliers via the median ±
    MAD, so the bounds aren't inflated by the very outliers being detected)
    and corrects each to the **mean**.
  Both corrections are statistically valid central estimates.
- The per-Key chart shows the **original Sales History** and the
  **cleansed history** (after correction) as solid lines, the threshold
  band, outlier markers, and **History For Forecast (kg)** unchanged for
  comparison.
- A table of all corrected outliers and CSV exports of the cleansed
  history (per-Key and all-Keys).

### Tab 5 — Year-over-Year Seasonality
- Compares the **monthly seasonal shape** of the last **3 historical
  years** of Sales History against the forecast period. Each line is a
  seasonal index (1.0 = that period's average; 1.10 = ~10% above average).
- Shows each recent year individually, the 3-year average, and the
  forecast profile, with a reference line at 1.0.
- An **automatic interpretation** reports the correlation between the
  forecast and historical seasonal shapes, names the historical peak and
  trough months, and tells you whether the forecast preserves seasonality
  (≥ 0.7 = strong match), partially matches, diverges, or is flat where
  history is seasonal.

### Tab 6 — Statistical Forecast Adoption %
The **last tab**. Shows how much demand volume is being driven by the
**statistical forecast** rather than requiring manual review.

- **On-stat definition (granular level):** the granular unit is Business
  Line + Material + Ship To Sub Region + Arkieva Active Status. A
  combination is **on the Statistical Forecast** only when
  **`Arkieva Review Req` = No** AND its **`Arkieva Active Status`** is
  **Active** or **Sparse**.
- **Adoption % (region-based denominator):**

  ```
  Adoption = Σ on-stat forecast
             ──────────────────────────────────────────
             Σ total forecast of the same Ship To Sub Region
  ```

  Higher-level views (overall, by Business Line) roll up these granular
  numerators and denominators.
- **Active year onwards only:** the app finds the **active year** — the
  calendar year of the last active month in **Sales History (kg)** — and
  computes adoption for that year and every **future** year. Past years
  (whose forecast is already historical) are excluded. The active year is
  shown at the top of the tab.
- Uses the **same cascading filters** as the other tabs, with Arkieva
  Active Status defaulting to **Active + Sparse**. Every chart reacts
  instantly to filter changes.
- Visuals: an adoption-by-fiscal-year bar+trend combo, grouped bar charts
  **by Business Line** and **by Ship To Sub Region**, and a Business
  Line × year **heatmap**.
- At the bottom, an **interactive table** of the granular combinations on
  the Statistical Forecast (Business Line + Material + Ship To Sub Region +
  Active Status), with forecast volume per fiscal year and each
  combination's share of its Ship To Sub Region's total forecast — filtered
  live by the selections above and exportable to CSV.

### Input normalisation
- The statistical-forecast line may appear in the input as either
  **"Statistical Forecast (kg)"** or **"Statistical Forecast Committed
  (kg)"**. On every upload, any row whose `Data` is *"Statistical
  Forecast Committed (kg)"* is automatically relabelled to *"Statistical
  Forecast (kg)"*, so the rest of the dashboard is label-agnostic.

### Production-grade behavior
- **Auto-reset** when the uploaded file is removed or replaced.
- **Independent cascading filter state per tab.**
- **Robust validation** with clear error messages.
- **Determinism**: all results are reproducible.
- **Performance:** the wide file is reshaped once and cached; the value
  column is stored as 32-bit floats to roughly halve memory; every heavy
  computation (anomaly scan, STF variance/drivers, adoption roll-ups,
  outlier correction) is wrapped in `@st.cache_data`, so repeat renders and
  tab switches are instant for a given filter selection. The anomaly scan
  also pre-filters to Keys with enough history before the per-Key work.

---

## Expected input file format

Upload the **raw Arkieva "Forecast vs Actuals" export** directly — the
required data is always on the **first sheet**. The first sheet must
contain these identifier columns followed by month/date columns (one
column per month):

| Raw identifier columns | Date columns |
|---|---|
| `Business Line`, `Material`, `Ship To Sub Region`, `Arkieva ABC`, `Arkieva Pattern`, `Arkieva Active Status`, `Arkieva Review Req`, `Data` | `2023-06-01`, `2023-07-01`, …, `2028-05-01` |

The app **derives** two columns automatically and places them at the
front, so you do not need to prepare them:

- **Material code** — extracted from the `Material` string (the part after
  the double underscore, e.g. `2-ETHYL HEXANOL BULK__3000924` → `3000924`).
- **Key** — built as `Business Line_<material code>_Ship To Sub Region`
  (e.g. `Cleaning__BL02_3000924_North America`).

`Data` should take the values: `Sales History (kg)`,
`History For Forecast (kg)`, `Statistical Forecast (kg)` *(or
"Statistical Forecast Committed (kg)", which is auto-normalised)*,
`Statistical Forecast Committed Lag 1 (kg)` *(the prior cycle's committed
forecast, used by the STF Variation & Exceptions tab)*, and
`Final Demand Plan Lag 1 (kg)`.

**Notes on the new columns:**
- **Arkieva Review Req** replaces the legacy *Stat Flag*; boolean
  `True`/`False` values are shown as `Yes`/`No`.
- **Arkieva Active Status** is a new filter. The app **defaults to showing
  only `Active` and `Sparse` keys**; clear or change the filter on any tab
  to include other statuses (`Inactive`, `Obsolete`, `New Combination`,
  `Active New`, `-`). The 🔄 Reset button restores the Active + Sparse
  default.

> The same `Business Line_code_region` Key can appear with several
> `Arkieva Active Status` rows (each carrying its own monthly values); the
> default Active + Sparse filter plus the app's summing aggregation handle
> this correctly.

Older exports that already contain a built `Key` and a `Stat Flag` column
are still accepted (Stat Flag is mapped to *Arkieva Review Req*; a missing
*Arkieva Active Status* defaults to `Active`).

---

## Method notes

### Outlier detection & correction

For a single time series `x` (the Sales History of one Key):

**IQR (Tukey's fences)**
1. Compute the first and third quartiles `Q1`, `Q3` and `IQR = Q3 − Q1`.
2. A point is an outlier when it falls outside `[Q1 − k·IQR, Q3 + k·IQR]`
   (default `k = 1.5`).
3. **Correction**: each outlier is replaced by the series **median** — a
   robust central value that is not distorted by the outliers it replaces.

**Sigma (z-score / empirical rule)**
1. Estimate the mean `μ` and standard deviation `σ` **robustly**: first
   exclude gross outliers (those beyond `median ± n·1.4826·MAD`), then
   compute `μ`, `σ` on what remains, so the bounds aren't inflated by the
   very points being detected.
2. A point is an outlier when `|x − μ| > n·σ` (default `n = 3`).
3. **Correction**: each outlier is replaced by the (robust) **mean**.

In both methods, non-outlier points are left unchanged, producing the
*cleansed history*. The implementation is pure NumPy/pandas, so it deploys
cleanly to Streamlit Cloud. (A Hampel-filter implementation also remains in
the codebase for reference, but the UI exposes IQR and Sigma.)

### Trend analysis
A linear (ordinary least-squares) trend is fitted separately to the
history and forecast periods. The slope is reported per year and as a % of
the mean level, with R² as goodness-of-fit. The interpretation compares
the two directions (rising / falling / flat) and flags when the forecast
trend diverges from history or grows implausibly faster than history.

### Year-over-year seasonality
For each period a **seasonal index** is computed: the average value in each
calendar month divided by the overall average (so 1.0 = average month).
The forecast's seasonal shape is compared to the last 3 historical years
via Pearson correlation of the two 12-point profiles. ≥ 0.7 indicates the
forecast preserves the historical seasonal pattern; a flat forecast against
a seasonal history is flagged as a mismatch.

---

## Installation (local)

```bash
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

Python 3.9 or newer is required. Streamlit prints a local URL (typically
`http://localhost:8501`); open it and drag-and-drop your Excel file.

---

## Deploying to Streamlit Community Cloud via GitHub

Follow these steps to put the app online for free.

### 1. Prepare the project folder
Make a folder containing these files (note the `.streamlit` folder):

```
forecast-vs-actuals/
├── app.py
├── requirements.txt
├── README.md
└── .streamlit/
    └── config.toml
```

> Do **not** commit any customer Excel files — the app takes its input by
> upload at runtime.

### 2. Create a GitHub repository
1. Sign in at <https://github.com> (create a free account if needed).
2. Click **New repository** (the green button or `+` → *New repository*).
3. Name it e.g. `forecast-vs-actuals`, choose **Public** (Streamlit
   Community Cloud requires public repos on the free tier), and click
   **Create repository**.

### 3. Push your files to GitHub
Using the command line (with Git installed):

```bash
cd forecast-vs-actuals
git init
git add app.py requirements.txt README.md .streamlit/config.toml
git commit -m "Initial commit: Forecast vs Actuals dashboard"
git branch -M main
git remote add origin https://github.com/<your-username>/forecast-vs-actuals.git
git push -u origin main
```

(Alternatively, on the new repo page choose **uploading an existing
file** and drag the three files into the browser, then **Commit
changes** — no command line needed.)

### 4. Deploy on Streamlit Community Cloud
1. Go to <https://share.streamlit.io> and click **Sign in** → continue
   with GitHub, authorising access.
2. Click **Create app** → **Deploy a public app from GitHub**.
3. Fill in:
   - **Repository**: `<your-username>/forecast-vs-actuals`
   - **Branch**: `main`
   - **Main file path**: `app.py`
4. (Optional) Under **Advanced settings**, set the **Python version**
   (3.11 is a good choice) to match local testing.
5. Click **Deploy**.

Streamlit installs everything in `requirements.txt` and launches the app.
First build takes 1–3 minutes; afterwards you get a public URL like
`https://<your-app-name>.streamlit.app`.

### 5. Updating the deployed app
Any push to the `main` branch redeploys automatically:

```bash
git add app.py
git commit -m "Update dashboard"
git push
```

Streamlit detects the push and rebuilds within a minute.

### Troubleshooting deployment
| Symptom | Fix |
|---|---|
| Build fails on dependencies | Confirm `requirements.txt` is at the repo root and spelled exactly. |
| "Main module does not exist" | The **Main file path** must be `app.py` (case-sensitive). |
| App sleeps after inactivity | Free-tier apps sleep; the first visit wakes them in ~30 s. |
| Large uploads rejected | Streamlit's default upload limit is 200 MB; raise it with a `.streamlit/config.toml` containing `[server]\nmaxUploadSize = 500` if needed. |

---

## Project structure

```
.
├── app.py                    # Streamlit application
├── requirements.txt          # Pinned dependency ranges
├── README.md
└── .streamlit/
    └── config.toml           # Dark theme + max upload size
```

> Commit the `.streamlit/config.toml` file too — it is what makes the
> deployed app dark-themed from the first paint.

---

## How filters cascade

The filters behave like Excel slicers: selecting values in one narrows
the options in every other, the active filter still shows all values
consistent with the others, and stale selections are silently pruned
(earlier filters dominate, in the order Business Line → Arkieva ABC →
Ship To Sub Region → Material → Arkieva Review Req → Arkieva Active Status
→ Arkieva Pattern → Data). The 🔄 **Reset filters** button restores the
defaults (Arkieva Active Status → Active + Sparse) and clears every other
selection on that tab.

---

## License

Internal use. Adapt freely within your organization.
