# Forecast vs Actuals — Interactive Dashboard

A production-grade Streamlit application that replicates the Excel
**"Forecast vs Actuals" pivot dashboard** and adds an interactive
**Hampel-filter outlier detection & correction** view for the
sales-history time series.

The app accepts the standard *Forecast vs Actuals* monthly Excel file
(file name typically ends with the reporting month, e.g.
`Forecast_vs_Actuals_Cleaning_and_Personal_Care_April.xlsx`) and renders
every chart dynamically from the most recently uploaded file.

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

### Tab 1 — Forecast vs Actuals
A faithful reproduction of the original Excel pivot dashboard:

- **All seven multi-select cascading filters at the top of the page**
  (no sidebar): Business Line, Arkieva ABC, Ship To Sub Region, Material,
  Stat Flag, Arkieva Pattern, Data. Picking a value in one filter narrows
  the choices shown by every other filter (Excel-slicer style).
- **Date-range slider** for the months shown.
- **Maximised chart area**: the legend sits *below* the chart, so all
  series are clearly visible.
- **Clear History vs Forecast highlighting**: a tinted blue *History*
  band and orange *Forecast* band separated by a vertical boundary at the
  last month of Sales History; History series are solid lines, Forecast
  series are dashed/dotted.
- KPI cards (unique Keys, Materials, Months, Total volume).

### Tab 2 — Outlier Detection & Correction (Hampel filter)
- The **same six cascading filters at the top** (the *Data* filter is
  omitted because this view is intrinsically scoped to *Sales History
  (kg)*). Filter state is independent from Tab 1.
- Detects **and corrects** outliers in **Sales History (kg)** using the
  **Hampel filter** (see method notes below).
- Produces a **"Hampel filter cleansed history"** for each Key — the
  Sales History with every flagged outlier replaced by its local
  (window) median. This cleansed series is drawn on the per-Key chart
  alongside the original Sales History.
- **History For Forecast (kg)** is shown **unchanged** on the same chart
  for comparison — it is never modified.
- Adjustable Hampel parameters: rolling-window half-width and `n_sigma`
  threshold (default 3.0).
- Per-Key inspection chart (original + cleansed + History-For-Forecast +
  threshold band + outlier markers), a table of all corrected outliers,
  and CSV exports of the cleansed history (per-Key and all-Keys).

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
- **Caching** for instant filter changes.
- **Determinism**: all results are reproducible.

---

## Expected input file format

The first sheet must contain these identifier columns followed by
month/date columns (one column per month):

| Identifier columns | Date columns |
|---|---|
| `Key`, `Business Line`, `Material`, `Material code`, `Ship To Sub Region`, `Arkieva ABC`, `Arkieva Pattern`, `Stat Flag`, `Data` | `2023-04-01`, `2023-05-01`, …, `2028-03-01` |

`Data` should take the values: `Sales History (kg)`,
`History For Forecast (kg)`, `Statistical Forecast (kg)` *(or
"Statistical Forecast Committed (kg)", which is auto-normalised)*, and
`Final Demand Plan Lag 1 (kg)`.

---

## Method notes — Hampel filter

For a single time series `x` (the Sales History of one Key):

1. A **centred rolling window** of half-width `w` slides over the series,
   so the full window spans `2·w + 1` points.
2. For each point `i`, compute the window **median** `m_i` and the
   **Median Absolute Deviation** `MAD_i = median(|x − m_i|)` over the window.
3. The robust scale estimate is `σ_i = 1.4826 · MAD_i` (the `1.4826`
   factor makes the MAD a consistent estimator of the standard deviation
   for Gaussian data), and the threshold is `n_sigma · σ_i`.
4. Point `i` is an **outlier** when `|x_i − m_i| > n_sigma · σ_i`.
5. **Correction**: each outlier is replaced by the window median `m_i`,
   producing the *cleansed history*. All non-outlier points are unchanged.

This matches the algorithm in
<https://medium.com/@migueloteropedrido/hampel-filter-with-python-17db1d265375>
and the `hampel` PyPI library (verified to give identical outlier
indices in tests). The implementation here is dependency-free (pure
NumPy/pandas) so it deploys cleanly to Streamlit Cloud.

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
Make a folder containing exactly these three files:

```
forecast-vs-actuals/
├── app.py
├── requirements.txt
└── README.md
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
git add app.py requirements.txt README.md
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
├── app.py              # Streamlit application
├── requirements.txt    # Pinned dependency ranges
└── README.md
```

---

## How filters cascade

The filters behave like Excel slicers: selecting values in one narrows
the options in every other, the active filter still shows all values
consistent with the others, and stale selections are silently pruned
(earlier filters dominate, in the order Business Line → Arkieva ABC →
Ship To Sub Region → Material → Stat Flag → Arkieva Pattern → Data). The
🔄 **Reset filters** button clears every selection on that tab.

---

## License

Internal use. Adapt freely within your organization.
