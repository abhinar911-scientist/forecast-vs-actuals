"""Synthetic Arkieva 'Forecast vs Actuals' export for tests.

Mirrors the raw export layout the app ingests (first sheet, one row per
grain x Data line, one column per month) without shipping customer data in
the repository. Deterministic (fixed seed).

Usage:
    from fixture import write_fixture
    path = write_fixture("/tmp/fva_fixture.xlsx")
"""
from __future__ import annotations

import numpy as np
import pandas as pd

BUSINESS_LINES = ["Cleaning__BL02", "Personal Care__BL03", "Agro__BL05"]
REGIONS = ["North America", "Europe", "Asia Pacific", "Latin America"]
PATTERNS = ["Stable", "Seasonal", "Erratic", "Linear"]
DATA_LINES = [
    "Sales History (kg)",
    "History For Forecast (kg)",
    "Statistical Forecast Committed (kg)",
    "Statistical Forecast Committed Lag 2 (kg)",
    "Final Demand Plan Lag 1 (kg)",
]
MONTHS = pd.date_range("2023-07-01", "2028-06-01", freq="MS")
BOUNDARY = pd.Timestamp("2026-07-01")  # last month with Sales History


def _grains(rng: np.random.Generator) -> list[dict]:
    grains = []
    code = 3000100
    for bl_i, bl in enumerate(BUSINESS_LINES):
        for m in range(8):
            code += 7
            name = f"Dissolvine {chr(65 + m)}-{bl_i + 1}{m}, bulk__{code}"
            regions = rng.choice(REGIONS, size=rng.integers(1, 4), replace=False)
            for region in regions:
                status = rng.choice(["Active", "Sparse", "Inactive", "Obsolete"],
                                    p=[0.6, 0.2, 0.1, 0.1])
                pattern = rng.choice(PATTERNS)
                grains.append(dict(
                    bl=bl, material=name, region=str(region), status=str(status),
                    pattern=str(pattern), abc=str(rng.choice(["A", "B", "C"])),
                    review=bool(rng.random() < 0.4)))
    # A multi-status Key (same BL + material + region, two statuses).
    first = dict(grains[0])
    first["status"] = "Sparse" if first["status"] != "Sparse" else "Active"
    grains.append(first)
    return grains


def build_raw(seed: int = 7) -> pd.DataFrame:
    rng = np.random.default_rng(seed)
    rows = []
    for i, g in enumerate(_grains(rng)):
        base = rng.uniform(5_000, 80_000)
        season = 1 + 0.25 * np.sin(np.arange(len(MONTHS)) * 2 * np.pi / 12 + i)
        trend = 1 + np.linspace(0, rng.uniform(-0.2, 0.3), len(MONTHS))
        hist = base * season * trend * rng.normal(1, 0.08, len(MONTHS))
        spikes = rng.random(len(MONTHS)) < 0.04
        hist[spikes] *= rng.uniform(2.5, 4.0)
        hist[MONTHS > BOUNDARY] = 0.0
        cleansed = np.where(spikes, base * season * trend, hist)
        stf = base * season * trend * rng.normal(1, 0.03, len(MONTHS))
        stf_lag = stf * rng.normal(1 + rng.uniform(-0.15, 0.15), 0.02, len(MONTHS))
        fdp = stf * rng.normal(1.02, 0.05, len(MONTHS))
        series = [hist, cleansed, stf, stf_lag, fdp]

        # Prior-cycle attributes: exercise all four change buckets.
        lag_status, lag_pattern = g["status"], g["pattern"]
        if i % 9 == 1:
            lag_status = "Active New"                       # status only
        elif i % 9 == 2:
            lag_pattern = "Sporadic"                        # pattern only
        elif i % 9 == 3:
            lag_status, lag_pattern = "Unknown", "Unknown"  # both

        for data, values in zip(DATA_LINES, series):
            row = {
                "Business Line": g["bl"], "Material": g["material"],
                "Ship To Sub Region": g["region"], "Arkieva ABC": g["abc"],
                "Arkieva Pattern": g["pattern"],
                "Arkieva Active Status": g["status"],
                "Arkieva Review Req": g["review"], "Data": data,
                "Arkieva Active Status Lag 1": lag_status,
                "Arkieva Pattern Lag 1": lag_pattern,
            }
            row.update({m.to_pydatetime(): round(float(v), 1)
                        for m, v in zip(MONTHS, values)})
            rows.append(row)
    return pd.DataFrame(rows)


def write_fixture(path: str, seed: int = 7) -> str:
    build_raw(seed).to_excel(path, index=False)
    return path


if __name__ == "__main__":
    import sys
    print(write_fixture(sys.argv[1] if len(sys.argv) > 1 else "fva_fixture.xlsx"))
