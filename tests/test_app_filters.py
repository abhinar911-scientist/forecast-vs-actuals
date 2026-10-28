"""End-to-end filter tests on the REAL app, using Streamlit's own runtime.

These tests drive genuine widget clicks through streamlit.testing (AppTest),
so they exercise Streamlit's widget-identity and session-state machinery —
the layer where the "filters vanish" bug lived. (A mock of st.multiselect
cannot reproduce that bug, which is why earlier mock-based tests passed
while the app still reset filters.)

Run:   python tests/test_app_filters.py [optional_real_export.xlsx]
With no argument a synthetic Arkieva-format workbook is generated.
"""
from __future__ import annotations

import os
import sys
import tempfile
import time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from streamlit.testing.v1 import AppTest  # noqa: E402

from fixture import write_fixture  # noqa: E402

APP_PATH = os.path.join(os.path.dirname(HERE), "app.py")
STATUS = "Arkieva Active Status"
TAB_PREFIXES = ["anom", "stf", "dash", "outlier", "season", "adopt"]
DEFAULTED = {"anom", "dash", "outlier", "season", "adopt"}  # Active + Sparse

results: list[tuple[bool, str]] = []


def check(name: str, ok: bool, info: str = "") -> None:
    results.append((bool(ok), name))
    print(f"{'PASS' if ok else 'FAIL'}  {name}" + (f"  [{info}]" if info and not ok else ""))


# ---- helpers ---------------------------------------------------------------
def ms(at, prefix, col):
    key = f"{prefix}::{col}::w"
    found = [m for m in at.multiselect if m.key == key]
    return found[0] if found else None


def val(at, prefix, col):
    m = ms(at, prefix, col)
    return None if m is None else list(m.value)


def store(at, prefix, col):
    key = f"{prefix}::{col}"
    return list(at.session_state[key]) if key in at.session_state else []


def no_exceptions(at) -> bool:
    return len(at.exception) == 0


def first_other(options, exclude):
    for o in options:
        if o not in exclude:
            return o
    return None


def main(xlsx_path: str) -> int:
    os.environ["FVA_APP_PATH"] = APP_PATH
    os.environ["FVA_XLSX_PATH"] = xlsx_path
    at = AppTest.from_file(os.path.join(HERE, "app_harness.py"),
                           default_timeout=900)
    t0 = time.perf_counter()
    at.run()
    print(f"first run: {time.perf_counter() - t0:.1f}s")
    check("App runs without exceptions", no_exceptions(at),
          str([e.value for e in at.exception][:2]))
    if not no_exceptions(at):
        return 1

    # ---- 1. Defaults ---------------------------------------------------------
    print("\n== Defaults")
    for p in TAB_PREFIXES:
        status = val(at, p, STATUS)
        if p in DEFAULTED:
            check(f"[{p}] Active Status defaults to Active + Sparse",
                  status is not None and set(status) <= {"Active", "Sparse"}
                  and len(status) >= 1, str(status))
        else:
            check(f"[{p}] no status default", status == [], str(status))

    # ---- 2. Persistence through cascading changes, per tab -------------------
    print("\n== Selections persist while other filters cascade")
    for p in TAB_PREFIXES:
        bl = ms(at, p, "Business Line")
        if bl is None:
            check(f"[{p}] Business Line filter present", False)
            continue
        pick_bl = bl.options[0]
        bl.select(pick_bl).run()
        region = ms(at, p, "Ship To Sub Region")
        pick_region = region.options[0]
        region.select(pick_region).run()
        mat = ms(at, p, "Material")
        pick_mat = mat.options[0]
        mat.select(pick_mat).run()
        ok = (val(at, p, "Business Line") == [pick_bl]
              and val(at, p, "Ship To Sub Region") == [pick_region]
              and val(at, p, "Material") == [pick_mat])
        check(f"[{p}] 3 filters stay selected after cascading changes", ok,
              f"BL={val(at, p, 'Business Line')} R={val(at, p, 'Ship To Sub Region')} "
              f"M={val(at, p, 'Material')}")
        at.run()
        check(f"[{p}] selections survive a plain rerun",
              val(at, p, "Business Line") == [pick_bl]
              and val(at, p, "Material") == [pick_mat])
        check(f"[{p}] widget value matches persistent store",
              store(at, p, "Business Line") == [pick_bl]
              and store(at, p, "Material") == [pick_mat])
        check(f"[{p}] no exceptions after filtering", no_exceptions(at),
              str([e.value for e in at.exception][:1]))

    # ---- 3. Excel-slicer cascade -------------------------------------------
    print("\n== Cascade narrows the other filters, never itself")
    p = "dash"
    ms(at, p, "Business Line").set_value([]).run()
    ms(at, p, "Ship To Sub Region").set_value([]).run()
    ms(at, p, "Material").set_value([]).run()
    all_bl = list(ms(at, p, "Business Line").options)
    all_mat = list(ms(at, p, "Material").options)
    pick = all_bl[0]
    ms(at, p, "Business Line").select(pick).run()
    narrowed = list(ms(at, p, "Material").options)
    check("Material list narrows when a Business Line is chosen",
          0 < len(narrowed) < len(all_mat), f"{len(narrowed)} of {len(all_mat)}")
    check("Business Line keeps its full list (self not narrowed)",
          list(ms(at, p, "Business Line").options) == all_bl)
    ms(at, p, "Business Line").select(all_bl[1]).run()
    check("Multi-select within one filter widens the others",
          len(ms(at, p, "Material").options) > len(narrowed))
    ms(at, p, "Business Line").unselect(all_bl[1]).run()
    check("Removing a value persists",
          val(at, p, "Business Line") == [pick])

    # ---- 4. Tabs are independent ----------------------------------------------
    print("\n== Tabs filter independently")
    before = {q: val(at, q, "Business Line") for q in TAB_PREFIXES}
    other = first_other(ms(at, "season", "Business Line").options,
                        before["season"] or [])
    ms(at, "season", "Business Line").set_value([other]).run()
    after = {q: val(at, q, "Business Line") for q in TAB_PREFIXES}
    check("Changing one tab's filter leaves every other tab untouched",
          all(after[q] == before[q] for q in TAB_PREFIXES if q != "season"),
          f"{before} -> {after}")

    # ---- 5. Reset ----------------------------------------------------------------
    print("\n== Reset")
    reset = [b for b in at.button if b.key == "outlier::reset"][0]
    reset.click().run()
    check("Reset clears the tab's selections",
          val(at, "outlier", "Business Line") == []
          and val(at, "outlier", "Material") == [])
    status = val(at, "outlier", STATUS)
    check("Reset restores the Active + Sparse default",
          status is not None and set(status) <= {"Active", "Sparse"} and status,
          str(status))
    check("Reset on one tab leaves other tabs' filters",
          val(at, "dash", "Business Line") == [pick])
    at.run()
    check("Post-reset state is stable on rerun",
          val(at, "outlier", "Business Line") == [])

    # ---- 6. Sticky pickers ------------------------------------------------------
    print("\n== Data-dependent pickers keep their choice")
    for q in TAB_PREFIXES:  # widen everything so pickers have options
        for col in ["Business Line", "Ship To Sub Region", "Material"]:
            if ms(at, q, col) is not None and val(at, q, col):
                ms(at, q, col).set_value([]).run()

    sb = [s for s in at.selectbox if s.key == "outlier::chosen_key"]
    if sb and len(sb[0].options) > 1:
        target = sb[0].options[1]
        sb[0].select(target).run()
        status_ms = ms(at, "outlier", STATUS)
        status_ms.set_value(list(status_ms.options)).run()  # widen another filter
        now = [s for s in at.selectbox if s.key == "outlier::chosen_key"][0].value
        check("Outlier Key picker keeps its Key when another filter changes",
              now == target, f"{target} -> {now}")
    else:
        check("Outlier Key picker present with >1 Key", False)

    dk = [s for s in at.selectbox if s.key == "anom::drill_key"]
    if dk and len(dk[0].options) > 1:
        target = dk[0].options[1]
        dk[0].select(target).run()
        at.run()
        now = [s for s in at.selectbox if s.key == "anom::drill_key"][0].value
        check("Anomaly drill-down Key survives reruns", now == target,
              f"{target} -> {now}")

    rng = [s for s in at.select_slider if s.key == "stf::range"]
    if rng and len(rng[0].options) > 3:
        opts = list(rng[0].options)
        rng[0].set_range(opts[1], opts[-2]).run()
        set_to = tuple(str(v) for v in
                       [s for s in at.select_slider if s.key == "stf::range"][0].value)
        st_ms = ms(at, "stf", "Arkieva ABC")
        st_ms.select(st_ms.options[0]).run()
        now = tuple(str(v) for v in
                    [s for s in at.select_slider if s.key == "stf::range"][0].value)
        check("STF horizon range survives a filter change", now == set_to,
              f"{set_to} -> {now}")
        mode = [r for r in at.radio if r.key == "stf::mode"][0]
        fy = [o for o in mode.options if o.startswith("FY")]
        if fy:
            mode.set_value(fy[0]).run()
            st_ms = ms(at, "stf", "Arkieva ABC")
            st_ms.set_value([]).run()
            now_mode = [r for r in at.radio if r.key == "stf::mode"][0].value
            check("STF fiscal-year preset survives a filter change",
                  now_mode == fy[0], str(now_mode))
    else:
        # Older exports have no prior-cycle line; the tab must then explain
        # that instead of showing the horizon controls.
        explains = any("prior-cycle" in str(i.value) for i in at.info)
        check("STF tab: horizon slider, or a clear note when the file has "
              "no prior-cycle line", explains)

    check("No exceptions at the end of the session", no_exceptions(at),
          str([e.value for e in at.exception][:1]))

    passed = sum(ok for ok, _ in results)
    print(f"\nAPP FILTER TESTS: {passed}/{len(results)} passed")
    return 0 if passed == len(results) else 1


if __name__ == "__main__":
    if len(sys.argv) > 1:
        path = sys.argv[1]
    else:
        path = os.path.join(tempfile.mkdtemp(), "fva_fixture.xlsx")
        write_fixture(path)
    sys.exit(main(path))
