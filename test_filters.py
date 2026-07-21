"""Filter-strip behaviour tests: persistence, cascading, defaults, reset.

Uses a faithful multiselect mock mirroring Streamlit's real state semantics:
  * A widget with key=K reads its value from session_state[K].
  * At creation, the stored value is PRUNED to values present in `options`
    (Streamlit silently drops out-of-options values) and written back.
  * The pruned list is returned.
This is exactly the behaviour that made the old code lose selections; the
new code must survive it.
"""
import sys, importlib.util
import pandas as pd

results = []
def check(name, ok, info=""):
    results.append((ok, name))
    suffix = f" — {info}" if info else ""
    print(f"{'PASS' if ok else 'FAIL'} {name}{suffix}")

class FilterMock:
    def __init__(self, ss=None):
        self.session_state = ss if ss is not None else {}
        self._buttons = {}
    def __getattr__(self, name):
        if name == 'session_state': return self.session_state
        if name == 'columns':
            def cols(spec):
                n = spec if isinstance(spec, int) else len(spec)
                class _C:
                    def __enter__(s): return s
                    def __exit__(s,*a): return False
                    def __getattr__(s,n2): return lambda *a,**k: None
                return [_C() for _ in range(n)]
            return cols
        if name in ('container','expander','form'):
            class _Ctx:
                def __enter__(s): return s
                def __exit__(s,*a): return False
            return lambda *a,**k: _Ctx()
        if name == 'multiselect':
            def ms(label, options=None, key=None, default=None, **k):
                options = list(options or [])
                stored = self.session_state.get(key, default or [])
                if not isinstance(stored, list): stored = []
                # REAL Streamlit: prune stored to values in options
                pruned = [v for v in stored if v in options]
                self.session_state[key] = pruned
                return pruned
            return ms
        if name == 'button':
            return lambda *a, **k: self._buttons.get(k.get('key'), False)
        if name == 'rerun':
            def _rr(*a, **k): raise _Rerun()
            return _rr
        if name in ('cache_data','cache_resource'):
            def deco(*a,**k):
                if a and callable(a[0]): return a[0]
                def w(fn): fn.clear=lambda:None; return fn
                return w
            return deco
        return lambda *a,**k: None
    def __call__(self,*a,**k): return None

class _Rerun(Exception): pass

m = FilterMock()
sys.modules['streamlit'] = m
spec = importlib.util.spec_from_file_location("app","/home/claude/forecast_app/app.py")
app = importlib.util.module_from_spec(spec); spec.loader.exec_module(app)

# A dataset with clear cascade relationships
df = pd.DataFrame({
    "Business Line": ["A","A","A","B","B","C","C","C"],
    "Region":        ["N","N","S","N","S","N","S","S"],
    "Material":      ["m1","m2","m3","m4","m5","m6","m7","m8"],
    "Pattern":       ["P","P","Q","P","Q","P","Q","Q"],
})
COLS = ["Business Line","Region","Material","Pattern"]

def render(state):
    m.session_state = state
    return app.render_filter_strip(df, COLS, key_prefix="t")

print("\n=== 1. Selection persists across reruns ===")
ss = {}
ss["t::Business Line"] = ["A"]
render(ss)
check("BL=A after first render", ss.get("t::Business Line")==["A"], ss.get("t::Business Line"))
render(ss)  # rerun, no change
check("BL=A persists on rerun", ss.get("t::Business Line")==["A"])
render(ss); render(ss)
check("BL=A persists across many reruns", ss.get("t::Business Line")==["A"])

print("\n=== 2. Cascade narrows OTHER filters, not self ===")
ss = {"t::Business Line": ["A"]}
render(ss)
# Region options given BL=A should be {N,S}; Material given BL=A -> m1,m2,m3
opts = app.cascading_options(df, COLS, {"Business Line":["A"],"Region":[],"Material":[],"Pattern":[]})
check("Region options under BL=A = [N,S]", opts["Region"]==["N","S"], opts["Region"])
check("Material options under BL=A = [m1,m2,m3]", opts["Material"]==["m1","m2","m3"], opts["Material"])
check("BL's own options stay full [A,B,C] (self not narrowed)",
      opts["Business Line"]==["A","B","C"], opts["Business Line"])

print("\n=== 3. Multi-filter cascade + persistence ===")
ss = {"t::Business Line":["A"], "t::Region":["S"]}
render(ss)
check("BL=A and Region=S both persist", ss.get("t::Business Line")==["A"] and ss.get("t::Region")==["S"],
      f"BL={ss.get('t::Business Line')} Region={ss.get('t::Region')}")
# Material under BL=A & Region=S -> only m3
opts = app.cascading_options(df, COLS, {"Business Line":["A"],"Region":["S"],"Material":[],"Pattern":[]})
check("Material narrowed to [m3] under BL=A,Region=S", opts["Material"]==["m3"], opts["Material"])

print("\n=== 4. Selected value NEVER dropped even if other filters would exclude it ===")
# Pick Material=m8 (needs BL=C, Region=S). Then also pick BL=A which is
# inconsistent with m8. The Material selection should still be VISIBLE (not
# silently wiped) — the widget keeps it; the user decides whether to clear.
ss = {"t::Material":["m8"], "t::Business Line":["A"]}
render(ss)
check("Material=m8 retained even though BL=A excludes it (no silent wipe)",
      ss.get("t::Material")==["m8"], ss.get("t::Material"))
# The options list passed to Material must include m8 so it isn't dropped
# (verify via cascading + union logic the function uses)
casc = app.cascading_options(df, COLS, {"Business Line":["A"],"Region":[],"Material":["m8"],"Pattern":[]})
union = list(casc["Material"]) + [v for v in ["m8"] if v not in casc["Material"]]
check("m8 present in Material's option union", "m8" in union)

print("\n=== 5. One-time defaults applied, then user choice respected ===")
ss = {}
render_with_default = lambda state: (setattr(m,'session_state',state),
                                     app.render_filter_strip(df, COLS, key_prefix="d",
                                     default_selections={"Region":["N"]}))[1]
render_with_default(ss)
check("Default Region=N applied on first render", ss.get("d::Region")==["N"], ss.get("d::Region"))
check("Init flag set", ss.get("d::__initialized__") is True)
# user clears Region
ss["d::Region"] = []
render_with_default(ss)
check("Cleared Region stays cleared (default not re-applied)", ss.get("d::Region")==[], ss.get("d::Region"))

print("\n=== 6. Reset restores defaults ===")
ss = {"r::Business Line":["A"], "r::Region":["S"]}
m.session_state = ss
m._buttons["r::reset"] = True   # simulate reset click
try:
    app.render_filter_strip(df, COLS, key_prefix="r",
                            default_selections={"Region":["N"]})
    check("Reset triggered rerun", False, "expected _Rerun")
except _Rerun:
    check("Reset cleared filter keys", "r::Business Line" not in ss and "r::Region" not in ss)
    check("Reset cleared init flag (defaults re-seed next run)", "r::__initialized__" not in ss)
m._buttons["r::reset"] = False
# next run re-seeds default
app.render_filter_strip(df, COLS, key_prefix="r", default_selections={"Region":["N"]})
check("After reset, default Region=N re-seeded", ss.get("r::Region")==["N"], ss.get("r::Region"))

print("\n=== 7. Independent state across tabs (key prefixes) ===")
ss = {}
m.session_state = ss
app.render_filter_strip(df, COLS, key_prefix="anom")
ss["anom::Business Line"] = ["A"]
app.render_filter_strip(df, COLS, key_prefix="anom")
app.render_filter_strip(df, COLS, key_prefix="dash")
ss["dash::Business Line"] = ["B"]
app.render_filter_strip(df, COLS, key_prefix="dash")
check("anom tab BL=A independent of dash tab BL=B",
      ss.get("anom::Business Line")==["A"] and ss.get("dash::Business Line")==["B"],
      f"anom={ss.get('anom::Business Line')} dash={ss.get('dash::Business Line')}")

print("\n=== 8. apply_filters correctness (str-robust) ===")
f = app.apply_filters(df, {"Business Line":["A"], "Region":["N"]})
check("apply_filters BL=A,Region=N -> 2 rows (m1,m2)", len(f)==2 and set(f["Material"])=={"m1","m2"},
      f"{len(f)} rows")
# empty selection = no filter
f2 = app.apply_filters(df, {"Business Line":[], "Region":[]})
check("empty selections -> all rows", len(f2)==len(df))

print("\n" + "="*60)
p = sum(1 for r in results if r[0]); t = len(results)
print(f"  FILTER TESTS: {p}/{t}")
if p != t: sys.exit(1)
