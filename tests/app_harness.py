"""Runs the REAL app (`app.main()`) under Streamlit's test runtime.

Streamlit's AppTest cannot click through a login form or drive a file
uploader, so this harness stubs exactly those two things and nothing else:
it marks the session as signed in and makes st.file_uploader return the
workbook named by FVA_XLSX_PATH. Every tab, filter and widget is the real
code path.
"""
import importlib.util
import os
import time

import streamlit as st

APP_PATH = os.environ["FVA_APP_PATH"]
XLSX_PATH = os.environ["FVA_XLSX_PATH"]


class _Upload:
    name = os.path.basename(XLSX_PATH)

    def __init__(self):
        with open(XLSX_PATH, "rb") as fh:
            self._bytes = fh.read()
        self.size = len(self._bytes)

    def getvalue(self):
        return self._bytes

    def read(self):
        return self._bytes


_upload = _Upload()
st.file_uploader = lambda *a, **k: _upload
st.session_state.setdefault("authenticated", True)
st.session_state.setdefault("auth_user", "abhi")

_spec = importlib.util.spec_from_file_location("fva_app", APP_PATH)
_app = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_app)

_t0 = time.perf_counter()
_app.main()
st.session_state["__run_secs__"] = round(time.perf_counter() - _t0, 2)
