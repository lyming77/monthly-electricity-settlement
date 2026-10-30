"""Streamlit UI interactions. Private fixtures are optional and remain external."""
from __future__ import annotations

import hashlib
import importlib.util
import io
import os
from pathlib import Path
import sys
import unittest
import zipfile

HAS_STREAMLIT = importlib.util.find_spec("streamlit") is not None
CLOUD_ROOT = Path(__file__).resolve().parents[1]
UI = CLOUD_ROOT / "streamlit_app.py"
if str(CLOUD_ROOT) not in sys.path:
    sys.path.insert(0, str(CLOUD_ROOT))
if HAS_STREAMLIT:
    from streamlit.testing.v1 import AppTest


def button_named(app, label):
    return next(button for button in app.button if button.label == label)


def fixture_paths():
    root = Path(os.environ["SETTLEMENT_FIXTURES"]) if os.environ.get("SETTLEMENT_FIXTURES") else CLOUD_ROOT.parent
    target = root / "2026年月度电价分析(目标文件).xlsx"
    for folder in ("2026.9（正式结算）", "9月月清单（10.10公布）"):
        sources = sorted((root / folder).glob("*.xlsx"))
        if target.is_file() and len(sources) == 8:
            return target, sources
    raise unittest.SkipTest("Private September workbooks unavailable; set SETTLEMENT_FIXTURES to test real uploads")


@unittest.skipUnless(HAS_STREAMLIT, "Streamlit is not installed; UI tests have not run")
class StreamlitUITests(unittest.TestCase):
    def test_initial_page(self):
        app = AppTest.from_file(str(UI), default_timeout=20).run()
        self.assertFalse(app.exception)
        self.assertEqual(app.title[0].value, "月度结算数据汇总")
        self.assertEqual(len(app.file_uploader), 2)
        self.assertTrue(button_named(app, "分析并预览").disabled)
        self.assertTrue(button_named(app, "生成更新后的Excel").disabled)

    def test_two_month_inputs_and_stale_result_invalidation(self):
        wrapper = """
import sys
sys.path.insert(0, %r)
import streamlit as st
from streamlit_app import render_supplements, parse_supplements
raw = render_supplements([10, 11], 0)
parsed, errors = parse_supplements(raw)
st.session_state['parsed'] = parsed
st.session_state['parse_errors'] = errors
for error in errors:
    st.error(error)
""" % str(CLOUD_ROOT)
        app = AppTest.from_string(wrapper, default_timeout=20).run()
        self.assertFalse(app.exception)
        self.assertEqual(len(app.text_input), 8)
        self.assertEqual(app.session_state["parsed"], {})
        app.session_state["preview"] = {"test": True}
        app.session_state["result"] = {"test": True}
        app.text_input(key="supplement_0_10_风电").set_value("300.5").run()
        self.assertFalse(app.exception)
        self.assertEqual(app.session_state["parsed"][10]["风电"], 300.5)
        self.assertNotIn("preview", app.session_state)
        self.assertNotIn("result", app.session_state)
        app.text_input(key="supplement_0_10_分布式电量").set_value("0").run()
        self.assertTrue(app.session_state["parse_errors"])
        app.text_input(key="supplement_0_10_分布式电费").set_value("0").run()
        self.assertEqual(app.session_state["parsed"][10]["分布式电量"], 0)
        self.assertFalse(app.session_state["parse_errors"])
        app.text_input(key="supplement_0_11_光伏").set_value("nan").run()
        self.assertTrue(app.session_state["parse_errors"])
        app.text_input(key="supplement_0_11_光伏").set_value("").run()
        self.assertFalse(app.session_state["parse_errors"])

    def test_content_sensitive_upload_signature(self):
        from streamlit_app import upload_signature
        first = {"name": "same.xlsx", "data": b"aaa"}
        second = {"name": "same.xlsx", "data": b"bbb"}
        self.assertNotEqual(upload_signature(first, []), upload_signature(second, []))

    def test_real_upload_preview_download_and_invalidation(self):
        target, sources = fixture_paths()
        target_bytes = target.read_bytes()
        original_hash = hashlib.sha256(target_bytes).hexdigest()
        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        app = AppTest.from_file(str(UI), default_timeout=40).run()
        app.file_uploader(key="target_upload_0").set_value((target.name, target_bytes, mime)).run()
        app.file_uploader(key="sources_upload_0").set_value([(p.name, p.read_bytes(), mime) for p in sources]).run()
        self.assertFalse(app.exception)
        self.assertEqual(len(app.text_input), 4)
        for field, value in (("风电", "300"), ("光伏", "250"), ("分布式电量", "0"), ("分布式电费", "0")):
            app.text_input(key=f"supplement_0_9_{field}").set_value(value).run()
        button_named(app, "分析并预览").click().run()
        self.assertFalse(app.exception)
        self.assertFalse(app.session_state["preview"]["errors"])
        self.assertEqual(len(app.metric), 4)
        button_named(app, "生成更新后的Excel").click().run()
        self.assertFalse(app.exception)
        self.assertEqual(len(app.download_button), 4)
        result = app.session_state["result"]
        self.assertTrue(zipfile.is_zipfile(io.BytesIO(result["workbook_bytes"])))
        self.assertTrue(zipfile.is_zipfile(io.BytesIO(result["bundle_bytes"])))
        app.text_input(key="supplement_0_9_风电").set_value("301").run()
        self.assertEqual(len(app.download_button), 0)
        self.assertNotIn("result", app.session_state)
        self.assertNotIn("preview", app.session_state)
        self.assertTrue(button_named(app, "生成更新后的Excel").disabled)
        button_named(app, "清空本次会话").click().run()
        self.assertFalse(app.exception)
        self.assertIsNone(app.file_uploader(key="target_upload_1").value)
        self.assertFalse(app.file_uploader(key="sources_upload_1").value)
        self.assertEqual(len(app.text_input), 0)
        self.assertEqual(hashlib.sha256(target.read_bytes()).hexdigest(), original_hash)


if __name__ == "__main__":
    unittest.main(verbosity=2)
