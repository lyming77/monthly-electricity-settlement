"""Cloud byte-boundary, isolation, archive, and real-workbook regressions.

Private workbooks are read only from SETTLEMENT_FIXTURES or the parent workspace.
They are never copied into this distribution. Run: python -m unittest discover -s tests -v
"""
from __future__ import annotations

import hashlib
import importlib
import io
import json
import os
from pathlib import Path
import sys
import struct
import unittest
from unittest import mock
import warnings
import zipfile
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor

CLOUD_ROOT = Path(__file__).resolve().parents[1]
if str(CLOUD_ROOT) not in sys.path:
    sys.path.insert(0, str(CLOUD_ROOT))


def zip_bytes(entries):
    buffer = io.BytesIO()
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for name, value in entries:
                if isinstance(value, zipfile.ZipInfo):
                    archive.writestr(value, b"linked")
                elif "\\" in name:
                    # Windows' ZipInfo constructor otherwise normalizes the malicious separator.
                    info = zipfile.ZipInfo("fixture")
                    info.filename = info.orig_filename = name
                    info.compress_type = zipfile.ZIP_DEFLATED
                    archive.writestr(info, value)
                else:
                    archive.writestr(name, value)
    return buffer.getvalue()


def upload(name, data):
    return {"name": name, "data": data}


def minimal_xlsx():
    main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    return zip_bytes([
        ("[Content_Types].xml", b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'),
        ("_rels/.rels", b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'),
        ("xl/workbook.xml", ('<workbook xmlns="'+main+'" xmlns:r="'+rel+'"><sheets><sheet name="数据表" sheetId="1" r:id="rId1"/></sheets></workbook>').encode()),
        ("xl/_rels/workbook.xml.rels", b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/></Relationships>'),
        ("xl/worksheets/sheet1.xml", ('<worksheet xmlns="'+main+'"><sheetData><row r="2"><c r="C2" t="inlineStr"><is><t>结算单元名称</t></is></c><c r="BH2" t="inlineStr"><is><t>总结算电费</t></is></c><c r="BQ2" t="inlineStr"><is><t>财政补贴收入</t></is></c></row></sheetData></worksheet>').encode()),
    ])


def fixture_directory():
    explicit = os.environ.get("SETTLEMENT_FIXTURES")
    candidates = [Path(explicit)] if explicit else [CLOUD_ROOT.parent]
    for candidate in candidates:
        if (candidate / "2026年月度电价分析(目标文件).xlsx").is_file():
            return candidate
    return None


def fixture_uploads():
    root = fixture_directory()
    if root is None:
        raise unittest.SkipTest("Set SETTLEMENT_FIXTURES to run private workbook regressions")
    target_path = root / "2026年月度电价分析(目标文件).xlsx"
    source_dir = root / "9月月清单（10.10公布）"
    if not source_dir.is_dir():
        raise unittest.SkipTest("September fixture directory is unavailable")
    paths = sorted(source_dir.glob("*.xlsx"))
    if len(paths) != 8:
        raise unittest.SkipTest("Expected eight September source fixtures")
    return root, upload(target_path.name, target_path.read_bytes()), [upload(p.name, p.read_bytes()) for p in paths]


def month10_copy(data):
    """Return a synthetic October copy; alter no real source or target path."""
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        parts = {info.filename: archive.read(info) for info in archive.infolist()}
    shared = []
    if "xl/sharedStrings.xml" in parts:
        shared = ["".join(node.itertext()) for node in ET.fromstring(parts["xl/sharedStrings.xml"])]
    tree = ET.fromstring(parts["xl/worksheets/sheet1.xml"])
    for cell in tree.findall("{0}sheetData/{0}row/{0}c".format("{"+namespace+"}")):
        address = cell.attrib["r"]
        value = cell.find("{"+namespace+"}v")
        if address == "A1":
            title = shared[int(value.text)] if cell.attrib.get("t") == "s" else "".join(cell.itertext())
            cell.attrib["t"] = "inlineStr"
            for child in list(cell):
                cell.remove(child)
            inline = ET.SubElement(cell, "{"+namespace+"}is")
            ET.SubElement(inline, "{"+namespace+"}t").text = title.replace("2026年9月", "2026年10月")
        elif address.startswith("B") and address[1:].isdigit() and value is not None and cell.attrib.get("t") not in ("s", "inlineStr"):
            try:
                number = float(value.text)
                if 46266 <= number <= 46295:
                    value.text = str(number+30)
            except (ValueError, TypeError):
                pass
    parts["xl/worksheets/sheet1.xml"] = ET.tostring(tree, encoding="utf-8", xml_declaration=True)
    for name, content in list(parts.items()):
        if name.endswith(".xml"):
            root = ET.fromstring(content)
            changed = False
            for node in root.iter():
                if node.text and "2026-09-" in node.text:
                    node.text = node.text.replace("2026-09-", "2026-10-")
                    changed = True
            if changed:
                parts[name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return zip_bytes(list(parts.items()))


class CloudArchiveTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.service = importlib.import_module("service")
        cls.target = upload("target.xlsx", minimal_xlsx())

    def assert_bad_archive(self, source, message_pattern=None):
        with self.assertRaises(ValueError) as raised:
            self.service.inspect_inputs(self.target, [source])
        if message_pattern:
            self.assertRegex(str(raised.exception), message_pattern)

    def test_outer_paths_cannot_escape(self):
        for member in ("../escape.xlsx", "/escape.xlsx", "C:/escape.xlsx", "folder\\escape.xlsx", "folder/../../escape.xlsx"):
            with self.subTest(member=member):
                self.assert_bad_archive(upload("month.zip", zip_bytes([(member, minimal_xlsx())])), "路径")

    def test_duplicate_outer_paths_rejected(self):
        self.assert_bad_archive(upload("month.zip", zip_bytes([("same.xlsx", minimal_xlsx()), ("same.xlsx", minimal_xlsx())])), "重复")

    def test_duplicate_canonical_paths_rejected(self):
        self.assert_bad_archive(upload("month.zip", zip_bytes([("folder/same.xlsx", minimal_xlsx()), ("folder/./same.xlsx", minimal_xlsx())])), "重复")

    def test_encrypted_outer_member_rejected_before_read(self):
        data = bytearray(zip_bytes([("locked.xlsx", minimal_xlsx())]))
        local = data.index(b"PK\x03\x04")
        central = data.index(b"PK\x01\x02")
        for position in (local+6, central+8):
            flags = struct.unpack_from("<H", data, position)[0]
            struct.pack_into("<H", data, position, flags | 1)
        self.assert_bad_archive(upload("month.zip", bytes(data)), "加密")

    def test_outer_symlink_rejected(self):
        info = zipfile.ZipInfo("linked.xlsx")
        info.create_system = 3
        info.external_attr = 0o120777 << 16
        self.assert_bad_archive(upload("month.zip", zip_bytes([("linked.xlsx", info)])), "链接")

    def test_inner_xlsx_paths_also_validated(self):
        self.assert_bad_archive(upload("bad.xlsx", zip_bytes([("../escape.xml", b"bad")])), "路径")

    def test_inner_duplicate_member_rejected(self):
        self.assert_bad_archive(upload("bad.xlsx", zip_bytes([("xl/workbook.xml", b"a"), ("xl/workbook.xml", b"b")])), "重复")

    def test_xml_entity_declaration_rejected(self):
        bad = b'<!DOCTYPE a [<!ENTITY x "expanded">]><a>&x;</a>'
        with zipfile.ZipFile(io.BytesIO(minimal_xlsx())) as z:
            entries = [(name, z.read(name)) for name in z.namelist()]
        for markup in (bad, bad.decode().encode("utf-16")):
            with self.subTest(encoding="utf16" if markup.startswith(b"\xff\xfe") else "utf8"):
                self.assert_bad_archive(upload("bad.xlsx", zip_bytes(entries+[("xl/custom.xml", markup)])), "XML.*实体")

    def test_high_expansion_archive_rejected(self):
        self.assert_bad_archive(upload("month.zip", zip_bytes([("payload.xlsx", b"x" * (4*1024*1024))])), "压缩比")
        with zipfile.ZipFile(io.BytesIO(minimal_xlsx())) as z:
            entries = [(name, z.read(name)) for name in z.namelist()]
        self.assert_bad_archive(upload("bad.xlsx", zip_bytes(entries+[("xl/payload.bin", b"x" * (4*1024*1024))])), "压缩比")

    def test_member_count_budget_enforced(self):
        with zipfile.ZipFile(io.BytesIO(minimal_xlsx())) as z:
            entries = [(name, z.read(name)) for name in z.namelist()]
        with mock.patch.object(self.service, "MAX_MEMBERS", len(entries)):
            self.assert_bad_archive(upload("bad.xlsx", zip_bytes(entries+[("extra.bin", b"one")])), "文件过多")

    def test_member_size_budget_enforced(self):
        with zipfile.ZipFile(io.BytesIO(minimal_xlsx())) as z:
            entries = [(name, z.read(name)) for name in z.namelist()]
        limit = max(len(data) for name, data in entries)+10
        with mock.patch.object(self.service, "MAX_MEMBER_BYTES", limit):
            self.assert_bad_archive(upload("bad.xlsx", zip_bytes(entries+[("extra.bin", b"x"*(limit+1))])), "单个文件过大")

    def test_cumulative_expansion_budget_enforced(self):
        with zipfile.ZipFile(io.BytesIO(minimal_xlsx())) as z:
            expanded = sum(info.file_size for info in z.infolist())
        with mock.patch.object(self.service, "MAX_BATCH_EXPANDED_BYTES", expanded*2-1):
            self.assert_bad_archive(upload("source.xlsx", minimal_xlsx()), "Excel 数据过多")

    def test_nonfinite_or_boolean_supplement_rejected(self):
        for value in (float("nan"), float("inf"), True):
            with self.subTest(value=str(value)), self.assertRaisesRegex(ValueError, "有效数字"):
                self.service.preview_update(self.target, [upload("source.xlsx", minimal_xlsx())], {9: {"风电": value}})

    def test_invalid_workbook_bytes_rejected(self):
        self.assert_bad_archive(upload("bad.xlsx", b"not a workbook"))


class CloudWorkbookTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.service = importlib.import_module("service")
        cls.root, cls.target, cls.sources = fixture_uploads()
        cls.target_path = cls.root / cls.target["name"]
        cls.original_hash = hashlib.sha256(cls.target_path.read_bytes()).hexdigest()

    @classmethod
    def tearDownClass(cls):
        if hasattr(cls, "target_path"):
            assert hashlib.sha256(cls.target_path.read_bytes()).hexdigest() == cls.original_hash, "Private original workbook was modified"

    def test_inspect_and_uppercase_xlsx(self):
        sources = [dict(item, name=item["name"][:-5]+".XLSX") for item in self.sources]
        result = self.service.inspect_inputs(self.target, sources)
        self.assertEqual(result["months"], [9])
        self.assertEqual(result["file_count"], 8)
        self.assertEqual(result["record_count"], 20)

    def test_nested_month_zip_and_mac_metadata(self):
        entries = [("结算/"+item["name"], item["data"]) for item in self.sources]
        entries += [("__MACOSX/._one.xlsx", b"metadata"), ("结算/._two.xlsx", b"metadata")]
        result = self.service.inspect_inputs(self.target, [upload("月结单.zip", zip_bytes(entries))])
        self.assertEqual(result["file_count"], 8)
        self.assertEqual(result["record_count"], 20)

    def test_preview_public_has_no_temp_paths(self):
        result = self.service.preview_update(self.target, self.sources)
        self.assertFalse(result["errors"], result["errors"])
        public = json.dumps(result, ensure_ascii=False)
        self.assertNotIn(str(self.root), public)
        self.assertNotIn("TemporaryDirectory", public)
        self.assertNotIn("AppData\\Local\\Temp", public)
        self.assertNotIn("/tmp/", public)
        self.assertEqual(result["months"], [9])

    def test_fingerprint_changes_with_name_bytes_and_inputs(self):
        original = self.service.preview_update(self.target, self.sources)
        renamed = dict(self.target, name="renamed.xlsx")
        renamed_plan = self.service.preview_update(renamed, self.sources)
        self.assertNotEqual(original["fingerprint"], renamed_plan["fingerprint"])
        supplements = {9: {"风电": 358.276, "光伏": 306.43}}
        supplied = self.service.preview_update(self.target, self.sources, supplements)
        self.assertNotEqual(original["fingerprint"], supplied["fingerprint"])
        changed_buffer = io.BytesIO(self.target["data"])
        with zipfile.ZipFile(changed_buffer, "a") as archive:
            archive.comment = b"content fingerprint regression"
        changed_target = dict(self.target, data=changed_buffer.getvalue())
        changed_plan = self.service.preview_update(changed_target, self.sources)
        self.assertNotEqual(original["fingerprint"], changed_plan["fingerprint"])
        with self.assertRaises(ValueError):
            self.service.generate_update(changed_target, self.sources, None, original["fingerprint"])
        with self.assertRaises(ValueError):
            self.service.generate_update(self.target, self.sources, supplements, original["fingerprint"])

    def test_generate_roundtrip_is_idempotent_and_reports_are_downloadable(self):
        plan = self.service.preview_update(self.target, self.sources)
        self.assertFalse(plan["errors"], plan["errors"])
        result = self.service.generate_update(self.target, self.sources, None, plan["fingerprint"])
        self.assertTrue(zipfile.is_zipfile(io.BytesIO(result["workbook_bytes"])))
        self.assertTrue(zipfile.is_zipfile(io.BytesIO(result["bundle_bytes"])))
        self.assertEqual(result["fingerprint"], plan["fingerprint"])
        json.loads(result["report_json"].decode("utf-8-sig"))
        self.assertIn("月份", result["report_csv"].decode("utf-8-sig"))
        again = self.service.preview_update(upload(result["workbook_name"], result["workbook_bytes"]), self.sources)
        self.assertFalse(again["errors"], again["errors"])
        self.assertEqual(again["summary"]["changed_cells"], 0)

    def test_new_october_requires_inputs_then_generates(self):
        sources = [upload(item["name"].replace("2026-09", "2026-10"), month10_copy(item["data"])) for item in self.sources]
        missing = self.service.preview_update(self.target, sources)
        self.assertTrue(missing["errors"])
        with self.assertRaises(ValueError):
            self.service.generate_update(self.target, sources, None, missing["fingerprint"])
        inputs = {10: {"风电": 350.1, "光伏": 280.2, "分布式电量": 0, "分布式电费": 0}}
        plan = self.service.preview_update(self.target, sources, inputs)
        self.assertFalse(plan["errors"], plan["errors"])
        self.assertEqual(plan["summary"]["new_rows"], 21)
        result = self.service.generate_update(self.target, sources, inputs, plan["fingerprint"])
        again = self.service.preview_update(upload(result["workbook_name"], result["workbook_bytes"]), sources, inputs)
        self.assertFalse(again["errors"], again["errors"])
        self.assertEqual(again["summary"]["new_rows"], 0)
        self.assertEqual(again["summary"]["changed_cells"], 0)

    def test_parallel_sessions_are_isolated(self):
        def preview(label):
            return self.service.preview_update(dict(self.target, name=label+".xlsx"), self.sources)
        with ThreadPoolExecutor(max_workers=2) as pool:
            first, second = list(pool.map(preview, ("会话甲", "会话乙")))
        self.assertFalse(first["errors"])
        self.assertFalse(second["errors"])
        self.assertNotEqual(first["fingerprint"], second["fingerprint"])
        self.assertEqual(first["summary"], second["summary"])
        with self.assertRaises(ValueError):
            self.service.generate_update(dict(self.target, name="会话乙.xlsx"), self.sources, None, first["fingerprint"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
