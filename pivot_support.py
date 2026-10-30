"""Keep native PivotTable caches in step with the source data.

This module patches XML fragments instead of reserializing complete parts.  In
particular, Office extension namespaces, slicers and existing item indices stay
untouched.  Excel remains responsible for refreshing arbitrary pivot layouts.
"""
from __future__ import annotations

import ast
import math
import posixpath
import re
import xml.etree.ElementTree as ET
from xml.sax.saxutils import quoteattr

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"m": MAIN}


def _local(tag):
    return tag.rsplit("}", 1)[-1]


def _key(value):
    if value is None or value == "":
        return ("m", None)
    if isinstance(value, bool):
        return ("b", "1" if value else "0")
    if isinstance(value, (int, float)):
        if not math.isfinite(float(value)):
            raise ValueError("透视缓存不能包含无穷值或非数值 NaN")
        return ("n", float(value))
    return ("s", str(value))


def _node_key(node):
    kind = _local(node.tag)
    value = node.attrib.get("v")
    if kind == "n":
        return (kind, float(value))
    if kind == "m":
        return (kind, None)
    return (kind, value)


def _value(node):
    kind, value = _node_key(node)
    if kind == "n":
        return float(value)
    if kind == "m":
        return None
    if kind == "b":
        return value == "1"
    return value


def _item_xml(key):
    kind, value = key
    if kind == "m":
        return "<m/>"
    if kind == "n":
        value = format(value, ".16g")
    return "<" + kind + " v=" + quoteattr(str(value)) + "/>"


def _set_attrs(xml, tag, **attrs):
    def replace(match):
        opening = match.group(0)
        for key, value in attrs.items():
            expression = rf'\s{re.escape(key)}="[^"]*"'
            new = " " + key + "=" + quoteattr(str(value))
            if re.search(expression, opening):
                opening = re.sub(expression, lambda _: new, opening, count=1)
            else:
                opening = opening[:-1] + new + ">"
        return opening
    return re.sub(r"<" + re.escape(tag) + r"\b[^>]*>", replace, xml, count=1)


def _cache_rebuild(xml, rows):
    tree = ET.fromstring(xml)
    fields = list(tree.find("m:cacheFields", NS))
    source_fields = [f for f in fields if f.attrib.get("databaseField", "1") != "0"]
    all_maps, fragments, appended = [], [], {}
    field_fragments = list(re.finditer(r"<cacheField\b[^>]*(?:/>|>.*?</cacheField>)", xml, re.S))
    if len(fields) != len(field_fragments):
        raise ValueError("无法匹配透视缓存字段，已停止更新该缓存")
    for index, field in enumerate(source_fields):
        shared = field.find("m:sharedItems", NS)
        if shared is None:
            raise ValueError("目标透视缓存缺少 sharedItems")
        items = list(shared)
        lookup = {}
        for old_index, item in enumerate(items):
            lookup.setdefault(_node_key(item), old_index)
        extras = []
        for row in rows:
            key = _key(row.get(field.attrib["name"]))
            if key not in lookup:
                lookup[key] = len(items) + len(extras)
                extras.append(key)
        all_maps.append(lookup)
        appended[index] = list(range(len(items), len(items) + len(extras)))
        old_fragment = field_fragments[index].group(0)
        shared_match = re.search(r"<sharedItems\b[^>]*(?:/>|>.*?</sharedItems>)", old_fragment, re.S)
        if not shared_match:
            raise ValueError("透视缓存共享值格式无法识别")
        previous = shared_match.group(0)
        opening = previous[:previous.index(">") + 1]
        opening = re.sub(r"/\s*>$", ">", opening)
        new_keys = list(lookup)
        kinds = {key[0] for key in new_keys}
        # Preserve old entries verbatim, including unused-item metadata.
        old_inner = "" if previous.endswith("/>") else previous[previous.index(">") + 1:previous.rindex("</sharedItems>")]
        opening = _set_attrs(opening, "sharedItems", count=len(items) + len(extras),
                             containsBlank=int("m" in kinds), containsNumber=int("n" in kinds),
                             containsString=int("s" in kinds),
                             containsMixedTypes=int(len(kinds - {"m"}) > 1),
                             containsNonDate=int(bool(kinds - {"d", "m"})))
        numbers = [key[1] for key in new_keys if key[0] == "n"]
        if numbers:
            opening = _set_attrs(opening, "sharedItems", minValue=format(min(numbers), ".16g"), maxValue=format(max(numbers), ".16g"))
        replacement = opening + old_inner + "".join(_item_xml(key) for key in extras) + "</sharedItems>"
        fragments.append((field_fragments[index].start(), field_fragments[index].end(),
                          old_fragment[:shared_match.start()] + replacement + old_fragment[shared_match.end():]))
    for start, end, replacement in reversed(fragments):
        xml = xml[:start] + replacement + xml[end:]
    xml = _set_attrs(xml, "pivotCacheDefinition", recordCount=len(rows), refreshOnLoad="1", enableRefresh="1")
    records = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
               '<pivotCacheRecords xmlns="' + MAIN + '" count="' + str(len(rows)) + '">']
    for row in rows:
        records.append("<r>" + "".join('<x v="' + str(all_maps[i][_key(row.get(field.attrib["name"]))]) + '"/>'
                                             for i, field in enumerate(source_fields)) + "</r>")
    records.append("</pivotCacheRecords>")
    return xml.encode("utf-8"), "".join(records).encode("utf-8"), fields, appended


def _sum_values(rows, name):
    return sum(row[name] for row in rows if isinstance(row.get(name), (int, float)))


def _evaluate(formula, sums):
    # Cache formulas contain field names, arithmetic, and optional quote marks.
    expression = formula
    for index, name in enumerate(sorted(sums, key=len, reverse=True)):
        expression = expression.replace("'" + name + "'", "(" + repr(sums[name]) + ")")
        expression = expression.replace(name, "(" + repr(sums[name]) + ")")
    parsed = ast.parse(expression, mode="eval")
    allowed = (ast.Expression, ast.BinOp, ast.UnaryOp, ast.Constant, ast.Add,
               ast.Sub, ast.Mult, ast.Div, ast.USub, ast.UAdd, ast.Load)
    if any(not isinstance(node, allowed) for node in ast.walk(parsed)):
        raise ValueError("透视表含有不支持的计算字段")
    value = eval(compile(parsed, "<pivot-calculation>", "eval"), {"__builtins__": {}}, {})
    if not isinstance(value, (int, float)) or not math.isfinite(float(value)):
        raise ValueError("透视表计算字段结果不可用")
    return value


def _patch_cell(xml, address, value):
    pattern = r'<c\b(?=[^>]*\br="' + re.escape(address) + r'")[^>]*(?:/>|>.*?</c>)'
    def replace(match):
        original = match.group(0)
        opening = original[:original.index(">") + 1]
        opening = re.sub(r'\s+t="[^"]*"', "", opening)
        opening = re.sub(r"/\s*>$", ">", opening)
        return opening + "<v>" + format(float(value), ".16g") + "</v></c>"
    result, count = re.subn(pattern, replace, xml, count=1, flags=re.S)
    if count != 1:
        raise ValueError("透视表显示区域的单元格不存在：" + address)
    return result


def _pivot_display(parts, pivot_xml, fields, rows):
    pivot = ET.fromstring(pivot_xml)
    location = pivot.find("m:location", NS)
    row_fields = pivot.find("m:rowFields", NS)
    data_fields = pivot.find("m:dataFields", NS)
    if (location is None or location.attrib.get("ref") != "A3:D6"
            or row_fields is None or len(row_fields) != 1
            or data_fields is None or len(data_fields) != 3):
        raise ValueError("透视表布局已变化，显示区域需在 Excel/WPS 中刷新")
    if [int(field.attrib["fld"]) for field in data_fields] != [48, 83, 59]:
        raise ValueError("透视表指标已变化，显示区域需在 Excel/WPS 中刷新")
    pivot_fields = list(pivot.find("m:pivotFields", NS))
    filtered = rows
    for field, pivot_field in zip(fields, pivot_fields):
        items = pivot_field.find("m:items", NS)
        shared = field.find("m:sharedItems", NS)
        if items is None or shared is None:
            continue
        hidden = [item for item in items if item.attrib.get("h") == "1" and "x" in item.attrib]
        if not hidden:
            continue
        values = list(shared)
        excluded = {_key(_value(values[int(item.attrib["x"])])) for item in hidden}
        name = field.attrib["name"]
        if all(key[0] == "m" for key in excluded):
            filtered = [row for row in filtered if _key(row.get(name))[0] != "m"]
        else:
            allowed = {_key(_value(values[int(item.attrib["x"])])) for item in items
                       if "x" in item.attrib and item.attrib.get("h") != "1"}
            filtered = [row for row in filtered if _key(row.get(name)) in allowed]
    row_field_index = int(row_fields[0].attrib["x"])
    grouping_field = fields[row_field_index]
    grouping_shared = list(grouping_field.find("m:sharedItems", NS))
    grouping_items = list(pivot_fields[row_field_index].find("m:items", NS))
    row_items = pivot.find("m:rowItems", NS)
    if row_items is None or len(row_items) != 3:
        raise ValueError("透视表行分组已变化，显示区域需在 Excel/WPS 中刷新")
    groups = []
    for item in list(row_items)[:2]:
        position = int(item.find("m:x", NS).attrib.get("v", "0"))
        shared_index = int(grouping_items[position].attrib["x"])
        groups.append(_value(grouping_shared[shared_index]))
    known_names = {field.attrib["name"] for field in fields if field.attrib.get("databaseField", "1") != "0"}
    results = {}
    for row_number, group in [(4, groups[0]), (5, groups[1]), (6, None)]:
        selected = filtered if group is None else [row for row in filtered if row.get(grouping_field.attrib["name"]) == group]
        sums = {name: _sum_values(selected, name) for name in known_names}
        for column, data_field in zip("BCD", data_fields):
            field = fields[int(data_field.attrib["fld"])]
            formula = field.attrib.get("formula")
            if formula:
                value = _evaluate(formula, sums)
            elif data_field.attrib.get("subtotal", "sum") == "sum":
                value = sums[field.attrib["name"]]
            else:
                raise ValueError("透视表聚合方式已变化，显示区域需在 Excel/WPS 中刷新")
            results[column + str(row_number)] = value
    # Find the worksheet containing this specific native pivot.
    pivot_name = next((name for name, payload in parts.items() if name.startswith("xl/pivotTables/") and payload == pivot_xml), None)
    if pivot_name is None:
        raise ValueError("无法定位透视表工作表")
    sheet_path = None
    for rel_path, payload in parts.items():
        if not rel_path.startswith("xl/worksheets/_rels/") or not rel_path.endswith(".rels"):
            continue
        rel_tree = ET.fromstring(payload)
        for relation in rel_tree:
            if not relation.attrib.get("Type", "").endswith("/pivotTable"):
                continue
            target = relation.attrib["Target"]
            resolved = posixpath.normpath(posixpath.join("xl/worksheets", target)) if not target.startswith("/") else target.lstrip("/")
            if resolved == pivot_name:
                sheet_path = "xl/worksheets/" + posixpath.basename(rel_path)[:-5]
                break
    if sheet_path not in parts:
        raise ValueError("无法定位透视表显示工作表")
    xml = parts[sheet_path].decode("utf-8")
    for address, value in results.items():
        xml = _patch_cell(xml, address, value)
    return {sheet_path: xml.encode("utf-8")}


def refresh_pivots(parts: dict[str, bytes], rows: list[dict]) -> tuple[dict[str, bytes], list[str]]:
    """Return changed native cache/display parts without changing ``parts``.

    ``rows`` contains all data rows keyed by the exact source header names.
    Unknown pivot layouts retain their old display but receive a refreshed
    cache and an explicit warning requesting a native Excel/WPS refresh.
    """
    changed, warnings = {}, []
    for name, payload in parts.items():
        if not name.startswith("xl/pivotCache/pivotCacheDefinition") or not name.endswith(".xml"):
            continue
        tree = ET.fromstring(payload)
        source = tree.find("m:cacheSource/m:worksheetSource", NS)
        if source is None or source.attrib.get("sheet") != "数据表":
            continue
        rebuilt, records, fields, appended = _cache_rebuild(payload.decode("utf-8"), rows)
        suffix = re.search(r"Definition(\d+)\.xml$", name).group(1)
        records_name = "xl/pivotCache/pivotCacheRecords" + suffix + ".xml"
        if records_name not in parts:
            raise ValueError("未找到数据表的透视缓存记录文件")
        changed[name] = rebuilt
        changed[records_name] = records
        # Cache indices are retained.  Native refresh updates numeric item
        # lists and slicer options, avoiding destructive re-authoring.
        if appended.get(3):
            warnings.append("已扩展透视缓存；新增月份的切片器选项将在 Excel/WPS 刷新后显示。")
        workbook = ET.fromstring(parts["xl/workbook.xml"])
        cache_id = None
        relations = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
        relation_id = None
        for relation in relations:
            target = relation.attrib.get("Target", "")
            resolved = posixpath.normpath(posixpath.join("xl", target)) if not target.startswith("/") else target.lstrip("/")
            if resolved == name:
                relation_id = relation.attrib["Id"]
        for cache in workbook.findall("m:pivotCaches/m:pivotCache", NS):
            if cache.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id") == relation_id:
                cache_id = cache.attrib["cacheId"]
        for pivot_name, pivot_payload in parts.items():
            if not pivot_name.startswith("xl/pivotTables/") or not re.search(r"pivotTable\d+\.xml$", pivot_name):
                continue
            pivot = ET.fromstring(pivot_payload)
            if pivot.attrib.get("cacheId") != cache_id:
                continue
            try:
                changed.update(_pivot_display(parts, pivot_payload, fields, rows))
            except (ValueError, ZeroDivisionError, SyntaxError, KeyError, IndexError) as exc:
                warnings.append(str(exc) + "；缓存已更新并设置打开时刷新。")
    return changed, warnings
