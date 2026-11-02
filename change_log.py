"""Append a readable audit worksheet without rebuilding existing Excel objects.

Only this worksheet, its registration and (on first use) appended styles change.
Existing OOXML namespaces, pivot/slicer parts and historical log rows stay intact.
"""
from __future__ import annotations

from datetime import datetime, timedelta, timezone
import io
import math
import posixpath
import re
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape, quoteattr
import zipfile


MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'m': MAIN}
TITLE = '月度电价分析表变更明细'
HEADERS = ('更新日期', '月份', '场站', '单元格', '字段', '更新前', '更新后', '来源')
HEADER_ROW = 4
MAX_ROW = 1048576
BEIJING = timezone(timedelta(hours=8))
_ILLEGAL_XML = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]')


def _text(value):
    value = _ILLEGAL_XML.sub('', str(value))
    # Excel's cell limit is measured in UTF-16 code units.
    if len(value.encode('utf-16-le')) // 2 > 32767:
        raise ValueError('变更明细包含超过32767字符的单元格，请缩短后重试。')
    return escape(value)


def _prefix(xml, local):
    match = re.search(r'<((?:[A-Za-z_][\w.-]*:)?' + re.escape(local) + r')(?=[\s/>])', xml)
    if match is None:
        raise ValueError('工作簿XML缺少必要节点：' + local)
    return match[1][:-len(local)]


def _append(xml, local, fragment, *, count=None):
    """Append inside an existing element, retaining every original namespace."""
    prefix = _prefix(xml, local)
    tag = re.escape(prefix + local)
    # A single explicit match handles both a self-closing and paired container.
    match = re.search(r'<'+tag+r'\b[^>]*(?:/\s*>|>.*?</'+tag+r'\s*>)', xml, re.S)
    if match is None:
        raise ValueError('无法追加工作簿节点：' + local)
    element = match[0]
    opening = re.match(r'<'+tag+r'\b[^>]*>', element)[0]
    if count is not None:
        if re.search(r'\bcount\s*=', opening):
            opening = re.sub(r'\bcount\s*=\s*([\"\']).*?\1', f'count="{count}"', opening, count=1)
        else:
            opening = re.sub(r'\s*/?>$', f' count="{count}">', opening)
    if opening.rstrip().endswith('/>'):
        opening = re.sub(r'/\s*>$', '>', opening)
        replacement = opening + fragment + '</' + prefix + local + '>'
    else:
        body = element[element.find('>')+1:element.rfind('</')]
        replacement = opening + body + fragment + '</' + prefix + local + '>'
    return xml[:match.start()] + replacement + xml[match.end():]


def _new_id(ids):
    number = 1
    while f'rId{number}' in ids:
        number += 1
    return f'rId{number}'


def _part(target):
    return target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)


def _cell_value(cell, strings):
    if cell is None:
        return None
    if cell.get('t') == 'inlineStr':
        return ''.join(cell.find('m:is', NS).itertext())
    value = cell.find('m:v', NS)
    if value is None:
        return None
    if cell.get('t') == 's':
        return strings[int(value.text)]
    return value.text


def _styles(xml):
    """Append new records to style collections; all existing indices stay fixed."""
    root = ET.fromstring(xml)
    prefix = _prefix(xml, 'styleSheet')
    fonts = root.find('m:fonts', NS)
    fills = root.find('m:fills', NS)
    xfs = root.find('m:cellXfs', NS)
    if fonts is None or fills is None or xfs is None:
        raise ValueError('工作簿样式表不完整，无法添加变更明细。')
    font_index, fill_index, xf_index = len(fonts), len(fills), len(xfs)
    formats = root.find('m:numFmts', NS)
    fmt_ids = {int(node.get('numFmtId')) for node in formats} if formats is not None else set()
    date_id = max({163, *fmt_ids}) + 1
    number_id = date_id + 1

    def tags(text):
        return re.sub(r'<(/?)([A-Za-z][A-Za-z0-9]*)', lambda m: '<'+m[1]+prefix+m[2], text)

    format_xml = tags(f'<numFmt numFmtId="{date_id}" formatCode="yyyy-mm-dd hh:mm:ss"/>'
                      f'<numFmt numFmtId="{number_id}" formatCode="#,##0.########"/>')
    if formats is None:
        # numFmts precedes fonts in the OOXML stylesheet sequence.
        match = re.search(r'<'+re.escape(prefix+'fonts')+r'\b', xml)
        xml = xml[:match.start()] + tags('<numFmts count="2">') + format_xml + tags('</numFmts>') + xml[match.start():]
    else:
        xml = _append(xml, 'numFmts', format_xml, count=len(formats)+2)
    font_xml = tags(
        '<font><sz val="10"/><color rgb="FF263445"/><name val="Arial"/></font>'
        '<font><b/><sz val="14"/><color rgb="FF1F4E78"/><name val="Arial"/></font>'
        '<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>'
        '<font><i/><sz val="10"/><color rgb="FF65758B"/><name val="Arial"/></font>')
    xml = _append(xml, 'fonts', font_xml, count=font_index+4)
    xml = _append(xml, 'fills', tags('<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/>'
                                   '<bgColor indexed="64"/></patternFill></fill>'), count=fill_index+1)
    xf_xml = []
    for font, fill, fmt, alignment in (
        (font_index, 0, 0, 'horizontal="left" vertical="center" wrapText="1" indent="1"'),
        (font_index+1, 0, 0, 'vertical="center"'),
        (font_index+2, fill_index, 0, 'horizontal="center" vertical="center"'),
        (font_index+3, 0, 0, 'vertical="center"'),
        (font_index, 0, date_id, 'vertical="center"'),
        (font_index, 0, number_id, 'horizontal="right" vertical="center" indent="1"'),
    ):
        xf_xml.append(tags(f'<xf numFmtId="{fmt}" fontId="{font}" fillId="{fill}" borderId="0" xfId="0" '
                           f'applyNumberFormat="1" applyFont="1" applyFill="1" applyAlignment="1">'
                           f'<alignment {alignment}/></xf>'))
    xml = _append(xml, 'cellXfs', ''.join(xf_xml), count=xf_index+6)
    return xml, dict(body=xf_index, title=xf_index+1, header=xf_index+2,
                     note=xf_index+3, date=xf_index+4, number=xf_index+5)


def _cell(address, value, style):
    if value is None:
        return f'<c r="{address}" s="{style}"/>'
    if isinstance(value, bool):
        return f'<c r="{address}" s="{style}" t="b"><v>{int(value)}</v></c>'
    if isinstance(value, (int, float)):
        if not math.isfinite(value):
            raise ValueError('变更明细含无效数值，无法写入Excel。')
        return f'<c r="{address}" s="{style}"><v>{value}</v></c>'
    # Formula-like source strings remain text, never executable formulas.
    return f'<c r="{address}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{_text(value)}</t></is></c>'


def _rows(changes, first, completed, date1904, styles):
    date = completed.astimezone(BEIJING) if completed.tzinfo is not None else completed
    date = date.replace(tzinfo=None, microsecond=0)
    epoch = datetime(1904, 1, 1) if date1904 else datetime(1899, 12, 30)
    serial = (date - epoch).total_seconds() / 86400
    result = []
    for row, change in enumerate(changes, first):
        values = (serial, change.get('month'), change.get('unit'), change.get('cell'),
                  change.get('field'), change.get('old'), change.get('new'), change.get('source'))
        cells = []
        for column, value in zip('ABCDEFGH', values):
            style = styles['date'] if column == 'A' else styles['number'] if column in 'FG' else styles['body']
            cells.append(_cell(f'{column}{row}', value, style))
        result.append(f'<row r="{row}" ht="30" customHeight="1">'+''.join(cells)+'</row>')
    return ''.join(result)


def _new_sheet(rows, last, styles, no_changes):
    title = _cell('A2', TITLE, styles['title'])
    note = '每次更新自动追加历史记录；更新日期为北京时间，数值单位同原字段。'
    note = _cell('A3', note, styles['note'])
    header = ''.join(_cell(f'{column}4', label, styles['header']) for column, label in zip('ABCDEFGH', HEADERS))
    columns = []
    for index, width in enumerate((23, 8, 34, 12, 30, 22, 22, 78), 1):
        style = styles['date'] if index == 1 else styles['number'] if index in (6, 7) else styles['body']
        columns.append(f'<col min="{index}" max="{index}" width="{width}" style="{style}" customWidth="1"/>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<worksheet xmlns="{MAIN}"><dimension ref="A1:H{last}"/>'
            '<sheetViews><sheetView showGridLines="0" workbookViewId="0">'
            '<pane xSplit="3" ySplit="4" topLeftCell="D5" activePane="bottomRight" state="frozen"/>'
            '<selection pane="bottomRight" activeCell="D5" sqref="D5"/>'
            '</sheetView></sheetViews><sheetFormatPr defaultRowHeight="20"/>'
            '<cols>'+''.join(columns)+'</cols><sheetData><row r="1" ht="8" customHeight="1"/>'
            '<row r="2" ht="26" customHeight="1">'+title+'</row>'
            '<row r="3" ht="22" customHeight="1">'+note+'</row>'
            '<row r="4" ht="24" customHeight="1">'+header+'</row>'+rows+'</sheetData>'
            f'<autoFilter ref="A4:H{last}"/><pageMargins left="0.3" right="0.3" top="0.4" bottom="0.4" header="0.2" footer="0.2"/>'
            '<pageSetup paperSize="9" orientation="landscape"/></worksheet>')


def append_change_log(workbook_bytes: bytes, changes: list[dict], completed: datetime) -> bytes:
    """Append sanitized changes, retaining prior log rows and all unrelated parts."""
    with zipfile.ZipFile(io.BytesIO(workbook_bytes)) as source:
        infos = source.infolist()
        parts = {info.filename: source.read(info) for info in infos}
        archive_comment = source.comment
    workbook_xml = parts['xl/workbook.xml'].decode('utf-8-sig')
    rels_xml = parts['xl/_rels/workbook.xml.rels'].decode('utf-8-sig')
    types_xml = parts['[Content_Types].xml'].decode('utf-8-sig')
    workbook = ET.fromstring(workbook_xml)
    relationships = ET.fromstring(rels_xml)
    sheets = workbook.find('m:sheets', NS)
    if sheets is None:
        raise ValueError('工作簿没有工作表列表，无法添加变更明细。')
    rels = {node.get('Id'): node for node in relationships}
    strings = []
    if 'xl/sharedStrings.xml' in parts:
        strings = [''.join(node.itertext()) for node in ET.fromstring(parts['xl/sharedStrings.xml'])]
    existing = None
    for sheet in sheets:
        relationship = rels.get(sheet.get('{'+REL+'}id'))
        if relationship is None or relationship.get('Type') != REL+'/worksheet':
            continue
        path = _part(relationship.get('Target'))
        tree = ET.fromstring(parts[path])
        cells = {node.get('r'): node for node in tree.findall('m:sheetData/m:row/m:c', NS)}
        if _cell_value(cells.get('A2'), strings) == TITLE and tuple(
                _cell_value(cells.get(f'{column}4'), strings) for column in 'ABCDEFGH') == HEADERS:
            if existing is not None:
                raise ValueError('分析表中有多张相同的自动变更明细，请保留一张后重试。')
            existing = (path, tree)
    if existing is not None and not changes:
        return workbook_bytes

    changed = {}
    props = workbook.find('m:workbookPr', NS)
    date1904 = props is not None and props.get('date1904', '').lower() in ('1', 'true')
    if existing is not None:
        path, tree = existing
        previous_rows = tree.findall('m:sheetData/m:row', NS)
        first = max((int(row.get('r')) for row in previous_rows), default=HEADER_ROW)+1
        columns = tree.findall('m:cols/m:col', NS)
        def column_style(number):
            return next(int(col.get('style')) for col in columns
                        if int(col.get('min')) <= number <= int(col.get('max')))
        try:
            styles = {'body': column_style(3), 'date': column_style(1), 'number': column_style(6)}
        except (StopIteration, TypeError, ValueError) as exc:
            raise ValueError('变更明细列样式已改变，请恢复原表后重试。') from exc
        if first + len(changes)-1 > MAX_ROW:
            raise ValueError('变更明细超过Excel行数上限，请先将历史明细归档后重试。')
        xml = parts[path].decode('utf-8-sig')
        fragment = _rows(changes, first, completed, date1904, styles)
        # Generated audit sheets use the default namespace, but honor a later
        # namespace-prefix rewrite by Excel or another compatible editor.
        prefix = _prefix(xml, 'worksheet')
        if prefix:
            fragment = re.sub(r'<(/?)(row|c|v|is|t)(?=[\s/>])', lambda m: '<'+m[1]+prefix+m[2], fragment)
        xml = _append(xml, 'sheetData', fragment)
        last = first+len(changes)-1
        for local, ref in (('dimension', f'A1:H{last}'), ('autoFilter', f'A4:H{last}')):
            tag = re.escape(prefix+local)
            xml, count = re.subn(r'(<'+tag+r'\b[^>]*\bref\s*=\s*)([\"\']).*?\2',
                                lambda m: m[1]+quoteattr(ref), xml, count=1)
            if count != 1:
                raise ValueError('变更明细表结构已改变，请恢复原表后重试。')
        changed[path] = xml.encode('utf-8')
    else:
        if HEADER_ROW+len(changes) > MAX_ROW:
            raise ValueError('变更明细超过Excel行数上限，请分批更新。')
        style_relationship = next((node for node in relationships if node.get('Type') == REL+'/styles'), None)
        if style_relationship is None:
            style_path, suffix = 'xl/styles.xml', 1
            while style_path in parts:
                style_path = f'xl/settlementStyles{suffix}.xml'
                suffix += 1
            style_xml = (f'<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="{MAIN}">'
                         '<fonts count="1"><font><sz val="11"/><name val="Arial"/></font></fonts>'
                         '<fills count="2"><fill><patternFill patternType="none"/></fill>'
                         '<fill><patternFill patternType="gray125"/></fill></fills>'
                         '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
                         '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
                         '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
                         '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
            style_rid = _new_id(rels)
            rels[style_rid] = None
            prefix = _prefix(rels_xml, 'Relationships')
            rels_xml = _append(rels_xml, 'Relationships',
                f'<{prefix}Relationship Id="{style_rid}" Type="{REL}/styles" Target={quoteattr(posixpath.relpath(style_path, "xl"))}/>')
            prefix = _prefix(types_xml, 'Types')
            types_xml = _append(types_xml, 'Types',
                f'<{prefix}Override PartName={quoteattr("/"+style_path)} ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
        else:
            style_path = _part(style_relationship.get('Target'))
            style_xml = parts[style_path].decode('utf-8-sig')
        style_xml, styles = _styles(style_xml)
        changed[style_path] = style_xml.encode('utf-8')
        names = {sheet.get('name', '').casefold() for sheet in sheets}
        name, suffix = '变更明细', 1
        while name.casefold() in names:
            name = f'变更明细（{suffix}）'
            suffix += 1
        path, suffix = 'xl/worksheets/settlementChanges.xml', 1
        while path in parts:
            path = f'xl/worksheets/settlementChanges{suffix}.xml'
            suffix += 1
        sheet_id = max((int(sheet.get('sheetId')) for sheet in sheets), default=0)+1
        rid = _new_id(rels)
        prefix = _prefix(workbook_xml, 'workbook')
        workbook_xml = _append(workbook_xml, 'sheets',
            f'<{prefix}sheet xmlns:r="{REL}" name={quoteattr(name)} sheetId="{sheet_id}" r:id="{rid}"/>')
        prefix = _prefix(rels_xml, 'Relationships')
        rels_xml = _append(rels_xml, 'Relationships',
            f'<{prefix}Relationship Id="{rid}" Type="{REL}/worksheet" Target={quoteattr(posixpath.relpath(path, "xl"))}/>')
        prefix = _prefix(types_xml, 'Types')
        types_xml = _append(types_xml, 'Types',
            f'<{prefix}Override PartName={quoteattr("/"+path)} ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
        changed['xl/workbook.xml'] = workbook_xml.encode('utf-8')
        changed['xl/_rels/workbook.xml.rels'] = rels_xml.encode('utf-8')
        changed['[Content_Types].xml'] = types_xml.encode('utf-8')
        fragment = _rows(changes, HEADER_ROW+1, completed, date1904, styles)
        changed[path] = _new_sheet(fragment, HEADER_ROW+len(changes), styles, not changes).encode('utf-8')

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, 'w', compression=zipfile.ZIP_DEFLATED) as destination:
        destination.comment = archive_comment
        for info in infos:
            destination.writestr(info, changed.pop(info.filename, parts[info.filename]))
        for path, data in changed.items():
            destination.writestr(path, data)
    return buffer.getvalue()
