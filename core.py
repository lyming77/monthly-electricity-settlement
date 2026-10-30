"""2026 月度结算导入。仅用标准库，定点修改 OOXML，保留原有工作簿对象。"""
from __future__ import annotations

import argparse
import csv
import hashlib
import io
import json
import math
import os
from pathlib import Path
import posixpath
import re
import shutil
import tempfile
from datetime import datetime, timedelta
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape, quoteattr
import zipfile

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
ALIASES = {'招远市龙源新能源有限公司(发电)': '禄山龙源（发电）',
           '招远市龙源新能源有限公司(用电)': '禄山龙源（用电）',
           '玉山龙源第九期第一批': '玉山龙源第三期第二批'}
STORAGE = {
    'R': ('实时用电量', -1), 'S': ('实时市场电价', 1), 'T': ('实时电能量电费', -1),
    'U': ('日前出清电量', -1), 'V': ('日前市场出清电价', 1), 'W': ('日前电能量电费', -1),
    'X': ('中长期合约电量', -1), 'Y': ('中长期合约电价', 1), 'Z': ('中长期合约电费', -1),
    'AA': ('月末调整电量', -1), 'AB': ('月末调整均价', 1), 'AC': ('月末调整电费', -1),
    'AG': ('运行成本补偿费用', -1), 'AH': ('运行成本补偿费用分摊', -1),
    'AK': ('市场超额收益回收返还', -1), 'AT': ('用户侧中长期偏差收益回收', -1),
    'AM': ('调频辅助服务费用分摊', -1), 'AP': ('爬坡辅助服务费用分摊', -1),
    'BB': ('跨月追退补结算电量', -1), 'BC': ('跨月追退补结算电费', -1),
    'BD': ('退补调整分配费用', -1), 'BE': ('清算费用', -1),
    'BF': ('总结算电量', -1), 'BG': ('总结算均价', 1), 'BH': ('总结算费用', -1),
}


def normal(value):
    return re.sub(r'\s+', '', str(value or '')).replace('（', '(').replace('）', ')')


def colnum(col):
    result = 0
    for ch in col:
        result = result * 26 + ord(ch) - 64
    return result


def colname(number):
    result = ''
    while number:
        number, rem = divmod(number - 1, 26)
        result = chr(65 + rem) + result
    return result


def number(value):
    if value is None or value == '' or value in ('/', '-', '—'):
        return None
    try:
        result = float(value)
        if not math.isfinite(result):
            raise ValueError('数值必须有限')
        return result
    except (ValueError, TypeError):
        return None


def filehash(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def same(old, new):
    if isinstance(old, (int, float)) and isinstance(new, (int, float)):
        return math.isclose(old, new, rel_tol=1e-12, abs_tol=1e-8)
    return old == new


def shift_formula(formula, delta, column_delta=0):
    # This template uses A1 row references. Do not shift absolute row references or quoted strings.
    chunks = re.split(r'("[^"]*(?:""[^"]*)*")', formula)
    for i in range(0, len(chunks), 2):
        def translate(m):
            col = m[2] if m[1] else colname(colnum(m[2]) + column_delta)
            return m[1] + col + m[3] + str(int(m[4]) + (0 if m[3] else delta))
        chunks[i] = re.sub(r'(?<![A-Za-z0-9_])([$]?)([A-Z]{1,3})([$]?)([0-9]+)', translate, chunks[i])
    return ''.join(chunks)


class Book:
    def __init__(self, path):
        self.path = Path(path).resolve()
        with zipfile.ZipFile(self.path) as z:
            self.parts = {info.filename: z.read(info) for info in z.infolist()}
            self.infos = z.infolist()
        self.strings = []
        if 'xl/sharedStrings.xml' in self.parts:
            self.strings = [''.join(x.itertext()) for x in ET.fromstring(self.parts['xl/sharedStrings.xml'])]
        rels = {x.attrib['Id']: x.attrib['Target'] for x in ET.fromstring(self.parts['xl/_rels/workbook.xml.rels'])}
        self.sheets = {}
        for sh in ET.fromstring(self.parts['xl/workbook.xml']).find('m:sheets', NS):
            target = rels[sh.attrib['{' + RNS + '}id']]
            part = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
            tree = ET.fromstring(self.parts[part])
            cells, masters = {}, {}
            for c in tree.findall('m:sheetData/m:row/m:c', NS):
                a = c.attrib['r']
                v, f = c.find('m:v', NS), c.find('m:f', NS)
                value = v.text if v is not None else None
                typ = c.attrib.get('t')
                if typ == 's':
                    value = self.strings[int(value)] if value is not None else None
                elif typ == 'inlineStr':
                    value = ''.join(c.find('m:is', NS).itertext())
                elif value is not None and typ not in ('e', 'str'):
                    value = number(value)
                formula = f.text if f is not None else None
                if f is not None and f.attrib.get('t') == 'shared' and formula:
                    masters[f.attrib['si']] = (a, formula)
                cells[a] = {'v': value, 'f': formula, 'style': c.attrib.get('s'),
                            'shared': f.attrib.get('si') if f is not None else None}
            for a, c in cells.items():
                if c['shared'] and not c['f']:
                    base, formula = masters[c['shared']]
                    c['f'] = shift_formula(formula,
                        int(re.search(r'\d+', a)[0]) - int(re.search(r'\d+', base)[0]),
                        colnum(re.sub(r'\d+', '', a)) - colnum(re.sub(r'\d+', '', base)))
            self.sheets[sh.attrib['name']] = {'path': part, 'cells': cells, 'tree': tree}


def parse_source(path):
    book = Book(path)
    if '结算依据' not in book.sheets:
        raise ValueError(f'{Path(path).name}：没有“结算依据”工作表，请选择原始结算单。')
    cells = book.sheets['结算依据']['cells']
    title = str(cells.get('A1', {}).get('v', ''))
    period = re.search(r'(20\d{2})年\s*(\d{1,2})月', title)
    if not period:
        raise ValueError(f'{Path(path).name}：无法从结算单标题识别年月。')
    year, month = map(int, period.groups())
    if year != 2026 or not 1 <= month <= 12:
        raise ValueError(f'{Path(path).name}：本程序用于 2026 年，此文件期间为 {year}年{month}月。')
    header_row = next((r for r in range(1, 13)
                       if {'总结算电量'} <= {c.get('v') for a, c in cells.items()
                                              if int(re.search(r'\d+', a)[0]) == r}
                       and any(c.get('v') in ('总结算电费', '总结算费用') for a, c in cells.items()
                               if int(re.search(r'\d+', a)[0]) == r)), None)
    if header_row is None:
        raise ValueError(f'{Path(path).name}：未找到总结算电费表头，结算单格式可能已经变化。')
    headers = {re.sub(r'\d+', '', a): normal(c['v']) for a, c in cells.items()
               if int(re.search(r'\d+', a)[0]) == header_row and c['v'] is not None}
    byrow = {}
    for a, c in cells.items():
        row = int(re.search(r'\d+', a)[0])
        if row > header_row:
            byrow.setdefault(row, {})[re.sub(r'\d+', '', a)] = c['v']
    records, unit = [], None
    for row, values in sorted(byrow.items()):
        label = values.get('A')
        if normal(label) == '合计':
            if not unit:
                raise ValueError(f'{Path(path).name}：第{row}行合计缺少场站名称。')
            data = {headers[c]: v for c, v in values.items() if c in headers and c not in ('A', 'B')}
            side = '用电' if '(用电)' in normal(unit) else '发电'
            required = ('总结算电量', '总结算费用' if side == '用电' else '总结算电费')
            if any(number(data.get(k)) is None for k in required):
                raise ValueError(f'{Path(path).name}：{unit} 的总结算电量/电费缺失，不能当成零。')
            records.append({'month': month, 'year': year, 'unit': unit, 'side': side,
                            'data': data, 'source': str(Path(path).resolve()), 'source_row': row})
        elif label:
            unit = str(label).strip()
        if values.get('B') and normal(label) != '合计':
            date = values['B']
            try:
                actual = datetime(1899, 12, 30) + timedelta(days=float(date)) if isinstance(date, (int, float)) else datetime.strptime(str(date)[:10], '%Y-%m-%d')
            except (ValueError, TypeError, OverflowError):
                continue
            if (actual.year, actual.month) != (year, month):
                raise ValueError(f'{Path(path).name}：明细日期 {actual:%Y-%m-%d} 与标题年月不一致。')
    if not records:
        raise ValueError(f'{Path(path).name}：没有识别到场站合计行。')
    return records


def expand_sources(sources):
    result = []
    for item in sources:
        path = Path(item).resolve()
        result.extend(sorted(path.rglob('*.xlsx')) if path.is_dir() else [path])
    return list(dict.fromkeys(p for p in result if p.suffix.lower() == '.xlsx' and not p.name.startswith('~$')))


def prepare_update(target, sources, reference_prices=None):
    target = Path(target).resolve()
    plan = {'target': str(target), 'months': [], 'records': [], 'changes': [], 'warnings': [], 'errors': [],
            'summary': {}, '_patches': {}, '_new_rows': {}, '_sources': {}, '_target_hash': filehash(target)}
    book = Book(target)
    if '数据表' not in book.sheets:
        raise ValueError('目标文件没有“数据表”工作表，请选择原来的月度电价分析文件。')
    sh = book.sheets['数据表']
    cells = sh['cells']
    headers = {colname(i): cells.get(colname(i) + '2', {}).get('v') for i in range(1, 70)}
    if headers.get('C') != '结算单元名称' or headers.get('BH') != '总结算电费' or headers.get('BQ') != '财政补贴收入':
        raise ValueError('目标表表头已变化，需调整字段配置，已停止写入。')
    rows = {}
    for a, c in cells.items():
        r = int(re.search(r'\d+', a)[0])
        if r >= 3:
            rows.setdefault(r, {})[re.sub(r'\d+', '', a)] = c['v']
    rows = {r: v for r, v in rows.items() if v.get('C') and re.fullmatch(r'\d{1,2}月', str(v.get('D', '')))}
    keyed, templates = {}, {}
    for r, values in rows.items():
        key = (int(values['D'][:-1]), normal(values['C']))
        if key in keyed:
            plan['errors'].append(f'目标表有重复月份/场站：{values["D"]} {values["C"]}。')
        keyed[key] = r
        templates[normal(values['C'])] = r
    aliases = {normal(k): normal(v) for k, v in ALIASES.items()}
    sources = expand_sources(sources)
    sources = [p for p in sources if p != target]
    if not sources:
        plan['errors'].append('没有选择有效的原始结算单。')
    seen = {}
    duplicate_count = 0
    for path in sources:
        try:
            recs = parse_source(path)
            plan['_sources'][str(path)] = filehash(path)
            for rec in recs:
                unit = aliases.get(normal(rec['unit']), normal(rec['unit']))
                rec['target_unit'] = unit
                key = (rec['month'], unit)
                if key in seen:
                    previous = seen[key]
                    if previous['data'] == rec['data']:
                        duplicate_count += 1
                        continue
                    plan['errors'].append(f'{rec["month"]}月 {rec["unit"]} 在多个结算单中数值不同，请只保留最终版本。')
                    continue
                seen[key] = rec
                plan['records'].append(rec)
        except (ValueError, KeyError, zipfile.BadZipFile, ET.ParseError, OSError) as exc:
            plan['errors'].append(str(exc))
    plan['months'] = sorted({r['month'] for r in plan['records']})
    if duplicate_count:
        plan['warnings'].append(f'已忽略 {duplicate_count} 条完全相同的重复场站合计，不会重复累加。')
    refs = {}
    for key, value in (reference_prices or {}).items():
        if str(key).isdigit() and isinstance(value, dict):
            refs[int(key)] = value
    if reference_prices and not refs:
        if len(plan['months']) == 1:
            refs[plan['months'][0]] = reference_prices
        else:
            plan['errors'].append('多个待导入月份的补充数据必须分别按月份提供。')
    for month, vals in refs.items():
        if month not in plan['months']:
            plan['errors'].append(f'补充数据填写了{month}月，但选中的结算单不含该月。')
        for label, value in vals.items():
            if number(value) is None:
                plan['errors'].append(f'{month}月 {label} 不是有效数字。')
        if ('分布式电量' in vals) != ('分布式电费' in vals):
            plan['errors'].append(f'{month}月分布式电量和电费需要同时填写。')
    maxrow = max(rows, default=2)
    work = {r: dict(vals) for r, vals in rows.items()}
    patches = plan['_patches']

    def assign(r, col, val, rec, formula=None):
        old = work.setdefault(r, {}).get(col)
        if not same(old, val):
            plan['changes'].append({'sheet': '数据表', 'cell': col + str(r), 'month': rec['month'],
                                    'unit': work[r].get('C', rec.get('unit', '')), 'field': headers.get(col, col),
                                    'old': old, 'new': val, 'source': rec.get('source', '人工补充/模板公式')})
            work[r][col] = val
            patches[col + str(r)] = {'value': val, 'formula': formula}
        elif formula is not None and cells.get(col + str(r), {}).get('f') != formula:
            # Only update a formula when one of the row's imported values changes, below.
            return

    newmonths = [m for m in plan['months'] if not any(k[0] == m for k in keyed)]
    # Append whole template blocks in chronological order, maintaining the template's unit order.
    for month in newmonths:
        if any(k[0] > month for k in keyed):
            plan['errors'].append(f'目标表缺少{month}月但已有更晚月份，不能在末尾错序插入，请先补齐目标月份。')
            continue
        for template in sorted(templates.values()):
            maxrow += 1
            work[maxrow] = {c: rows[template].get(c) for c in ('A', 'B', 'C', 'E', 'BI', 'BO')}
            work[maxrow]['D'] = f'{month}月'
            keyed[(month, normal(rows[template]['C']))] = maxrow
            plan['_new_rows'][str(maxrow)] = template
            for c, val in work[maxrow].items():
                patches[c + str(maxrow)] = {'value': val, 'formula': None}
    imported = set()
    for rec in plan['records']:
        month, unit, data = rec['month'], rec['target_unit'], rec['data']
        if unit not in templates:
            plan['errors'].append(f'未配置的新场站：{rec["unit"]}。请先确认目标表的公司、分类和补贴规则。')
            continue
        r = keyed.get((month, unit))
        if not r:
            continue
        imported.add((month, unit))
        user = rec['side'] == '用电'
        mapping = STORAGE if user else {colname(i): (normal(headers[colname(i)]), 1) for i in range(6, 61)}
        used = set()
        for col, (field, sign) in mapping.items():
            field = normal(field)
            raw = data.get(field)
            if raw is None:
                continue  # A newly introduced or genuinely absent field must not erase historical data.
            used.add(field)
            val = number(raw)
            if val is None:
                if raw not in ('/', '-', '', None):
                    plan['errors'].append(f'{month}月 {rec["unit"]} {field}不是有效数值：{raw}')
                continue
            val *= sign
            old = work[r].get(col)
            if val == 0 and old in (None, '/') and str(r) not in plan['_new_rows']:
                continue  # Preserve intentional non-applicable placeholders for zero-value fields.
            formula = cells.get(col + str(r), {}).get('f') if col == 'BG' else None
            assign(r, col, val, rec, formula)
        extras = [k for k, v in data.items() if k not in used and number(v) not in (None, 0)
                  and k not in ('结算单元名称', '日期')]
        if extras:
            plan['warnings'].append(f'{month}月 {rec["unit"]} 有目标表无专列的非零项目：{"、".join(extras)}。总结算电费已按源文件完整导入，不重复拆加。')
        if user:
            plan['warnings'].append(f'{month}月储能用电按“支出负、电量负、价格不变”导入，含日前、月末和跨月调整。')
        vals = work[r]
        category = vals['E']
        supplied = refs.get(month, {})
        if category in ('风电', '光伏'):
            if category in supplied and number(supplied[category]) is not None:
                assign(r, 'BJ', float(supplied[category]), rec)
            elif month in newmonths and number(vals.get('BI')) not in (None, 0):
                plan['errors'].append(f'{month}月需要填写{category}机制结算参考价。')
            elif number(vals.get('BI')) not in (None, 0) and number(vals.get('BJ')) is None:
                plan['errors'].append(f'{month}月 {rec["unit"]} 的机制参考价为空，请补充。')
            bi, bj = number(vals.get('BI')), number(vals.get('BJ'))
            qty, export_qty = number(vals.get('BF')), number(vals.get('F'))
            if qty is None:
                plan['errors'].append(f'{month}月 {rec["unit"]} 缺少总结算电量。')
                continue
            enabled = bi not in (None, 0)
            if enabled and bj is not None and export_qty is not None:
                derived = {'BK': (bi-bj, f'BI{r}-BJ{r}'),
                           'BL': (qty*0.8-export_qty, f'BF{r}*0.8-F{r}'),
                           'BM': ((qty*0.8-export_qty)*(bi-bj), f'BL{r}*BK{r}'),
                           'BN': ((qty*0.8-export_qty)*bi, f'BI{r}*BL{r}')}
            elif not enabled:
                derived = {c: (0, None) for c in ('BK', 'BL', 'BM', 'BN')}
            else:
                derived = {}
            subsidy = number(vals.get('BO')) or 0
            ratio = 3/4.75 if unit == normal('冯家龙源第一期第一批') else (1 if subsidy else 0)
            derived['BP'] = (qty*ratio, f'BF{r}*3/4.75' if ratio not in (0, 1) else (f'BF{r}' if ratio else None))
            derived['BQ'] = (qty*ratio*subsidy, f'BO{r}*BP{r}')
            for col, (val, formula) in derived.items():
                assign(r, col, val, rec, formula)
        else:
            for col in ('BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ'):
                assign(r, col, 0, rec)
        if month in newmonths:
            assign(r, 'BD', number(vals.get('BD')) or 0, rec)
    for month in plan['months']:
        expected = {u for u in templates if u != normal('分布式光伏')}
        missing = sorted(expected - {u for m, u in imported if m == month})
        if missing:
            message = f'{month}月未提供 {len(missing)} 个场站的结算单：{"、".join(missing)}。'
            if month in newmonths:
                plan['errors'].append(message + '新增月份需要提供完整场站数据。')
            else:
                plan['warnings'].append(message + '这些场站原值保留。')
        supplied = refs.get(month, {})
        r = keyed.get((month, normal('分布式光伏')))
        if r and '分布式电量' in supplied and '分布式电费' in supplied:
            rec = {'month': month, 'unit': '分布式光伏'}
            qty, fee = number(supplied['分布式电量']), number(supplied['分布式电费'])
            if qty is not None and fee is not None:
                if qty == 0 and fee != 0:
                    plan['warnings'].append(f'{month}月分布式电量为0但电费不为0，电价保留空白。')
                assign(r, 'BF', qty, rec)
                assign(r, 'BH', fee, rec)
                assign(r, 'BG', fee/qty if qty else None, rec, f'IF(BF{r}=0,"",BH{r}/BF{r})')
                assign(r, 'R', qty, rec, f'BF{r}')
                assign(r, 'S', fee/qty if qty else None, rec, f'BG{r}')
                assign(r, 'T', fee, rec, f'BH{r}')
                for col in ('BO', 'BP', 'BQ'):
                    assign(r, col, 0, rec)
        elif month in newmonths:
            plan['errors'].append(f'{month}月需要填写分布式电量和电费（确认为0时请填写0）。')
        else:
            plan['warnings'].append(f'{month}月分布式光伏没有新输入，保留目标表原值。')
    plan['_headers'] = headers
    plan['_rows'] = work
    plan['_sheet_path'] = sh['path']
    plan['_maxrow'] = maxrow
    historical = [c for c in plan['changes'] if c['month'] not in newmonths
                  and isinstance(c['old'], (int, float))
                  and colnum(re.sub(r'\d+', '', c['cell'])) <= 60]
    if historical:
        plan['warnings'].insert(0, f'历史重导含 {len(historical)} 处已有数值变化。点击更新后将按所选结算单替换，请先核对结算版本和变更明细。')
    plan['warnings'] = list(dict.fromkeys(plan['warnings']))
    plan['errors'] = list(dict.fromkeys(plan['errors']))
    plan['summary'] = {'files': len(plan['_sources']), 'records': len(plan['records']),
                       'changed_cells': len(plan['changes']), 'new_rows': len(plan['_new_rows']),
                       'duplicates': duplicate_count, 'months': len(plan['months'])}
    return plan


def cell_xml(address, value, style=None, formula=None):
    attrs = ' r=' + quoteattr(address) + (' s=' + quoteattr(str(style)) if style is not None else '')
    if formula is not None:
        v = '' if value is None else ('<v>' + format(value, '.16g') + '</v>')
        typ = ' t="str"' if isinstance(value, str) else ''
        if isinstance(value, str):
            v = '<v>' + escape(value) + '</v>'
        return '<c' + attrs + typ + '><f>' + escape(formula) + '</f>' + v + '</c>'
    if value is None:
        return '<c' + attrs + '/>'
    if isinstance(value, str):
        return '<c' + attrs + ' t="inlineStr"><is><t>' + escape(value) + '</t></is></c>'
    return '<c' + attrs + '><v>' + format(value, '.16g') + '</v></c>'


def patch_sheet(xml_bytes, patches, cells, new_rows, maxrow):
    """Preserve namespace declarations/unknown extensions by replacing only affected cell XML."""
    text = xml_bytes.decode('utf-8')
    rowpat = re.compile(r'<row\b[^>]*\br="(\d+)"[^>]*>.*?</row>', re.S)
    cellpat = re.compile(r'<c\b[^>]*\br="([A-Z]+\d+)"[^>]*(?:/>|>.*?</c>)', re.S)
    patches = dict(patches)
    # If a shared-formula anchor is replaced, materialize the surviving group members.
    # Otherwise Excel would be left with followers whose anchor no longer exists.
    expand_groups = set()
    for match in cellpat.finditer(text):
        if match[1] not in patches:
            continue
        fm = re.search(r'<f\b[^>]*\bt="shared"[^>]*>(.+?)</f>', match[0], re.S)
        if fm:
            expand_groups.add(cells[match[1]]['shared'])
    for a, cell in cells.items():
        if cell['shared'] in expand_groups and a not in patches:
            patches[a] = {'value': cell['v'], 'formula': cell['f']}
    byrow = {}
    for a, patch in patches.items():
        byrow.setdefault(int(re.search(r'\d+', a)[0]), {})[a] = patch
    def replace_row(match):
        rn = int(match[1])
        if rn not in byrow:
            return match[0]
        raw = match[0]
        existing = {m[1]: m[0] for m in cellpat.finditer(raw)}
        for a, patch in byrow.pop(rn).items():
            existing[a] = cell_xml(a, patch['value'], cells.get(a, {}).get('style'), patch.get('formula'))
        opening = raw[:raw.index('>')+1]
        return opening + ''.join(existing[a] for a in sorted(existing, key=lambda x: colnum(re.sub(r'\d+', '', x)))) + '</row>'
    text = rowpat.sub(replace_row, text)
    added = []
    for rn, rowpatch in sorted(byrow.items()):
        template = int(new_rows[str(rn)])
        body = []
        for i in range(1, 70):
            col = colname(i)
            patch = rowpatch.get(col+str(rn), {'value': None})
            style = cells.get(col+str(template), {}).get('style')
            body.append(cell_xml(col+str(rn), patch['value'], style, patch.get('formula')))
        added.append(f'<row r="{rn}" spans="1:69">' + ''.join(body) + '</row>')
    text = text.replace('</sheetData>', ''.join(added) + '</sheetData>')
    if new_rows:
        text = re.sub(r'(<dimension\b[^>]*\bref=")([^"]+)(")', lambda m: m[1] + 'A1:CC' + str(maxrow) + m[3], text, count=1)
        text = re.sub(r'(<autoFilter\b[^>]*\bref=")([^"]+)(")', lambda m: m[1] + 'A1:BQ' + str(maxrow) + m[3], text, count=1)
    return text.encode('utf-8')


def apply_update(plan, output=None):
    if plan['errors']:
        raise ValueError('仍有待解决问题：\n' + '\n'.join(plan['errors']))
    target = Path(plan['target'])
    output = Path(output).resolve() if output else target
    if output.suffix.lower() != '.xlsx':
        raise ValueError('输出文件必须是 .xlsx。')
    if filehash(target) != plan['_target_hash']:
        raise ValueError('目标文件在预览后发生变化，请重新预览。')
    for path, fingerprint in plan['_sources'].items():
        if filehash(path) != fingerprint:
            raise ValueError(f'结算单在预览后发生变化：{Path(path).name}。请重新预览。')
    if output in [Path(p) for p in plan['_sources']]:
        raise ValueError('输出路径不能覆盖原始结算单。')
    if target.with_name('~$' + target.name).exists() or output.with_name('~$' + output.name).exists():
        raise ValueError('请先关闭 Excel/WPS 中的目标文件，再执行更新。')
    if not plan['_patches']:
        return {'output': str(target), 'backup': None, 'report': None, 'changed_cells': 0, 'warnings': plan['warnings']}
    book = Book(target)
    changed = {plan['_sheet_path']: patch_sheet(book.parts[plan['_sheet_path']], plan['_patches'],
                    book.sheets['数据表']['cells'], plan['_new_rows'], plan['_maxrow'])}
    wbtext = book.parts['xl/workbook.xml'].decode('utf-8')
    wbtext = re.sub(r'<calcPr\b[^>]*/>', '<calcPr calcId="191029" fullCalcOnLoad="1" forceFullCalc="1"/>', wbtext)
    if plan['_new_rows']:
        wbtext = re.sub(r'(数据表!\$A\$1:\$BQ\$)\d+', lambda m: m[1]+str(plan['_maxrow']), wbtext)
    changed['xl/workbook.xml'] = wbtext.encode('utf-8')
    pivotwarnings = []
    try:
        from pivot_support import refresh_pivots
        headers = plan['_headers']
        records = [{headers[c]: v.get(c) for c in headers if headers[c]} for r, v in sorted(plan['_rows'].items())]
        pv, pivotwarnings = refresh_pivots(book.parts | changed, records)
        changed.update(pv)
    except ImportError:
        for name, payload in book.parts.items():
            if re.fullmatch(r'xl/pivotCache/pivotCacheDefinition\d+.xml', name) and b'sheet="\xe6\x95\xb0\xe6\x8d\xae\xe8\xa1\xa8"' in payload:
                txt = re.sub(r'(<pivotCacheDefinition\b)', r'\1 refreshOnLoad="1"', payload.decode('utf-8'), count=1)
                changed[name] = txt.encode('utf-8')
        pivotwarnings.append('已标记透视表打开时刷新；请在 Excel/WPS 中刷新数据。')
    output.parent.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    # Back up the bytes that would be replaced, never overwrite an earlier backup.
    backup = None
    if output.exists():
        backupdir = output.parent / '备份'
        backupdir.mkdir(exist_ok=True)
        backup = backupdir / f'{output.stem}_{stamp}.xlsx'
        shutil.copy2(output, backup)
    fd, tempname = tempfile.mkstemp(prefix='.结算更新_', suffix='.xlsx', dir=output.parent)
    os.close(fd)
    try:
        with zipfile.ZipFile(tempname, 'w', compression=zipfile.ZIP_DEFLATED) as z:
            for info in book.infos:
                z.writestr(info, changed.get(info.filename, book.parts[info.filename]))
        # Structural validation and direct numeric read-back before replacing the real target.
        verified = Book(tempname)
        for a, patch in plan['_patches'].items():
            actual = verified.sheets['数据表']['cells'].get(a, {}).get('v')
            if not same(actual, patch['value']):
                raise ValueError(f'写入校验失败：{a}，原目标文件保留。')
        if filehash(target) != plan['_target_hash']:
            raise ValueError('写入期间目标文件发生变化，已停止替换，请重新预览。')
        os.replace(tempname, output)
    except PermissionError as exc:
        raise PermissionError('文件正在被占用或没有写入权限，请关闭 Excel/WPS 后重新预览。') from exc
    finally:
        if Path(tempname).exists():
            Path(tempname).unlink()
    logdir = output.parent / '更新记录'
    logdir.mkdir(exist_ok=True)
    report = logdir / f'结算更新_{stamp}.json'
    public = {k: v for k, v in plan.items() if not k.startswith('_')}
    public.update({'output': str(output), 'backup': str(backup) if backup else None, 'completed': stamp,
                   'warnings': list(dict.fromkeys(plan['warnings'] + pivotwarnings))})
    report.write_text(json.dumps(public, ensure_ascii=False, indent=2), encoding='utf-8')
    csvpath = report.with_suffix('.csv')
    with csvpath.open('w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['月份', '场站', '单元格', '字段', '更新前', '更新后', '来源'])
        for c in plan['changes']:
            writer.writerow([c['month'], c['unit'], c['cell'], c['field'], c['old'], c['new'], c['source']])
    return {'output': str(output), 'backup': str(backup) if backup else None, 'report': str(report),
            'changed_cells': len(plan['changes']), 'warnings': public['warnings']}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='月度结算助手：默认仅预览，加 --update 才写入。')
    parser.add_argument('--target', required=True)
    parser.add_argument('--sources', nargs='+', required=True)
    parser.add_argument('--inputs', help='UTF-8 JSON：{"10":{"风电":300,"光伏":250,"分布式电量":0,"分布式电费":0}}')
    parser.add_argument('--update', action='store_true')
    parser.add_argument('--output', help='另存为路径，不填则备份并原位更新')
    parser.add_argument('--report', help='预览 JSON 输出路径')
    args = parser.parse_args()
    inputs = json.loads(Path(args.inputs).read_text(encoding='utf-8-sig')) if args.inputs else None
    p = prepare_update(args.target, args.sources, inputs)
    result = {k: v for k, v in p.items() if not k.startswith('_')}
    if args.report:
        Path(args.report).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({k: result[k] for k in ('months', 'summary', 'warnings', 'errors')}, ensure_ascii=False, indent=2))
    if p['errors']:
        raise SystemExit(2)
    if args.update:
        print(json.dumps(apply_update(p, args.output), ensure_ascii=False, indent=2))
