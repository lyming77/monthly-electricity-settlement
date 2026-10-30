"""Isolated upload/preview/download adapter around the existing settlement engine.

No user data is cached globally or persisted to the source directory. Every call
uses its own temporary workspace, and returns only names, values and bytes.
"""
from __future__ import annotations

from contextlib import contextmanager
import csv
from datetime import datetime, timedelta, timezone
import hashlib
import io
import json
import math
from pathlib import Path, PurePosixPath, PureWindowsPath
import re
import stat
import tempfile
from threading import BoundedSemaphore
from typing import TypedDict
import xml.etree.ElementTree as ET
import zipfile

import core


class Upload(TypedDict):
    name: str
    data: bytes


MIB = 1024 * 1024
MAX_UPLOAD_BYTES = 32 * MIB
MAX_TOTAL_UPLOAD_BYTES = 96 * MIB
MAX_WORKBOOK_BYTES = 24 * MIB
MAX_BATCH_EXPANDED_BYTES = 64 * MIB
MAX_ARCHIVE_EXPANDED_BYTES = 96 * MIB
MAX_MEMBER_BYTES = 32 * MIB
MAX_MEMBERS = 4096
MAX_SOURCE_FILES = 160
MAX_COMPRESSION_RATIO = 500
SUPPLEMENT_FIELDS = {'风电', '光伏', '分布式电量', '分布式电费'}
_PROCESSING_SLOT = BoundedSemaphore(1)


def _upload(item, *, target=False):
    if not isinstance(item, dict) or not isinstance(item.get('name'), str):
        raise ValueError('上传文件缺少名称。')
    name, data = item['name'], item.get('data')
    if not name or any(ord(c) < 32 for c in name) or '/' in name or '\\' in name or ':' in name:
        raise ValueError('上传文件名无效，请直接选择 Excel 或 ZIP 文件。')
    allowed = ('.xlsx',) if target else ('.xlsx', '.zip')
    if Path(name).suffix.lower() not in allowed:
        raise ValueError(f'{name}：请上传 {"Excel .xlsx" if target else ".xlsx 或 .zip"} 文件。')
    if not isinstance(data, bytes) or not data:
        raise ValueError(f'{name}：文件为空或无法读取。')
    if len(data) > MAX_UPLOAD_BYTES:
        raise ValueError(f'{name}：单个文件不能超过32 MB。')
    return {'name': name, 'data': data}


def _archive_members(z, label, expanded_limit):
    infos = z.infolist()
    if len(infos) > MAX_MEMBERS:
        raise ValueError(f'{label}：压缩包内文件过多。')
    seen, expanded = set(), 0
    for info in infos:
        path = info.orig_filename
        posix, windows = PurePosixPath(path), PureWindowsPath(path)
        if (not path or any(ord(c) < 32 for c in path) or '\\' in path or posix.is_absolute()
                or windows.drive or '..' in posix.parts or ':' in path):
            raise ValueError(f'{label}：压缩包包含无效或越界路径。')
        canonical = posix.as_posix().rstrip('/').casefold()
        if canonical in seen:
            raise ValueError(f'{label}：压缩包包含重复文件路径。')
        seen.add(canonical)
        if info.flag_bits & 1:
            raise ValueError(f'{label}：不支持加密压缩包，请先解密。')
        if stat.S_IFMT(info.external_attr >> 16) == stat.S_IFLNK:
            raise ValueError(f'{label}：压缩包包含链接文件，无法导入。')
        if info.file_size > MAX_MEMBER_BYTES:
            raise ValueError(f'{label}：压缩包内单个文件过大。')
        if info.file_size > MIB and info.file_size / max(1, info.compress_size) > MAX_COMPRESSION_RATIO:
            raise ValueError(f'{label}：压缩比异常，请拆分文件后上传。')
        expanded += info.file_size
        if expanded > expanded_limit:
            raise ValueError(f'{label}：解压后的数据过大，请分批导入。')
    return infos, expanded


def _read_member(z, info, label):
    with z.open(info) as stream:
        data = stream.read(MAX_MEMBER_BYTES + 1)
    if len(data) > MAX_MEMBER_BYTES or len(data) != info.file_size:
        raise ValueError(f'{label}：压缩包文件大小异常。')
    return data


def _validate_workbook(upload):
    label = upload['name']
    try:
        with zipfile.ZipFile(io.BytesIO(upload['data'])) as z:
            infos, expanded = _archive_members(z, label, MAX_WORKBOOK_BYTES)
            if not {'[Content_Types].xml', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels'} <= set(z.namelist()):
                raise ValueError(f'{label}：不是有效的 Excel 工作簿。')
            for info in infos:
                if info.is_dir():
                    continue
                data = _read_member(z, info, label)
                if info.filename.lower().endswith(('.xml', '.rels')):
                    # Check both UTF-8 and zero-padded UTF-16 before the XML parser.
                    markup = data.replace(b'\x00', b'').upper()
                    if b'<!DOCTYPE' in markup or b'<!ENTITY' in markup:
                        raise ValueError(f'{label}：包含不支持的 XML 外部实体声明。')
            return expanded
    except (zipfile.BadZipFile, RuntimeError, NotImplementedError, EOFError) as exc:
        raise ValueError(f'{label}：Excel 文件损坏或压缩格式不受支持。') from exc


def _materials(target, sources):
    target = _upload(target, target=True)
    if not isinstance(sources, (list, tuple)) or not sources:
        raise ValueError('请先上传当月原始结算单。')
    uploads = [_upload(item) for item in sources]
    if len(uploads) > MAX_SOURCE_FILES or sum(len(x['data']) for x in [target, *uploads]) > MAX_TOTAL_UPLOAD_BYTES:
        raise ValueError('上传数据过多，请按月份分批导入。')
    expanded_workbooks = _validate_workbook(target)
    workbooks = []
    for item in uploads:
        if item['name'].lower().endswith('.xlsx'):
            candidates = [item]
        else:
            candidates = []
            try:
                with zipfile.ZipFile(io.BytesIO(item['data'])) as z:
                    infos, _ = _archive_members(z, item['name'], MAX_ARCHIVE_EXPANDED_BYTES)
                    for info in infos:
                        parts = PurePosixPath(info.filename).parts
                        if info.is_dir() or '__MACOSX' in parts or any(p.startswith('._') for p in parts):
                            continue
                        if not info.filename.lower().endswith('.xlsx') or PurePosixPath(info.filename).name.startswith('~$'):
                            continue
                        candidates.append({'name': item['name'] + '/' + info.filename,
                                           'data': _read_member(z, info, item['name'])})
            except (zipfile.BadZipFile, RuntimeError, NotImplementedError, EOFError) as exc:
                raise ValueError(f'{item["name"]}：无法读取 ZIP 压缩包。') from exc
            if not candidates:
                raise ValueError(f'{item["name"]}：压缩包内没有 .xlsx 结算单。')
        for candidate in candidates:
            if len(candidate['data']) > MAX_UPLOAD_BYTES:
                raise ValueError(f'{candidate["name"]}：文件超过32 MB。')
            expanded_workbooks += _validate_workbook(candidate)
            workbooks.append(candidate)
            if len(workbooks) > MAX_SOURCE_FILES or expanded_workbooks > MAX_BATCH_EXPANDED_BYTES:
                raise ValueError('解压后的 Excel 数据过多，请按月份分批导入。')
    return target, workbooks


def _supplements(values):
    if values is None:
        return {}
    if not isinstance(values, dict):
        raise ValueError('补充数据需要按月份填写。')
    result = {}
    for month, fields in values.items():
        if not str(month).isdigit() or not 1 <= int(month) <= 12 or not isinstance(fields, dict):
            raise ValueError('补充数据月份必须为1至12。')
        parsed = {}
        for label, value in fields.items():
            if label not in SUPPLEMENT_FIELDS:
                raise ValueError(f'无法识别的补充字段：{label}')
            if value is None or value == '':
                continue
            try:
                if isinstance(value, bool):
                    raise ValueError()
                numeric = float(value)
                if not math.isfinite(numeric):
                    raise ValueError()
            except (ValueError, TypeError):
                raise ValueError(f'{int(month)}月 {label}请填写有效数字。') from None
            parsed[label] = numeric
        if parsed:
            if int(month) in result:
                raise ValueError(f'{int(month)}月补充数据重复。')
            result[int(month)] = parsed
    return result


def input_fingerprint(target, sources, supplements=None):
    """Content-sensitive revision shared by preview and generation, independent of temp paths."""
    target = _upload(target, target=True)
    uploads = [_upload(x) for x in sources]
    descriptor = {'target': [target['name'], hashlib.sha256(target['data']).hexdigest()],
                  'sources': sorted([u['name'], hashlib.sha256(u['data']).hexdigest()] for u in uploads),
                  'supplements': _supplements(supplements)}
    return hashlib.sha256(json.dumps(descriptor, ensure_ascii=False, sort_keys=True).encode('utf-8')).hexdigest()


@contextmanager
def _workspace(target, sources):
    # Limit simultaneous XML expansion on small cloud instances; data remains session-local.
    with _PROCESSING_SLOT:
        target, sources = _materials(target, sources)
        with tempfile.TemporaryDirectory(prefix='settlement_') as directory:
            base = Path(directory)
            target_path = base / 'target.xlsx'
            target_path.write_bytes(target['data'])
            paths, labels = [], {str(target_path): target['name'], target_path.name: target['name']}
            for index, upload in enumerate(sources):
                path = base / f'source_{index:04d}.xlsx'
                path.write_bytes(upload['data'])
                paths.append(path)
                labels[str(path)] = upload['name']
                labels[path.name] = upload['name']
            yield target_path, paths, labels, base, target


def _sanitize(value, labels):
    replacements = dict(labels)
    replacements.update({path.replace('\\', '/'): name for path, name in labels.items()})
    pattern = re.compile('|'.join(re.escape(k) for k in sorted(replacements, key=len, reverse=True)))
    def clean(item):
        if isinstance(item, dict):
            return {k: clean(v) for k, v in item.items() if not str(k).startswith('_')}
        if isinstance(item, list):
            return [clean(x) for x in item]
        if isinstance(item, str):
            return pattern.sub(lambda m: replacements[m[0]], item)
        return item
    return clean(value)


def _safe_prepare(target, sources, supplements, labels):
    try:
        return core.prepare_update(target, sources, reference_prices=supplements)
    except (ET.ParseError, KeyError, zipfile.BadZipFile, ValueError, OSError) as exc:
        raise ValueError(_sanitize('无法分析工作簿：' + str(exc), labels)) from None


def inspect_inputs(target: Upload, sources: list[Upload]):
    with _workspace(target, sources) as (target_path, paths, labels, _, uploaded):
        plan = _safe_prepare(target_path, paths, {}, labels)
        # Missing supplements are a later step; malformed source files are not.
        public = _sanitize(plan, labels)
        if not plan['records']:
            raise ValueError('\n'.join(public['errors']) or '未识别到结算单合计数据。')
        return {'months': plan['months'], 'file_count': len(paths),
                'record_count': len(plan['records']), 'target_name': uploaded['name']}


def preview_update(target: Upload, sources: list[Upload], supplements=None):
    fingerprint = input_fingerprint(target, sources, supplements)
    parsed = _supplements(supplements)
    with _workspace(target, sources) as (target_path, paths, labels, _, _uploaded):
        plan = _safe_prepare(target_path, paths, parsed, labels)
        public = _sanitize(plan, labels)
        public['fingerprint'] = fingerprint
        return public


def _csv_bytes(changes):
    stream = io.StringIO(newline='')
    writer = csv.writer(stream)
    writer.writerow(['月份', '场站', '单元格', '字段', '更新前', '更新后', '来源'])
    for change in changes:
        values = [change.get(k) for k in ('month', 'unit', 'cell', 'field', 'old', 'new', 'source')]
        values = ["'" + v if isinstance(v, str) and v.startswith(('=', '+', '-', '@')) else v for v in values]
        writer.writerow(values)
    return stream.getvalue().encode('utf-8-sig')


def generate_update(target: Upload, sources: list[Upload], supplements=None, expected_fingerprint=''):
    fingerprint = input_fingerprint(target, sources, supplements)
    if not expected_fingerprint or fingerprint != expected_fingerprint:
        raise ValueError('上传文件或补充数据已改变，请重新分析并预览。')
    parsed = _supplements(supplements)
    with _workspace(target, sources) as (target_path, paths, labels, base, uploaded):
        plan = _safe_prepare(target_path, paths, parsed, labels)
        if plan['errors']:
            raise ValueError('请先解决预览中的问题：\n' + '\n'.join(_sanitize(plan['errors'], labels)))
        output = base / 'result.xlsx'
        try:
            applied = core.apply_update(plan, output)
        except (ValueError, OSError, KeyError, ET.ParseError) as exc:
            raise ValueError(_sanitize('未能生成分析表：' + str(exc), labels)) from None
        workbook_bytes = Path(applied['output']).read_bytes()
        name = Path(uploaded['name']).stem[:100]
        workbook_name = name + '_更新后.xlsx'
        backup_name = name + '_更新前备份.xlsx'
        labels.update({str(output): workbook_name, str(base): '本次会话'})
        public = _sanitize(plan, labels)
        warnings = _sanitize(applied.get('warnings', []), labels)
        public.update({'fingerprint': fingerprint, 'output': workbook_name,
                       'backup': backup_name, 'warnings': warnings,
                       'completed': datetime.now(timezone(timedelta(hours=8))).isoformat(timespec='seconds')})
        report_json = json.dumps(public, ensure_ascii=False, indent=2).encode('utf-8')
        report_csv = _csv_bytes(public['changes'])
        bundle = io.BytesIO()
        with zipfile.ZipFile(bundle, 'w', compression=zipfile.ZIP_DEFLATED) as z:
            z.writestr(workbook_name, workbook_bytes)
            z.writestr(backup_name, uploaded['data'])
            z.writestr('更新报告.json', report_json)
            z.writestr('变更明细.csv', report_csv)
        return {'fingerprint': fingerprint, 'workbook_bytes': workbook_bytes,
                'workbook_name': workbook_name, 'report_json': report_json, 'report_csv': report_csv,
                'bundle_bytes': bundle.getvalue(), 'bundle_name': name + '_更新结果.zip',
                'summary': plan['summary'], 'warnings': warnings}
