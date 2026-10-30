"""月度结算助手的 Streamlit 界面。上传内容只属于当前会话。"""
from __future__ import annotations

import hashlib
import json
import math

import streamlit as st


INPUT_FIELDS = (
    ("风电", "风电机制结算参考价", "元/兆瓦时"),
    ("光伏", "光伏机制结算参考价", "元/兆瓦时"),
    ("分布式电量", "分布式光伏电量", "兆瓦时"),
    ("分布式电费", "分布式光伏电费", "元"),
)


def parse_supplements(raw_values):
    """空白与数字零分别保留含义，返回后台需要的月度输入。"""
    supplements, errors = {}, []
    for month, fields in raw_values.items():
        values = {}
        for key, _label, _unit in INPUT_FIELDS:
            raw = str(fields.get(key, "")).strip().replace(",", "").replace("，", "")
            if not raw:
                continue
            try:
                value = float(raw)
                if not math.isfinite(value):
                    raise ValueError()
            except (ValueError, OverflowError):
                errors.append(f"{month}月的{key}请填写有效数字。")
                continue
            if key == "分布式电量" and value < 0:
                errors.append(f"{month}月的分布式电量不能小于0。")
                continue
            values[key] = value
        if ("分布式电量" in values) != ("分布式电费" in values):
            errors.append(f"{month}月的分布式电量和电费需同时填写。")
        if values:
            supplements[int(month)] = values
    return supplements, list(dict.fromkeys(errors))


def to_upload(uploaded_file):
    return {"name": uploaded_file.name, "data": uploaded_file.getvalue()}


def upload_signature(target, sources):
    """以内容识别上传变化，文件名相同也不会复用旧结果。"""
    if target is None:
        return ""
    metadata = {
        "target": [target["name"], hashlib.sha256(target["data"]).hexdigest()],
        "sources": sorted([source["name"], hashlib.sha256(source["data"]).hexdigest()] for source in sources),
    }
    return hashlib.sha256(json.dumps(metadata, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()


def input_signature(upload_key, raw_values):
    content = json.dumps({"uploads": upload_key, "supplements": raw_values}, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(content.encode("utf-8")).hexdigest()


def invalidate_results():
    if st.session_state.get("preview") is not None or st.session_state.get("result") is not None:
        st.session_state["inputs_changed"] = True
    for key in ("preview", "result", "preview_input_signature", "result_input_signature"):
        st.session_state.pop(key, None)


def uploads_changed():
    invalidate_results()
    for key in ("inspection", "inspection_signature", "inspection_error"):
        st.session_state.pop(key, None)


def clear_session():
    epoch = st.session_state.get("upload_epoch", 0) + 1
    for key in list(st.session_state):
        del st.session_state[key]
    st.session_state["upload_epoch"] = epoch


def formatted_value(value):
    if value is None:
        return "（空）"
    if isinstance(value, float):
        return f"{value:,.8f}".rstrip("0").rstrip(".")
    return str(value)


def render_supplements(months, epoch):
    raw_values = {}
    for month in months:
        with st.expander(f"{month}月补充数据", expanded=len(months) == 1):
            columns = st.columns(4)
            raw_values[month] = {}
            for column, (field, label, unit) in zip(columns, INPUT_FIELDS):
                with column:
                    raw_values[month][field] = st.text_input(
                        label, placeholder="留空保留原值", help=f"单位：{unit}。确认为0时填写0。",
                        key=f"supplement_{epoch}_{month}_{field}", on_change=invalidate_results,
                    )
    return raw_values


def render_plan(plan):
    summary = plan.get("summary", {})
    metrics = st.columns(4)
    for column, label, value in zip(
        metrics, ("结算文件", "场站记录", "待更新单元格", "新增数据行"),
        (summary.get("files", "—"), summary.get("records", "—"),
         summary.get("changed_cells", len(plan.get("changes", []))), summary.get("new_rows", 0)),
    ):
        column.metric(label, value)

    errors = plan.get("errors", [])
    warnings = plan.get("warnings", [])
    if errors:
        st.error("需要处理后才能生成文件：\n\n" + "\n\n".join(f"- {item}" for item in errors))
    if warnings:
        with st.expander(f"检查提示（{len(warnings)}项）", expanded=True):
            for warning in warnings:
                st.write(f"• {warning}")

    changes = plan.get("changes", [])
    if changes:
        rows = [
            {
                "月份": f"{change.get('month', '')}月", "场站 / 单元": change.get("unit", ""),
                "字段": change.get("field", ""), "单元格": change.get("cell", ""),
                "原值": formatted_value(change.get("old")), "新值": formatted_value(change.get("new")),
                "来源": change.get("source", ""),
            }
            for change in changes
        ]
        st.dataframe(rows, hide_index=True, width="stretch", height=360)
    elif not errors:
        st.info("所选结算数据与目标表一致，无需改动。仍可生成文件并下载检查报告。")


def render_downloads(result):
    count = result.get("summary", {}).get("changed_cells", 0)
    if count:
        st.success(f"已生成更新文件，共更新 {count:,} 个单元格。下载后请保留，作为下个月的目标表。")
    else:
        st.success("已生成下载文件。数据与目标表一致，本次没有新增变更。")
    columns = st.columns(2)
    with columns[0]:
        st.download_button(
            "下载更新后的 Excel", data=result["workbook_bytes"], file_name=result["workbook_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", on_click="ignore", width="stretch",
        )
    with columns[1]:
        st.download_button(
            "下载 Excel 和报告压缩包", data=result["bundle_bytes"], file_name=result["bundle_name"],
            mime="application/zip", on_click="ignore", width="stretch",
        )
    with st.expander("单独下载检查报告"):
        report_columns = st.columns(2)
        with report_columns[0]:
            st.download_button("下载 CSV 变更明细", result["report_csv"], file_name="结算更新明细.csv", mime="text/csv", on_click="ignore")
        with report_columns[1]:
            st.download_button("下载 JSON 完整报告", result["report_json"], file_name="结算更新报告.json", mime="application/json", on_click="ignore")
    if result.get("warnings"):
        with st.expander("生成结果提示"):
            for warning in result["warnings"]:
                st.write(f"• {warning}")
    st.caption("压缩包包含更新前原表备份、新版Excel和检查报告。首次打开新版Excel后，请在“数据”菜单执行“全部刷新”，核对透视表。")


def main():
    st.set_page_config(page_title="月度结算助手", page_icon="📊", layout="wide", initial_sidebar_state="collapsed")
    if "upload_epoch" not in st.session_state:
        st.session_state["upload_epoch"] = 0
    epoch = st.session_state["upload_epoch"]

    with st.sidebar:
        st.subheader("月度结算助手")
        st.caption("2026年月度电价分析")
        st.markdown("1. 上传最新目标表和结算单\n2. 补充当月数据并预览\n3. 生成文件并下载保留")
        st.divider()
        st.caption("每月都使用上次下载的最新版目标表。上传内容仅在本次会话中处理，完成后请下载保存。")
        st.button("清空本次会话", on_click=clear_session, width="stretch")

    st.title("月度结算数据汇总")
    st.write("导入结算单，核对差异后更新现有电价分析表。")
    st.caption("浏览器会提供新版文件供你下载，电脑中的原文件不会被直接覆盖。")

    target_column, sources_column = st.columns(2)
    with target_column:
        target_file = st.file_uploader(
            "1. 上传最新目标分析表", type=["xlsx"], key=f"target_upload_{epoch}",
            help="使用2026年月度电价分析表；从第二次更新开始，上传上次下载的新版。", on_change=uploads_changed,
        )
    with sources_column:
        source_files = st.file_uploader(
            "2. 上传本月结算单", type=["xlsx", "zip"], accept_multiple_files=True,
            key=f"sources_upload_{epoch}", help="可多选Excel文件，也可将当月文件夹压缩成ZIP后上传。", on_change=uploads_changed,
        )

    target = to_upload(target_file) if target_file is not None else None
    sources = [to_upload(file) for file in source_files]
    uploads_key = upload_signature(target, sources)
    ready = target is not None and bool(sources)
    if not ready:
        invalidate_results()
        st.info("请先上传目标分析表和本次结算单。")

    if ready and st.session_state.get("inspection_signature") != uploads_key:
        invalidate_results()
        st.session_state.pop("inspection", None)
        st.session_state.pop("inspection_error", None)
        try:
            from service import inspect_inputs
            with st.spinner("正在识别结算月份和场站……"):
                st.session_state["inspection"] = inspect_inputs(target, sources)
        except Exception as exc:
            st.session_state["inspection_error"] = str(exc)
        st.session_state["inspection_signature"] = uploads_key

    inspection = st.session_state.get("inspection", {}) if ready else {}
    inspection_error = st.session_state.get("inspection_error") if ready else None
    if inspection_error:
        st.error("文件未能识别：" + inspection_error)
    months = sorted({int(month) for month in inspection.get("months", [])})
    if inspection and months:
        month_names = "、".join(f"{month}月" for month in months)
        st.caption(f"已识别 {month_names} · {inspection.get('file_count', len(sources))}份结算单 · {inspection.get('record_count', '—')}条场站记录")

    st.subheader("3. 补充每月数据")
    st.caption("新增月份需填写四项，确认为0时填0；历史月份留空保留原值。参考价为元/兆瓦时，电量为兆瓦时，电费为元。")
    raw_values = render_supplements(months, epoch)
    supplements, input_errors = parse_supplements(raw_values)
    for error in input_errors:
        st.error(error)
    if not months:
        st.caption("上传文件后，将按结算单月份显示输入框。")

    current_signature = input_signature(uploads_key, raw_values)
    if st.session_state.get("current_input_signature") != current_signature:
        invalidate_results()
        st.session_state["current_input_signature"] = current_signature
    if st.session_state.get("inputs_changed"):
        st.caption("输入已变化，请重新分析并预览。")

    can_preview = ready and bool(months) and not inspection_error and not input_errors
    if st.button("分析并预览", type="primary", disabled=not can_preview, width="content"):
        invalidate_results()
        try:
            from service import preview_update
            with st.spinner("正在核对结算单与目标分析表……"):
                plan = preview_update(target, sources, supplements)
            st.session_state["preview"] = plan
            st.session_state["preview_input_signature"] = current_signature
            st.session_state["inputs_changed"] = False
        except Exception as exc:
            st.error("预览未完成：" + str(exc))

    preview = st.session_state.get("preview")
    valid_preview = preview is not None and st.session_state.get("preview_input_signature") == current_signature
    if valid_preview:
        st.subheader("本次更新预览")
        render_plan(preview)

    can_generate = valid_preview and not preview.get("errors") and not input_errors
    if st.button("生成更新后的Excel", disabled=not can_generate, type="primary"):
        st.session_state.pop("result", None)
        try:
            from service import generate_update
            with st.spinner("正在生成Excel和检查报告……"):
                result = generate_update(target, sources, supplements, expected_fingerprint=preview["fingerprint"])
            st.session_state["result"] = result
            st.session_state["result_input_signature"] = current_signature
        except Exception as exc:
            st.error("文件未生成：" + str(exc) + " 请重新预览后再试。")

    result = st.session_state.get("result")
    if result is not None and st.session_state.get("result_input_signature") == current_signature:
        st.divider()
        render_downloads(result)


if __name__ == "__main__":
    main()
