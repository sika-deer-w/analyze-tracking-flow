#!/usr/bin/env python3
"""Capture tracking events from a curl request and render ordered event reports."""

from __future__ import annotations

import argparse
import csv
import dataclasses
import datetime as dt
import json
import os
import re
import shlex
import signal
import subprocess
import sys
import time
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree as ET
from zoneinfo import ZoneInfo


@dataclasses.dataclass
class PropertySpec:
    name: str
    data_type: str = ""
    display_name: str = ""
    enums: dict[str, str] = dataclasses.field(default_factory=dict)
    visitor_required: bool = False
    normal_required: bool = False


@dataclasses.dataclass
class EventSpec:
    name: str
    description: str = ""
    properties: list[PropertySpec] = dataclasses.field(default_factory=list)


@dataclasses.dataclass
class CurlRequest:
    url: str
    method: str
    headers: list[str]
    body: dict[str, Any]
    insecure: bool = False


HEADER_ALIASES = {
    "event": ("事件名称", "eventname", "event_name"),
    "description": ("事件描述", "description", "eventdescription"),
    "property": ("属性名称", "propertyname", "properties"),
    "type": ("数据类型", "datatype", "type"),
    "display": ("属性显示名", "propertydisplayname", "displayname"),
    "enum": ("枚举值", "enum"),
    "visitor_required": ("游客是否必传", "visitorrequired"),
    "normal_required": ("非游客是否必传", "normaluserrequired", "normalrequired"),
}

COMMON_CONTEXT_KEYS = ("stageId", "subjectId", "publisherId", "semesterId")
SESSION_KEYS = ("sessionId", "session", "session_id")
CARD_KEYS = ("cardId", "card_id")
TASK_KEYS = ("taskId", "task_id")


def normalize_header(value: Any) -> str:
    text = str(value or "").strip().lower()
    return re.sub(r"[\s（）()【】\[\]：:＆&/_-]+", "", text.replace("必填", ""))


def split_lines(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, list):
        return [str(x).strip() for x in value]
    return [x.strip() for x in re.split(r"\r?\n", str(value)) if x.strip()]


def truthy_required(value: Any) -> bool:
    return str(value or "").strip().lower() in {"是", "yes", "true", "1", "必传", "required"}


def parse_enums(value: Any) -> dict[str, str]:
    text = "\n".join(split_lines(value)).strip().strip("[]")
    if not text or text.lower() == "none":
        return {}
    result: dict[str, str] = {}
    for match in re.finditer(r"(?:^|[-,;，；\n])\s*([\w.]+)\s*[:：]\s*([^\-\n,;，；\]]+)", text):
        result[match.group(1).strip()] = match.group(2).strip()
    return result


def find_column(headers: list[Any], kind: str) -> int | None:
    normalized = [normalize_header(x) for x in headers]
    for i, cell in enumerate(normalized):
        for alias in HEADER_ALIASES[kind]:
            needle = normalize_header(alias)
            if needle and (cell == needle or needle in cell):
                return i
    return None


def rows_to_specs(rows: list[list[Any]]) -> dict[str, EventSpec]:
    header_index = None
    for i, row in enumerate(rows):
        if find_column(row, "event") is not None:
            header_index = i
            break
    if header_index is None:
        raise ValueError("埋点文件中未找到“事件名称”表头")
    headers = rows[header_index]
    cols = {kind: find_column(headers, kind) for kind in HEADER_ALIASES}
    event_col = cols["event"]
    assert event_col is not None
    result: dict[str, EventSpec] = {}

    def cell(row: list[Any], kind: str) -> Any:
        idx = cols[kind]
        return row[idx] if idx is not None and idx < len(row) else None

    for row in rows[header_index + 1 :]:
        if event_col >= len(row):
            continue
        name = str(row[event_col] or "").strip()
        if not name:
            continue
        prop_names = split_lines(cell(row, "property"))
        types = split_lines(cell(row, "type"))
        displays = split_lines(cell(row, "display"))
        enums = split_lines(cell(row, "enum"))
        visitor = split_lines(cell(row, "visitor_required"))
        normal = split_lines(cell(row, "normal_required"))
        properties = []
        for j, prop_name in enumerate(prop_names):
            properties.append(
                PropertySpec(
                    name=prop_name,
                    data_type=types[j] if j < len(types) else "",
                    display_name=displays[j] if j < len(displays) else "",
                    enums=parse_enums(enums[j] if j < len(enums) else ""),
                    visitor_required=truthy_required(visitor[j] if j < len(visitor) else ""),
                    normal_required=truthy_required(normal[j] if j < len(normal) else ""),
                )
            )
        result[name] = EventSpec(name=name, description=str(cell(row, "description") or "").strip(), properties=properties)
    if not result:
        raise ValueError("埋点文件未解析出任何事件")
    return result


def load_table_rows(path: Path, sheet: str | None) -> list[list[Any]]:
    suffix = path.suffix.lower()
    if suffix in {".csv", ".tsv"}:
        with path.open("r", encoding="utf-8-sig", newline="") as f:
            return list(csv.reader(f, delimiter="\t" if suffix == ".tsv" else ","))
    if suffix in {".xlsx", ".xlsm"}:
        return load_xlsx_rows(path, sheet)
    if suffix in {".md", ".markdown"}:
        rows = []
        for line in path.read_text(encoding="utf-8").splitlines():
            if "|" not in line:
                continue
            cells = [x.strip() for x in line.strip().strip("|").split("|")]
            if cells and all(re.fullmatch(r":?-{3,}:?", x or "") for x in cells):
                continue
            rows.append(cells)
        return rows
    raise ValueError(f"不支持的埋点文件格式：{suffix}")


def load_xlsx_rows(path: Path, sheet: str | None) -> list[list[Any]]:
    """Read plain XLSX cell values using only the Python standard library."""
    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    pkg_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"

    def text_content(element: ET.Element | None) -> str:
        return "" if element is None else "".join(element.itertext())

    with zipfile.ZipFile(path) as archive:
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            shared_root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in shared_root.findall(f"{{{main_ns}}}si"):
                shared_strings.append("".join(node.text or "" for node in item.iter(f"{{{main_ns}}}t")))

        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in relationships.findall(f"{{{pkg_rel_ns}}}Relationship")
        }
        available: list[tuple[str, str]] = []
        for item in workbook.findall(f".//{{{main_ns}}}sheet"):
            rel_id = item.attrib.get(f"{{{rel_ns}}}id", "")
            target = targets.get(rel_id, "")
            if target:
                available.append((item.attrib.get("name", ""), target))
        if not available:
            raise ValueError("XLSX 中没有工作表")
        selected = next((item for item in available if item[0] == sheet), None) if sheet else available[0]
        if selected is None:
            raise ValueError(f"XLSX 中不存在工作表 {sheet}；可选：{', '.join(name for name, _ in available)}")
        target = selected[1].lstrip("/")
        sheet_path = target if target.startswith("xl/") else "xl/" + target
        sheet_root = ET.fromstring(archive.read(sheet_path))
        rows: list[list[Any]] = []
        for row in sheet_root.findall(f".//{{{main_ns}}}sheetData/{{{main_ns}}}row"):
            values: dict[int, Any] = {}
            for cell in row.findall(f"{{{main_ns}}}c"):
                reference = cell.attrib.get("r", "A1")
                letters = re.match(r"[A-Z]+", reference)
                if not letters:
                    continue
                column = 0
                for char in letters.group(0):
                    column = column * 26 + ord(char) - 64
                index = column - 1
                cell_type = cell.attrib.get("t", "")
                raw = text_content(cell.find(f"{{{main_ns}}}v"))
                if cell_type == "s" and raw:
                    value: Any = shared_strings[int(raw)]
                elif cell_type == "inlineStr":
                    value = "".join(node.text or "" for node in cell.iter(f"{{{main_ns}}}t"))
                elif cell_type == "b":
                    value = raw == "1"
                else:
                    value = raw
                values[index] = value
            if values:
                width = max(values) + 1
                rows.append([values.get(i) for i in range(width)])
            else:
                rows.append([])
        return rows


def load_specs(path: Path, sheet: str | None = None) -> dict[str, EventSpec]:
    if path.suffix.lower() != ".json":
        return rows_to_specs(load_table_rows(path, sheet))
    payload = json.loads(path.read_text(encoding="utf-8"))
    items = payload.get("events", payload) if isinstance(payload, dict) else payload
    if not isinstance(items, list):
        raise ValueError("JSON 埋点文件必须是事件数组或 {events:[...]}")
    result: dict[str, EventSpec] = {}
    for item in items:
        if not isinstance(item, dict):
            continue
        name = str(item.get("eventName") or item.get("event_name") or item.get("name") or "").strip()
        if not name:
            continue
        props = []
        for prop in item.get("properties", []):
            if not isinstance(prop, dict) or not prop.get("name"):
                continue
            enums = prop.get("enums", {})
            props.append(PropertySpec(
                name=str(prop["name"]),
                data_type=str(prop.get("type", "")),
                display_name=str(prop.get("displayName") or prop.get("display_name") or ""),
                enums={str(k): str(v) for k, v in enums.items()} if isinstance(enums, dict) else parse_enums(enums),
                visitor_required=bool(prop.get("visitorRequired") or prop.get("visitor_required")),
                normal_required=bool(prop.get("normalRequired") or prop.get("normal_required") or prop.get("required")),
            ))
        result[name] = EventSpec(name=name, description=str(item.get("description", "")), properties=props)
    if not result:
        raise ValueError("JSON 埋点文件未解析出任何事件")
    return result


def parse_curl(command: str) -> CurlRequest:
    tokens = shlex.split(command)
    if tokens and Path(tokens[0]).name == "curl":
        tokens = tokens[1:]
    headers: list[str] = []
    method = ""
    body_text = ""
    url = ""
    insecure = False
    i = 0
    while i < len(tokens):
        token = tokens[i]
        if token in {"-H", "--header"} and i + 1 < len(tokens):
            headers.append(tokens[i + 1]); i += 2; continue
        if token in {"-X", "--request"} and i + 1 < len(tokens):
            method = tokens[i + 1].upper(); i += 2; continue
        if token in {"-d", "--data", "--data-raw", "--data-binary"} and i + 1 < len(tokens):
            body_text = tokens[i + 1]; i += 2; continue
        if token in {"-k", "--insecure"}:
            insecure = True; i += 1; continue
        if token.startswith("http://") or token.startswith("https://"):
            url = token
        i += 1
    if not url:
        raise ValueError("curl 中未找到 http(s) URL")
    if not body_text:
        raise ValueError("curl 中未找到 JSON 请求体（--data-raw/-d）")
    try:
        body = json.loads(body_text)
    except json.JSONDecodeError as exc:
        raise ValueError("curl 请求体必须是 JSON") from exc
    if not isinstance(body, dict):
        raise ValueError("curl JSON 请求体必须是对象")
    return CurlRequest(url=url, method=method or "POST", headers=headers, body=body, insecure=insecure)


def request_events(request: CurlRequest, body: dict[str, Any], timeout: int) -> dict[str, Any]:
    argv = ["curl", "-sS", "--connect-timeout", "10", "--max-time", str(timeout), "-X", request.method, request.url]
    if request.insecure:
        argv.append("--insecure")
    for header in request.headers:
        if not header.lower().startswith("content-length:"):
            argv.extend(["-H", header])
    argv.extend(["--data-raw", json.dumps(body, ensure_ascii=False, separators=(",", ":"))])
    completed = subprocess.run(argv, check=False, capture_output=True, text=True)
    if completed.returncode:
        raise RuntimeError(completed.stderr.strip() or f"curl 退出码 {completed.returncode}")
    try:
        payload = json.loads(completed.stdout)
    except json.JSONDecodeError as exc:
        raise RuntimeError("接口响应不是 JSON") from exc
    if not isinstance(payload, dict):
        raise RuntimeError("接口响应 JSON 顶层必须是对象")
    return payload


def get_first(data: dict[str, Any], keys: Iterable[str]) -> Any:
    for key in keys:
        if key in data and data[key] not in (None, ""):
            return data[key]
    return None


def event_sort_key(event: dict[str, Any]) -> tuple[float, int]:
    text = str(event.get("eventTime") or event.get("serverTime") or "")
    try:
        timestamp = dt.datetime.fromisoformat(text.replace("Z", "+00:00")).timestamp()
    except ValueError:
        timestamp = 0.0
    try:
        event_id = int(event.get("id") or 0)
    except (TypeError, ValueError):
        event_id = 0
    return timestamp, event_id


def display_time(event: dict[str, Any], timezone: str) -> str:
    text = str(event.get("eventTime") or event.get("serverTime") or "")
    if not text:
        return ""
    try:
        value = dt.datetime.fromisoformat(text.replace("Z", "+00:00"))
        return value.astimezone(ZoneInfo(timezone)).isoformat(timespec="milliseconds")
    except (ValueError, KeyError):
        return text


def type_matches(value: Any, expected: str) -> bool:
    kind = expected.strip().lower()
    if not kind:
        return True
    if kind in {"integer", "int"}:
        return isinstance(value, int) and not isinstance(value, bool)
    if kind in {"number", "float", "double"}:
        return isinstance(value, (int, float)) and not isinstance(value, bool)
    if kind in {"string", "str"}:
        return isinstance(value, str)
    if kind in {"boolean", "bool"}:
        return isinstance(value, bool)
    if kind in {"array", "list"}:
        return isinstance(value, list)
    if kind in {"object", "dict", "json"}:
        return isinstance(value, dict)
    return True


def try_json(value: Any) -> Any:
    if not isinstance(value, str):
        return value
    text = value.strip()
    if not text or text[0] not in "[{":
        return value
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return value


def parse_focus_fields(value: str) -> list[str]:
    """Parse comma-separated field selectors while preserving order."""
    result: list[str] = []
    for item in re.split(r"[,，]", value or ""):
        selector = item.strip()
        if selector and selector not in result:
            result.append(selector)
    return result


def nested_value(source: Any, path: str) -> tuple[bool, Any]:
    current = source
    for part in path.split("."):
        if not isinstance(current, dict) or part not in current:
            return False, None
        current = current[part]
    return True, current


def focus_value(event: dict[str, Any], selector: str) -> Any:
    """Return the untouched value selected from event data or the top-level event."""
    data = event.get("data") if isinstance(event.get("data"), dict) else {}
    event_name = str(event.get("eventName") or data.get("eventKey") or "")
    if "." in selector:
        prefix, path = selector.split(".", 1)
        if prefix == event_name:
            return nested_value(data, path)[1]
        if prefix == "data":
            return nested_value(data, path)[1]
        found, value = nested_value(data, selector)
        if found:
            return value
        return nested_value(event, selector)[1]
    if selector in data:
        return data[selector]
    return event.get(selector)


def focus_column(selector: str) -> str:
    return f"{selector}（重点）"


def validate_event(event: dict[str, Any], spec: EventSpec, audience: str) -> list[str]:
    data = event.get("data") if isinstance(event.get("data"), dict) else {}
    problems: list[str] = []
    for prop in spec.properties:
        value = data.get(prop.name)
        required = prop.normal_required if audience == "normal" else prop.visitor_required
        if required and value in (None, ""):
            problems.append(f"缺少必传字段 {prop.name}")
            continue
        if value is not None and not type_matches(value, prop.data_type):
            problems.append(f"{prop.name} 类型错误：期望 {prop.data_type}，实际 {type(value).__name__}")
        if value is not None and prop.enums and str(value) not in prop.enums:
            problems.append(f"{prop.name} 枚举外值：{value}")
    if event.get("status") is False:
        problems.append("平台 status=false")
    platform_failures = event.get("validateFailList") or []
    if platform_failures:
        problems.append("平台校验失败：" + json.dumps(platform_failures, ensure_ascii=False, separators=(",", ":")))
    return problems


def major_data(event: dict[str, Any], spec: EventSpec) -> dict[str, Any]:
    data = event.get("data") if isinstance(event.get("data"), dict) else {}
    if spec.properties:
        return {prop.name: try_json(data.get(prop.name)) for prop in spec.properties if prop.name in data}
    ignored = {"eventKey", "eventTime", "uuid", "packageName", "platform", "os", "d_app_version"}
    return {k: try_json(v) for k, v in data.items() if k not in ignored}


def behavior_text(spec: EventSpec, data: dict[str, Any]) -> str:
    base = re.sub(r"\s+", " ", spec.description).strip() or spec.name
    highlights = []
    for key in ("scene", "actId", "fromPageName", "button", "taskId", "state", "duration", "option"):
        if key in data and data[key] not in (None, ""):
            value = try_json(data[key])
            if key == "duration":
                highlights.append(f"duration={value}s")
            else:
                highlights.append(f"{key}={json.dumps(value, ensure_ascii=False, separators=(',', ':'))}")
    return base + ("｜" + "；".join(highlights) if highlights else "")


def make_records(events: list[dict[str, Any]], specs: dict[str, EventSpec], audience: str, timezone: str, include_unknown: bool, focus_fields: list[str]) -> list[dict[str, Any]]:
    records = []
    for event in sorted(events, key=event_sort_key):
        name = str(event.get("eventName") or (event.get("data") or {}).get("eventKey") or "")
        if name not in specs and not include_unknown:
            continue
        spec = specs.get(name, EventSpec(name=name, description="未在埋点文件中定义"))
        data = event.get("data") if isinstance(event.get("data"), dict) else {}
        problems = validate_event(event, spec, audience) if name in specs else ["未在埋点文件中定义"]
        record = {
            "序号": len(records) + 1,
            "时间": display_time(event, timezone),
            "事件ID": str(event.get("id") or event.get("eventId") or ""),
            "事件名称": name,
            "事件描述": spec.description,
            "会话ID": str(get_first(data, SESSION_KEYS) or ""),
            "卡片ID": str(get_first(data, CARD_KEYS) or ""),
            "任务ID": str(get_first(data, TASK_KEYS) or ""),
            "触发行为": behavior_text(spec, data),
            "主要数据": major_data(event, spec),
            "校验结果": "通过" if not problems else "异常",
            "异常详情": problems,
            "_event": event,
        }
        for selector in focus_fields:
            record[focus_column(selector)] = focus_value(event, selector)
        records.append(record)
    return records


def markdown_escape(value: Any) -> str:
    if isinstance(value, (dict, list)):
        value = json.dumps(value, ensure_ascii=False, separators=(",", ":"))
    return ("" if value is None else str(value)).replace("|", "\\|").replace("\n", "<br>")


def csv_cell(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False, separators=(",", ":"))
    return value


def untriggered_events(records: list[dict[str, Any]], specs: dict[str, EventSpec]) -> list[str]:
    triggered = {record["事件名称"] for record in records}
    return [name for name in specs if name not in triggered]


def write_table(records: list[dict[str, Any]], specs: dict[str, EventSpec], output: Path, focus_fields: list[str]) -> None:
    focus_columns = [focus_column(selector) for selector in focus_fields]
    columns = ["序号", "时间", "事件ID", "事件名称", "事件描述", "会话ID", "卡片ID", "任务ID", *focus_columns, "触发行为", "主要数据", "校验结果", "异常详情"]
    with (output / "event-table.csv").open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=columns)
        writer.writeheader()
        for record in records:
            row = {k: record[k] for k in columns}
            for column in focus_columns:
                row[column] = csv_cell(row[column])
            row["主要数据"] = json.dumps(record["主要数据"], ensure_ascii=False, separators=(",", ":"))
            row["异常详情"] = "；".join(record["异常详情"])
            writer.writerow(row)
    counts = Counter(r["事件名称"] for r in records)
    missing = untriggered_events(records, specs)
    with (output / "event-summary.csv").open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["事件名称", "定义描述", "次数", "触发状态"])
        for name in specs:
            writer.writerow([name, specs[name].description, counts.get(name, 0), "⚠ 未触发" if counts.get(name, 0) == 0 else "已触发"])
    lines = ["# 埋点数据表", "", f"共捕获并纳入报告 **{len(records)}** 条事件；埋点文件定义 **{len(specs)}** 个事件，实际触发 **{len(specs) - len(missing)}** 个。", ""]
    if missing:
        lines += [f"## ⚠️ 重点：未触发埋点（{len(missing)}）", "", "本次监听期间以下埋点一次也没有触发，需要重点确认是否漏报、触发路径未覆盖或埋点定义已失效。", "", "| 事件名称 | 定义描述 | 定义属性 |", "|---|---|---|"]
        for name in missing:
            properties = ", ".join(prop.name for prop in specs[name].properties)
            lines.append(f"| {markdown_escape(name)} | {markdown_escape(specs[name].description)} | {markdown_escape(properties)} |")
        lines.append("")
    else:
        lines += ["## 未触发埋点", "", "埋点文件中的事件均至少触发一次。", ""]
    lines += ["## 触发概览", "", "| 事件名称 | 定义描述 | 次数 | 触发状态 |", "|---|---|---:|---|"]
    for name in specs:
        count = counts.get(name, 0)
        lines.append(f"| {markdown_escape(name)} | {markdown_escape(specs[name].description)} | {count} | {'⚠️ 未触发' if count == 0 else '已触发'} |")
    lines += ["", "## 按时间排序的事件明细", "", "| " + " | ".join(columns) + " |", "|" + "|".join("---:" if c == "序号" else "---" for c in columns) + "|"]
    for record in records:
        values = []
        for col in columns:
            value = record[col]
            if col == "异常详情":
                value = "；".join(value)
            values.append(markdown_escape(value))
        lines.append("| " + " | ".join(values) + " |")
    (output / "event-table.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


def context_compatible(left: dict[str, Any], right: dict[str, Any]) -> bool:
    left_data = left.get("_event", {}).get("data") or {}
    right_data = right.get("_event", {}).get("data") or {}
    compared = 0
    for key in COMMON_CONTEXT_KEYS:
        if key in left_data and key in right_data:
            compared += 1
            if left_data[key] != right_data[key]:
                return False
    return compared > 0


def assign_flows(records: list[dict[str, Any]], lookback_seconds: int) -> dict[str, list[dict[str, Any]]]:
    assignment: list[str | None] = [r["会话ID"] or None for r in records]
    for i, record in enumerate(records):
        if assignment[i]:
            continue
        current_time = event_sort_key(record["_event"])[0]
        for j in range(i + 1, len(records)):
            if not assignment[j]:
                continue
            next_time = event_sort_key(records[j]["_event"])[0]
            if next_time - current_time > lookback_seconds:
                break
            if context_compatible(record, records[j]):
                assignment[i] = assignment[j]
                break
    flows: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for flow_id, record in zip(assignment, records):
        flows[flow_id or "未关联会话"].append(record)
    return dict(flows)


def write_flow(records: list[dict[str, Any]], specs: dict[str, EventSpec], output: Path, lookback_seconds: int, focus_fields: list[str]) -> None:
    flows = assign_flows(records, lookback_seconds)
    lines = ["# 事件流", "", "事件按时间排序，并优先按 `sessionId` 分组；无会话 ID 的前置事件会在时间窗口和公共上下文一致时关联到后续会话。", ""]
    missing = untriggered_events(records, specs)
    if missing:
        lines += [f"> ⚠️ **未触发埋点（{len(missing)}）**：" + "、".join(f"`{name}`" for name in missing), ""]
    focus_columns = [focus_column(selector) for selector in focus_fields]
    for flow_id, items in flows.items():
        lines += [f"## 会话 `{markdown_escape(flow_id)}`", ""]
        contexts: dict[str, Any] = {}
        for item in items:
            data = item["_event"].get("data") or {}
            for key in COMMON_CONTEXT_KEYS + CARD_KEYS:
                if key in data and data[key] not in (None, ""):
                    contexts[key] = data[key]
        if contexts:
            lines += ["共同上下文：`" + markdown_escape(json.dumps(contexts, ensure_ascii=False, separators=(",", ":"))) + "`", ""]
        flow_columns = ["顺序", "时间", "埋点事件", *focus_columns, "触发行为", "主要数据", "校验"]
        lines += ["| " + " | ".join(flow_columns) + " |", "|" + "|".join("---:" if column == "顺序" else "---" for column in flow_columns) + "|"]
        labels = []
        for i, item in enumerate(items, 1):
            values = [str(i), markdown_escape(item["时间"]), markdown_escape(item["事件名称"])]
            values.extend(markdown_escape(item[column]) for column in focus_columns)
            values.extend([markdown_escape(item["触发行为"]), markdown_escape(item["主要数据"]), markdown_escape(item["校验结果"])])
            lines.append("| " + " | ".join(values) + " |")
            labels.append(item["事件名称"])
        lines += ["", "流程串联：", "", "`" + " → ".join(labels) + "`", ""]
    (output / "event-flow.md").write_text("\n".join(lines), encoding="utf-8")


def load_events(path: Path, list_field: str) -> list[dict[str, Any]]:
    text = path.read_text(encoding="utf-8").strip()
    if not text:
        return []
    try:
        payload = json.loads(text)
        if isinstance(payload, list):
            return [x for x in payload if isinstance(x, dict)]
        if isinstance(payload, dict):
            items = payload.get(list_field, payload.get("events", []))
            return [x for x in items if isinstance(x, dict)] if isinstance(items, list) else []
    except json.JSONDecodeError:
        pass
    events = []
    for line_number, line in enumerate(text.splitlines(), 1):
        if not line.strip():
            continue
        try:
            item = json.loads(line)
        except json.JSONDecodeError as exc:
            raise ValueError(f"JSONL 第 {line_number} 行不是合法 JSON") from exc
        if isinstance(item, dict):
            events.append(item)
    return events


def render(events: list[dict[str, Any]], specs: dict[str, EventSpec], args: argparse.Namespace) -> list[dict[str, Any]]:
    output = Path(args.output).expanduser().resolve()
    output.mkdir(parents=True, exist_ok=True)
    focus_fields = parse_focus_fields(args.focus_fields)
    records = make_records(events, specs, args.audience, args.timezone, args.include_unknown, focus_fields)
    write_table(records, specs, output, focus_fields)
    write_flow(records, specs, output, args.session_lookback_seconds, focus_fields)
    return records


def monitor(args: argparse.Namespace) -> int:
    command = Path(args.curl_file).read_text(encoding="utf-8") if args.curl_file else args.curl_command
    request = parse_curl(command)
    specs = load_specs(Path(args.spec), args.sheet)
    output = Path(args.output).expanduser().resolve()
    output.mkdir(parents=True, exist_ok=True)
    raw_path = output / "raw-events.jsonl"
    body = dict(request.body)
    cursor = body.get(args.cursor_field, "")
    seen: set[str] = set()
    captured: list[dict[str, Any]] = []
    stop = False

    def stop_handler(_signum: int, _frame: Any) -> None:
        nonlocal stop
        stop = True

    signal.signal(signal.SIGINT, stop_handler)
    signal.signal(signal.SIGTERM, stop_handler)
    started = time.monotonic()
    print(json.dumps({"status": "listening", "definedEvents": len(specs), "cursor": cursor, "output": str(output)}, ensure_ascii=False), flush=True)
    with raw_path.open("w", encoding="utf-8") as raw:
        while not stop:
            if args.duration_seconds and time.monotonic() - started >= args.duration_seconds:
                break
            try:
                body[args.cursor_field] = cursor
                response = request_events(request, body, args.request_timeout)
                items = response.get(args.list_field) or []
                if not isinstance(items, list):
                    raise RuntimeError(f"响应字段 {args.list_field} 不是数组")
                for event in sorted((x for x in items if isinstance(x, dict)), key=event_sort_key):
                    identity = str(event.get("id") or event.get("eventId") or json.dumps(event, ensure_ascii=False, sort_keys=True))
                    if identity in seen:
                        continue
                    seen.add(identity)
                    name = str(event.get("eventName") or (event.get("data") or {}).get("eventKey") or "")
                    if name not in specs and not args.include_unknown:
                        continue
                    captured.append(event)
                    raw.write(json.dumps(event, ensure_ascii=False, separators=(",", ":")) + "\n")
                    raw.flush()
                    spec = specs.get(name, EventSpec(name=name))
                    problems = validate_event(event, spec, args.audience) if name in specs else ["未在埋点文件中定义"]
                    print(json.dumps({"id": identity, "time": display_time(event, args.timezone), "eventName": name, "behavior": behavior_text(spec, event.get("data") or {}), "validation": "通过" if not problems else problems}, ensure_ascii=False, separators=(",", ":")), flush=True)
                next_cursor = response.get(args.cursor_field)
                if next_cursor not in (None, ""):
                    cursor = next_cursor
                end = time.monotonic() + args.poll_interval
                while not stop and time.monotonic() < end:
                    time.sleep(min(0.2, end - time.monotonic()))
            except Exception as exc:
                print(json.dumps({"status": "poll_error", "cursor": cursor, "error": str(exc)}, ensure_ascii=False), file=sys.stderr, flush=True)
                end = time.monotonic() + args.retry_interval
                while not stop and time.monotonic() < end:
                    time.sleep(min(0.2, end - time.monotonic()))
    records = render(captured, specs, args)
    print(json.dumps({"status": "stopped", "cursor": cursor, "captured": len(records), "untriggeredEvents": untriggered_events(records, specs), "output": str(output)}, ensure_ascii=False), flush=True)
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="监听埋点接口并生成有序埋点表和事件流")
    sub = parser.add_subparsers(dest="command", required=True)

    def common(target: argparse.ArgumentParser) -> None:
        target.add_argument("--spec", required=True, help="埋点定义文件：CSV/JSON/Markdown/XLSX")
        target.add_argument("--sheet", help="XLSX 工作表名称")
        target.add_argument("--output", required=True, help="报告输出目录")
        target.add_argument("--audience", choices=("normal", "visitor"), default="normal", help="按哪类用户校验必传字段")
        target.add_argument("--timezone", default="Asia/Shanghai")
        target.add_argument("--include-unknown", action="store_true", help="同时收录埋点文件未定义的事件")
        target.add_argument("--focus-fields", default="", help="重点字段，逗号分隔；各字段单独成列并保留原始值，如 option,state,duration")
        target.add_argument("--session-lookback-seconds", type=int, default=60, help="将无 sessionId 的前置事件关联到会话的最大时间窗口")
        target.add_argument("--list-field", default="list", help="接口响应中的事件数组字段")

    mon = sub.add_parser("monitor", help="持续轮询，Ctrl+C 后生成报告")
    source = mon.add_mutually_exclusive_group(required=True)
    source.add_argument("--curl-file", help="保存 curl 命令的文本文件")
    source.add_argument("--curl-command", help="curl 命令字符串；不推荐用于含密钥的请求")
    common(mon)
    mon.add_argument("--cursor-field", default="currentMaxId", help="请求体和响应中的增量游标字段")
    mon.add_argument("--poll-interval", type=float, default=2.0)
    mon.add_argument("--retry-interval", type=float, default=3.0)
    mon.add_argument("--request-timeout", type=int, default=20)
    mon.add_argument("--duration-seconds", type=float, default=0, help="0 表示持续到 Ctrl+C")

    existing = sub.add_parser("render", help="从已有事件 JSON/JSONL 生成报告")
    existing.add_argument("--events", required=True)
    common(existing)
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    try:
        if args.command == "monitor":
            return monitor(args)
        specs = load_specs(Path(args.spec), args.sheet)
        events = load_events(Path(args.events), args.list_field)
        records = render(events, specs, args)
        print(json.dumps({"status": "rendered", "captured": len(records), "untriggeredEvents": untriggered_events(records, specs), "output": str(Path(args.output).resolve())}, ensure_ascii=False))
        return 0
    except Exception as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
