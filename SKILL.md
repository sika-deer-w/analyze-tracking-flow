---
name: analyze-tracking-flow
description: Scripted tracking-event capture and flow reconstruction from a tracking-platform curl command plus an event specification file, with configurable raw focus-field columns and explicit zero-trigger warnings. Use when the user asks to monitor/listen to 埋点 events, validate emitted event properties against an 埋点文档, produce an ordered 埋点数据表 with timestamps and major payload data, highlight selected fields, identify defined events that never fired, or reconstruct event/session flows. Supports CSV, JSON, Markdown-table, and XLSX specifications.
---

# 埋点事件流分析

使用 `scripts/tracking_flow.py`。始终保留事件顺序和规范中定义的主要属性，不要只输出聚合次数。

## 输入

必须取得两个输入：

1. 埋点平台 curl：优先保存为权限受限的文本文件并使用 `--curl-file`，避免把 Bearer Token 再次显示在命令或回复中。
2. 埋点定义文件：CSV、JSON、Markdown 表格或 XLSX。表格需包含“事件名称”；可选包含“事件描述、属性名称、数据类型、属性显示名、枚举值、游客是否必传、非游客是否必传”。多属性单元格按换行对齐。

不得用 shell 执行整段 curl。脚本会解析 URL、headers、method 和 JSON body，再通过参数数组调用 curl。

## 持续监听

```bash
python3 scripts/tracking_flow.py monitor \
  --curl-file /path/request.curl \
  --spec /path/tracking.xlsx \
  --focus-fields option,state,duration \
  --output /path/report
```

- 默认仅保留埋点文件中定义的事件，每 2 秒增量轮询。
- 用户说“停止、可以、完毕”时向进程发送 Ctrl+C；脚本停止后自动生成报告。
- 如果接口游标字段不是 `currentMaxId`，使用 `--cursor-field`；事件列表字段不是 `list` 时使用 `--list-field`。
- XLSX 有多张工作表时用 `--sheet` 指定名称。
- 需要定时结束时使用 `--duration-seconds`。
- 需要重点查看属性时使用 `--focus-fields`，逗号分隔，例如 `option,state,duration`。每个字段单独成列，直接保留事件中的原始值，不解析 JSON 字符串。使用 `事件名.字段名` 可只关注某个事件的字段，例如 `clickInteractionModuleElement.state`。
- 不要把 curl、Authorization、Cookie 或 Token 写入报告。

## 对已有事件生成报告

```bash
python3 scripts/tracking_flow.py render \
  --events /path/raw-events.jsonl \
  --spec /path/tracking.xlsx \
  --focus-fields option,state,duration \
  --output /path/report
```

`--events` 接受事件数组、平台响应 `{ "list": [...] }` 或 JSONL。

## 固定产出

- `event-summary.csv`：埋点定义全集、触发次数和触发状态；零触发事件标记为“⚠ 未触发”。
- `event-table.csv`：可继续加工的逐事件数据表。
- `event-table.md`：顶部突出零触发埋点，再展示次数概览和按时间排序的逐事件明细。
- `event-flow.md`：按 `sessionId` 组织的事件流，包含顺序、时间、重点字段、行为和主要数据。
- `raw-events.jsonl`：监听过程的可恢复原始事件，只用于审计和重新渲染。

报告必须先展示逐事件数据表，再解释事件流。把每个零触发埋点作为显著告警列出，不能只在次数表中显示 0。重点检查：缺失必传字段、类型/枚举错误、平台校验失败、重复/乱序、会话或卡片标识变化、环节时长与总时长差异、定义存在但未触发的事件、流程中断位置。
