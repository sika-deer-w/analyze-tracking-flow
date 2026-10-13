# analyze-tracking-flow

按时间顺序核对埋点。监听埋点平台的增量接口，对照事件定义表，输出逐事件明细、零触发告警和按会话排好的事件流。

仓库里没有埋点平台地址。接口 curl 放在本机权限受限的文件里，用 `--curl-file` 传入。不要把 curl、Authorization、Cookie 或 Token 写进仓库和报告。

示例只使用占位地址：

```text
https://<埋点平台域名>/<事件增量路径>
```

## 依赖

- Python 3
- 本机 `curl`
- 事件定义：CSV、JSON、Markdown 表格或 XLSX。表格需要「事件名称」

脚本里的 `http://schemas.openxmlformats.org/...` 是读取 XLSX 的 Office 标准命名空间，不是业务地址。

## 用法

```bash
python3 scripts/tracking_flow.py monitor \
  --curl-file /path/request.curl \
  --spec /path/tracking.xlsx \
  --focus-fields option,state,duration \
  --output /path/report
```

已有事件时：

```bash
python3 scripts/tracking_flow.py render \
  --events /path/raw-events.jsonl \
  --spec /path/tracking.xlsx \
  --output /path/report
```

停止监听时对进程发送 Ctrl+C，脚本会生成报告。

## 产出

| 文件 | 内容 |
|---|---|
| `event-summary.csv` | 定义全集、触发次数；未触发标为「⚠ 未触发」 |
| `event-table.csv` | 逐事件数据表 |
| `event-table.md` | 先列零触发，再列次数和按时间排序的明细 |
| `event-flow.md` | 按 `sessionId` 排列的事件流 |
| `raw-events.jsonl` | 可重新渲染的原始事件 |

## 目录

| 路径 | 内容 |
|---|---|
| `SKILL.md` | 给 Agent 的分析规程 |
| `scripts/tracking_flow.py` | 监听、校验和出报告 |
| `agents/openai.yaml` | 展示名称与默认提示 |

## 安装

```bash
git clone https://github.com/sika-deer-w/analyze-tracking-flow.git \
  ~/.cursor/skills/analyze-tracking-flow
```

Codex 可放到 `~/.codex/skills/analyze-tracking-flow`。
