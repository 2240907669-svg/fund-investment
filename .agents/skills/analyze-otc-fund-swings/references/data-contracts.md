# 数据契约

所有 CSV 使用 UTF-8、首行为字段名。日期使用 `YYYY-MM-DD`，时间使用带时区的 ISO 8601。布尔值使用 `true` 或 `false`，费率使用小数（`0.015` 表示 1.5%）。

## 输入文件

- `config/profile.json`：资金和策略阈值。百分比均使用小数。
- `config/agent-improvement-rules.json`：全市场优先、独立多空、行动比较、学习门槛和报告发布校验规则。
- `data/funds.csv`：基金代码、份额、类别、主题、QDII 标记、规模、成立天数、申赎状态、限额、锁定期、跟踪误差和费率来源。
- `data/fees.csv`：按持有天数区间记录申购费、赎回费、年销售服务费、来源和验证日；费率使用小数。
- `data/nav-history.csv`：日期、代码、净值、来源、抓取时间和估算标记。评分只使用非估算记录。
- `data/holdings.csv`、`data/transactions.csv`、`data/portfolio-history.csv`：持仓、确认交易和组合历史。

## 状态值

- `purchase_status`: `open`, `limited`, `suspended`, `unknown`。
- `redemption_status`: `open`, `suspended`, `unknown`。
- `side`: `buy`, `sell`；`status`: `submitted`, `confirmed`, `cancelled`, `failed`。

## 输出

- `data/signals.csv` 保存可复算评分。
- `reports/YYYY-MM-DD-{morning|action|weekly|backtest}.md` 保存报告。
- 每日研究报告发布前必须通过 `scripts/validate-investment-brief.mjs`；校验失败表示研究环节缺失，不能只添加空标题。

不得改写历史原始记录。同步时按 `date + fund_code + is_estimate` 去重，并保留最新抓取来源。
