# 决策日志使用规则

`decision-journal.csv` 保存报告发布时的原始建议与随后复核结果，用来检验建议是否真正创造费用后价值。

- 发布午报或晚报时立即追加一行，填写到 `evaluation_due_date`；原始字段不得事后覆盖。
- 首次运行从历史报告回填时将 `record_origin` 写为 `reconstructed_from_report`，原生实时记录写为 `native_realtime`；回填缺失项留空并在 `evidence_sources` 保存原报告路径。
- 用户实际操作未知时，`user_executed` 留空，不能假定执行。
- 到达评估日后，以正式基金净值和真实费用填写结果；同时计算“按建议模型执行”“完全不操作”和宽基基准。
- `evaluation_status` 使用 `pending`、`resolved` 或 `insufficient_data`。
- 判断错误时必须填写 `error_type` 和可执行的 `lesson`；规则变化必须满足 `config/review-loop.json` 的样本与测试门槛。
