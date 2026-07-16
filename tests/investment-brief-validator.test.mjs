import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { validateInvestmentBrief } from '../.agents/skills/analyze-otc-fund-swings/scripts/validate-investment-brief.mjs';

const rules = JSON.parse(fs.readFileSync(new URL('../config/agent-improvement-rules.json', import.meta.url), 'utf8'));

const completeBrief = `
# 午报
数据截止：2026-07-15 11:30；来源：https://example.invalid
正式净值与盘中估算已分列。
## 全市场机会扫描
创新药领涨，半导体设备领跌。
## 资金流与主力资金动向
流入前五与流出前五均已核验；判断为存量轮动。
## 可以买什么
今天不买。
## 独立多头论证
反弹扩散。
## 独立空头论证
成交不足。
## 费用后比较
费用和费率计入后，现金占优。
## 情景
基准 50%，偏强 25%，偏弱 25%。
## 只看这一页：行动结论
申请截止时间14:50；推翻条件为电子与通信同时转强。
`;

test('complete market-wide brief passes the publication gate', () => {
  assert.deepEqual(validateInvestmentBrief(completeBrief, rules), {
    ok: true,
    missingSections: [],
    missingPatterns: []
  });
});

test('holdings-only brief fails on missing market, debate, and buy research', () => {
  const result = validateInvestmentBrief('# 持仓复述\n正式净值为1.0。', rules);
  assert.equal(result.ok, false);
  assert.ok(result.missingSections.includes('全市场机会扫描'));
  assert.ok(result.missingSections.includes('可以买什么'));
  assert.ok(result.missingSections.includes('独立空头论证'));
});

test('user-facing action sections reject six-digit fund codes', () => {
  const result = validateInvestmentBrief(
    completeBrief.replace('申请截止时间14:50', '申请截止时间14:50；赎回019024 15%'),
    rules
  );
  assert.equal(result.ok, false);
  assert.ok(result.missingPatterns.some(({ id }) => id === 'fund_full_name_only'));
});
