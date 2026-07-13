#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';
import { loadProject, planActions, renderReport, scoreAll, signalRows, toCsv, writeCsv, number, dayDiff, formatPct } from './lib/engine.mjs';

function args(argv) {
  const result = { root: '.', mode: 'action', asOf: '' };
  for (let i = 2; i < argv.length; i += 1) {
    if (argv[i] === '--root') result.root = argv[++i];
    else if (argv[i] === '--mode') result.mode = argv[++i];
    else if (argv[i] === '--as-of') result.asOf = argv[++i];
  }
  return result;
}

function shanghaiDate() {
  const parts = new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Shanghai', year: 'numeric', month: '2-digit', day: '2-digit' }).formatToParts(new Date());
  const map = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${map.year}-${map.month}-${map.day}`;
}

function weeklyAppend(project, asOf) {
  const confirmed = project.transactions.filter((row) => row.status === 'confirmed' && String(row.confirmed_at).slice(0, 10) <= asOf);
  const start = new Date(`${asOf}T00:00:00Z`); start.setUTCDate(start.getUTCDate() - 6);
  const from = start.toISOString().slice(0, 10);
  const week = confirmed.filter((row) => String(row.confirmed_at).slice(0, 10) >= from);
  const buys = week.filter((row) => row.side === 'buy');
  const sells = week.filter((row) => row.side === 'sell');
  const fees = week.reduce((sum, row) => sum + number(row.fees), 0);
  const turnover = week.reduce((sum, row) => sum + number(row.amount), 0);
  const portfolio = project.portfolio.filter((row) => row.date >= from && row.date <= asOf).sort((a, b) => a.date.localeCompare(b.date));
  const values = portfolio.map((row) => number(row.total_value)).filter((value) => value > 0);
  const weekReturn = values.length > 1 ? values.at(-1) / values[0] - 1 : null;
  let peak = 0, drawdown = 0;
  values.forEach((value) => { peak = Math.max(peak, value); if (peak) drawdown = Math.max(drawdown, 1 - value / peak); });
  const held = sells.map((sell) => {
    const buy = [...confirmed].reverse().find((row) => row.fund_code === sell.fund_code && row.side === 'buy' && row.confirmed_at < sell.confirmed_at);
    return buy ? dayDiff(sell.confirmed_at, buy.confirmed_at) : null;
  }).filter((value) => value != null);
  const averageHold = held.length ? held.reduce((sum, value) => sum + value, 0) / held.length : null;
  const dueDecisions = (project.decisions ?? []).filter((row) => row.evaluation_status === 'resolved' && String(row.evaluation_due_date).slice(0, 10) <= asOf);
  const recentDecisions = dueDecisions.slice(-60);
  const decided = recentDecisions.length;
  const correct = recentDecisions.filter((row) => String(row.outcome_correct).toLowerCase() === 'true').length;
  const valueAdds = recentDecisions.map((row) => Number(row.advice_value_add_vs_hold)).filter(Number.isFinite);
  const meanValueAdd = valueAdds.length ? valueAdds.reduce((sum, value) => sum + value, 0) / valueAdds.length : null;
  const brierRows = recentDecisions.filter((row) => ['up', 'flat', 'down'].includes(row.realized_outcome) && [row.up_probability, row.flat_probability, row.down_probability].every((value) => Number.isFinite(Number(value))));
  const brier = brierRows.length ? brierRows.reduce((sum, row) => {
    const probs = [Number(row.up_probability), Number(row.flat_probability), Number(row.down_probability)];
    if (Math.max(...probs) > 1) probs.forEach((value, index) => { probs[index] = value / 100; });
    const actual = ['up', 'flat', 'down'].map((label) => label === row.realized_outcome ? 1 : 0);
    return sum + probs.reduce((score, probability, index) => score + (probability - actual[index]) ** 2, 0) / 3;
  }, 0) / brierRows.length : null;
  return ['', '## 本周执行与纪律', '', `- 区间：${from} 至 ${asOf}`, `- 费用后组合变化：${weekReturn == null ? '数据不足' : formatPct(weekReturn)}`,
    `- 周内最大回撤：${values.length ? formatPct(drawdown) : '数据不足'}`, `- 确认申购/赎回：${buys.length}/${sells.length}笔`,
    `- 周换手金额：${turnover.toFixed(2)}元`, `- 已记录费用：${fees.toFixed(2)}元`,
    `- 已完成持仓平均天数：${averageHold == null ? '数据不足' : averageHold.toFixed(1)}`,
    '', '## 建议滚动自我审计', '', `- 已到期且完成复核样本：${decided}`,
    `- 方向命中率：${decided ? formatPct(correct / decided) : '数据不足；不得猜测'}`,
    `- 相对完全不操作的平均费用后增益：${meanValueAdd == null ? '数据不足' : formatPct(meanValueAdd)}`,
    `- 概率Brier分数（越低越好）：${brier == null ? '数据不足' : brier.toFixed(4)}`,
    `- 审计原则：原始预测不可覆盖；实际执行、模型执行、不操作和宽基基准分别核算。`, ''].join('\n');
}

const options = args(process.argv);
if (!['morning', 'action', 'weekly'].includes(options.mode)) throw new Error('--mode must be morning, action, or weekly');
const project = loadProject(options.root);
const asOf = options.asOf || shanghaiDate();
const scores = scoreAll(project, asOf);
const plan = planActions(project, scores, asOf);
const signalFile = path.join(project.root, 'data', 'signals.csv');
writeCsv(signalFile, signalRows(scores, asOf), ['as_of','fund_code','eligible','score','return_5d','return_10d','return_20d','trend_20d','volatility_20d','drawdown_20d','roundtrip_cost','stale_days','vetoes']);
let report = renderReport(project, scores, plan, asOf, options.mode);
if (options.mode === 'weekly') report += weeklyAppend(project, asOf);
const reportDir = path.join(project.root, 'reports');
fs.mkdirSync(reportDir, { recursive: true });
const reportFile = path.join(reportDir, `${asOf}-${options.mode}.md`);
fs.writeFileSync(reportFile, report, 'utf8');
process.stdout.write(`${reportFile}\n`);
