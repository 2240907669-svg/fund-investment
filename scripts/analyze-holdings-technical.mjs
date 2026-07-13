#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';
import { readCsv, writeCsv, mean, maxDrawdown, number } from '../.agents/skills/analyze-otc-fund-swings/scripts/lib/engine.mjs';

const root = path.resolve(process.argv[2] || '.');
const asOf = process.argv[3] || new Intl.DateTimeFormat('en-CA', {
  timeZone: 'Asia/Shanghai', year: 'numeric', month: '2-digit', day: '2-digit'
}).format(new Date());
const nav = readCsv(path.join(root, 'data', 'nav-history.csv'))
  .filter((row) => row.is_estimate !== 'true' && row.date <= asOf && number(row.nav) > 0);
const funds = readCsv(path.join(root, 'data', 'funds.csv'));
const intake = readCsv(path.join(root, 'data', 'holding-intake.csv'));
const intakeByCode = new Map(intake.map((row) => [row.fund_code, row]));
const benchmarkCode = '007339';

function series(code) {
  return nav.filter((row) => row.fund_code === code)
    .sort((a, b) => a.date.localeCompare(b.date))
    .map((row) => ({ date: row.date, value: number(row.nav), source: row.source, fetchedAt: row.fetched_at }));
}

function dayDiff(later, earlier) {
  return Math.round((new Date(`${later}T00:00:00Z`) - new Date(`${earlier}T00:00:00Z`)) / 86400000);
}

function ret(values, observations) {
  return values.length > observations ? values.at(-1) / values.at(-(observations + 1)) - 1 : null;
}

function rsiWilder(values, period = 14) {
  if (values.length < period + 1) return null;
  const changes = [];
  for (let i = 1; i < values.length; i += 1) changes.push(values[i] / values[i - 1] - 1);
  let gain = mean(changes.slice(0, period).map((value) => Math.max(0, value)));
  let loss = mean(changes.slice(0, period).map((value) => Math.max(0, -value)));
  for (const change of changes.slice(period)) {
    gain = (gain * (period - 1) + Math.max(0, change)) / period;
    loss = (loss * (period - 1) + Math.max(0, -change)) / period;
  }
  if (loss === 0) return gain === 0 ? 50 : 100;
  return 100 - 100 / (1 + gain / loss);
}

function valueOnOrBefore(rows, date) {
  let answer = null;
  for (const row of rows) {
    if (row.date > date) break;
    answer = row.value;
  }
  return answer;
}

function relativeReturn(rows, benchmark, observations) {
  if (rows.length <= observations) return null;
  const start = rows.at(-(observations + 1));
  const end = rows.at(-1);
  const benchStart = valueOnOrBefore(benchmark, start.date);
  const benchEnd = valueOnOrBefore(benchmark, end.date);
  if (!benchStart || !benchEnd) return null;
  return end.value / start.value - benchEnd / benchStart;
}

const benchmark = series(benchmarkCode);
const rows = funds.map((fund) => {
  const data = series(fund.fund_code);
  const values = data.map((row) => row.value);
  const latest = data.at(-1);
  const ma20 = values.length >= 20 ? mean(values.slice(-20)) : null;
  const ma60 = values.length >= 60 ? mean(values.slice(-60)) : null;
  const state = latest && ma20 && ma60
    ? latest.value >= ma20 && ma20 >= ma60 ? '多头排列'
      : latest.value < ma20 && ma20 < ma60 ? '空头排列'
        : latest.value >= ma20 ? '站上20日线，结构未完全转强'
          : '跌破20日线，结构未完全转弱'
    : '历史不足';
  const intakeRow = intakeByCode.get(fund.fund_code) || {};
  return {
    as_of: asOf,
    fund_code: fund.fund_code,
    fund_name: fund.fund_name,
    status: intakeRow.status || '',
    is_qdii: fund.is_qdii,
    observations: values.length,
    latest_date: latest?.date || '',
    latest_nav: latest?.value?.toFixed(4) || '',
    stale_calendar_days: latest ? dayDiff(asOf, latest.date) : '',
    ma20: ma20?.toFixed(4) || '',
    nav_vs_ma20: ma20 ? (latest.value / ma20 - 1).toFixed(8) : '',
    ma60: ma60?.toFixed(4) || '',
    nav_vs_ma60: ma60 ? (latest.value / ma60 - 1).toFixed(8) : '',
    rsi14: rsiWilder(values)?.toFixed(2) || '',
    return_20obs: ret(values, 20)?.toFixed(8) || '',
    return_60obs: ret(values, 60)?.toFixed(8) || '',
    drawdown_60obs: values.length >= 60 ? maxDrawdown(values.slice(-60)).toFixed(8) : '',
    max_drawdown_full_history: values.length ? maxDrawdown(values).toFixed(8) : '',
    relative_20obs_vs_007339: fund.fund_code === benchmarkCode ? '0.00000000' : relativeReturn(data, benchmark, 20)?.toFixed(8) || '',
    relative_60obs_vs_007339: fund.fund_code === benchmarkCode ? '0.00000000' : relativeReturn(data, benchmark, 60)?.toFixed(8) || '',
    technical_state: state,
    source: latest?.source || '',
    fetched_at: latest?.fetchedAt || ''
  };
});

const output = path.join(root, 'reports', `${asOf}-holding-technical-indicators.csv`);
writeCsv(output, rows, Object.keys(rows[0]));
const counts = {
  funds: rows.length,
  latestOnAsOf: rows.filter((row) => row.latest_date === asOf).length,
  qdiiStale: rows.filter((row) => row.is_qdii === 'true' && number(row.stale_calendar_days) > 2).length,
  aboveMa20: rows.filter((row) => number(row.nav_vs_ma20) >= 0).length,
  belowMa20: rows.filter((row) => number(row.nav_vs_ma20) < 0).length,
  output
};
process.stdout.write(`${JSON.stringify(counts, null, 2)}\n`);
