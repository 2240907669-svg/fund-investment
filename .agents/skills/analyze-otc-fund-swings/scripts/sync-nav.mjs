#!/usr/bin/env node
import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { loadProject, readCsv, writeCsv } from './lib/engine.mjs';

let root = '.', strict = false;
for (let i = 2; i < process.argv.length; i += 1) {
  if (process.argv[i] === '--root') root = process.argv[++i];
  else if (process.argv[i] === '--strict') strict = true;
}
const absolute = path.resolve(root);
const project = loadProject(absolute);
const target = path.join(absolute, 'data', 'nav-history.csv');
const logFile = path.join(absolute, 'data', 'sync-log.jsonl');
const fetchedAt = new Date().toISOString();
const additions = [];
const logs = [];

function shanghaiDate(timestamp) {
  const parts = new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Shanghai', year: 'numeric', month: '2-digit', day: '2-digit' })
    .formatToParts(new Date(Number(timestamp)));
  const map = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${map.year}-${map.month}-${map.day}`;
}

for (const fund of project.funds) {
  const url = `https://fund.eastmoney.com/pingzhongdata/${encodeURIComponent(fund.fund_code)}.js?v=${Date.now()}`;
  try {
    const response = await fetch(url, { headers: { 'user-agent': 'OTC-Fund-Research/1.0 (manual-trading-only)', accept: 'text/javascript,*/*' }, signal: AbortSignal.timeout(15000) });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const text = await response.text();
    const match = text.match(/var\s+Data_netWorthTrend\s*=\s*(\[[\s\S]*?\]);/);
    if (!match) throw new Error('Data_netWorthTrend not found');
    const points = JSON.parse(match[1]);
    for (const point of points) {
      if (!Number.isFinite(Number(point.x)) || !Number.isFinite(Number(point.y))) continue;
      additions.push({
        date: shanghaiDate(point.x), fund_code: fund.fund_code,
        nav: Number(point.y), source: url, fetched_at: fetchedAt, is_estimate: false
      });
    }
    logs.push({ fetched_at: fetchedAt, fund_code: fund.fund_code, status: 'ok', rows: points.length, source: url, raw_sha256: crypto.createHash('sha256').update(text).digest('hex') });
  } catch (error) {
    logs.push({ fetched_at: fetchedAt, fund_code: fund.fund_code, status: 'failed-cached-data-retained', source: url, error: String(error.message ?? error) });
  }
  await new Promise((resolve) => setTimeout(resolve, 500));
}

const existing = readCsv(target);
const merged = new Map();
for (const row of [...existing, ...additions]) merged.set(`${row.date}|${row.fund_code}|${row.is_estimate}`, row);
const rows = [...merged.values()].sort((a, b) => a.date.localeCompare(b.date) || a.fund_code.localeCompare(b.fund_code));
writeCsv(target, rows, ['date','fund_code','nav','source','fetched_at','is_estimate']);
fs.appendFileSync(logFile, `${logs.map((item) => JSON.stringify(item)).join('\n')}${logs.length ? '\n' : ''}`, 'utf8');
const failures = logs.filter((item) => item.status !== 'ok');
process.stdout.write(JSON.stringify({ funds: project.funds.length, added: additions.length, failures: failures.length, cachedRows: existing.length }, null, 2) + '\n');
if (strict && failures.length) process.exitCode = 2;
