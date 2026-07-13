#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';
import { loadProject, renderBacktest, runBacktest } from './lib/engine.mjs';

let root = '.', asOf = '';
for (let i = 2; i < process.argv.length; i += 1) {
  if (process.argv[i] === '--root') root = process.argv[++i];
  else if (process.argv[i] === '--as-of') asOf = process.argv[++i];
}
const project = loadProject(root);
asOf ||= [...new Set(project.nav.map((row) => row.date))].sort().at(-1) || new Date().toISOString().slice(0, 10);
const result = runBacktest(project);
const report = renderBacktest(result, asOf);
const file = path.join(project.root, 'reports', `${asOf}-backtest.md`);
fs.mkdirSync(path.dirname(file), { recursive: true });
fs.writeFileSync(file, report, 'utf8');
process.stdout.write(`${file}\n`);
