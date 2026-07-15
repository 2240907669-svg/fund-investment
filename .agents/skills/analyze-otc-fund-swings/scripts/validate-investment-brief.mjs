import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export function validateInvestmentBrief(markdown, rules) {
  const gate = rules?.reportGate;
  if (!gate) throw new Error('缺少reportGate配置');

  const missingSections = gate.requiredSections.filter((section) => {
    const heading = new RegExp(`^#{1,6}\\s+${escapeRegExp(section)}(?:\\s|$)`, 'm');
    return !heading.test(markdown);
  });
  const missingPatterns = gate.requiredPatterns
    .filter(({ pattern }) => !new RegExp(pattern, 'i').test(markdown))
    .map(({ id, description }) => ({ id, description }));

  return {
    ok: missingSections.length === 0 && missingPatterns.length === 0,
    missingSections,
    missingPatterns
  };
}

function parseArgs(argv) {
  const args = [...argv];
  const reportPath = args.shift();
  let root = process.cwd();
  const rootIndex = args.indexOf('--root');
  if (rootIndex >= 0 && args[rootIndex + 1]) root = path.resolve(args[rootIndex + 1]);
  return { reportPath, root };
}

function main() {
  const { reportPath, root } = parseArgs(process.argv.slice(2));
  if (!reportPath) {
    console.error('用法: node validate-investment-brief.mjs <report.md> [--root <project-root>]');
    process.exitCode = 2;
    return;
  }

  const absoluteReport = path.resolve(reportPath);
  const rulesPath = path.join(root, 'config', 'agent-improvement-rules.json');
  const markdown = fs.readFileSync(absoluteReport, 'utf8');
  const rules = JSON.parse(fs.readFileSync(rulesPath, 'utf8'));
  const result = validateInvestmentBrief(markdown, rules);
  console.log(JSON.stringify({ report: absoluteReport, rules: rulesPath, ...result }, null, 2));
  if (!result.ok) process.exitCode = 1;
}

const isCli = process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1]);
if (isCli) main();
