#!/usr/bin/env node

import fs from "node:fs";
import path from "node:path";

function argument(name, fallback) {
  const index = process.argv.indexOf(name);
  return index >= 0 && process.argv[index + 1] ? process.argv[index + 1] : fallback;
}

function parseCsvLine(line) {
  const values = [];
  let value = "";
  let quoted = false;
  for (let index = 0; index < line.length; index += 1) {
    const character = line[index];
    if (character === '"') {
      if (quoted && line[index + 1] === '"') {
        value += '"';
        index += 1;
      } else {
        quoted = !quoted;
      }
    } else if (character === "," && !quoted) {
      values.push(value);
      value = "";
    } else {
      value += character;
    }
  }
  values.push(value);
  return values;
}

function readCsv(filePath) {
  const lines = fs.readFileSync(filePath, "utf8").trim().split(/\r?\n/);
  if (lines.length === 0 || !lines[0]) return [];
  const headers = parseCsvLine(lines[0]);
  return lines.slice(1).filter(Boolean).map((line) => {
    const values = parseCsvLine(line);
    return Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ""]));
  });
}

const requestedRoot = path.resolve(argument("--root", process.cwd()));
const projectRoot = fs.realpathSync(requestedRoot);
const requiredFiles = [
  "AGENTS.md",
  "config/profile.json",
  "config/trading-constraints.json",
  "config/noon-decision-rules.json",
  "data/holdings.csv",
  "data/holding-intake.csv",
  "data/transactions.csv",
  "data/fees.csv",
  "data/funds.csv",
];

const missingFiles = requiredFiles.filter((relativePath) => !fs.existsSync(path.join(projectRoot, relativePath)));
if (missingFiles.length > 0) {
  console.error(JSON.stringify({
    status: "blocked",
    requestedRoot,
    projectRoot,
    missingFiles,
  }, null, 2));
  process.exit(1);
}

const profile = JSON.parse(fs.readFileSync(path.join(projectRoot, "config/profile.json"), "utf8"));
const holdings = readCsv(path.join(projectRoot, "data/holdings.csv"));
const transactions = readCsv(path.join(projectRoot, "data/transactions.csv"));
const portfolioHistory = fs.existsSync(path.join(projectRoot, "data/portfolio-history.csv"))
  ? readCsv(path.join(projectRoot, "data/portfolio-history.csv"))
  : [];
const totalMarketValueCny = holdings.reduce((total, holding) => total + Number(holding.market_value || 0), 0);
const latestHoldingUpdate = holdings.map((holding) => holding.updated_at).filter(Boolean).sort().at(-1) ?? null;
const warnings = [];

if (holdings.length === 0) warnings.push("持仓表没有记录");
if (transactions.length === 0) warnings.push("交易账本为空，无法核验真实成交与持有期");
if (portfolioHistory.length === 0) warnings.push("组合历史为空，无法计算真实账户回撤");
if (Number(profile.capitalCny) > 0 && Math.abs(totalMarketValueCny - Number(profile.capitalCny)) / Number(profile.capitalCny) > 0.1) {
  warnings.push("持仓市值与配置资金相差超过10%，仓位和现金比例不可直接采用配置值");
}

console.log(JSON.stringify({
  status: warnings.length === 0 ? "ready" : "ready_with_warnings",
  requestedRoot,
  projectRoot,
  requiredFileCount: requiredFiles.length,
  holdingCount: holdings.length,
  totalMarketValueCny: Number(totalMarketValueCny.toFixed(2)),
  configuredCapitalCny: Number(profile.capitalCny),
  confirmedTransactionCount: transactions.filter((transaction) => transaction.status === "confirmed").length,
  portfolioHistoryRows: portfolioHistory.length,
  latestHoldingUpdate,
  warnings,
}, null, 2));
