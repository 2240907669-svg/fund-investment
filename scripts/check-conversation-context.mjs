#!/usr/bin/env node

import fs from "node:fs";
import path from "node:path";

function argument(name, fallback = "") {
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
      if (quoted && line[index + 1] === '"') { value += '"'; index += 1; }
      else quoted = !quoted;
    } else if (character === "," && !quoted) { values.push(value); value = ""; }
    else value += character;
  }
  values.push(value);
  return values;
}

function readCsv(file) {
  if (!fs.existsSync(file)) return [];
  const lines = fs.readFileSync(file, "utf8").trim().split(/\r?\n/).filter(Boolean);
  if (!lines.length) return [];
  const headers = parseCsvLine(lines[0]);
  return lines.slice(1).map((line) => {
    const values = parseCsvLine(line);
    return Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ""]));
  });
}

function shanghaiTimestamp() {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Shanghai", year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", minute: "2-digit", second: "2-digit", hourCycle: "h23"
  }).formatToParts(new Date());
  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${values.year}-${values.month}-${values.day}T${values.hour}:${values.minute}:${values.second}+08:00`;
}

function ageDays(openedAt, now) {
  const milliseconds = new Date(now).getTime() - new Date(openedAt).getTime();
  return Math.max(0, Math.floor(milliseconds / 86400000));
}

const requestedRoot = path.resolve(argument("--root", process.cwd()));
const root = fs.realpathSync(requestedRoot);
const role = argument("--role", "all");
const recordRun = process.argv.includes("--record-run");
const config = JSON.parse(fs.readFileSync(path.join(root, "config", "context-rotation.json"), "utf8"));
const registryPath = path.join(root, config.registryFile);
const registry = JSON.parse(fs.readFileSync(registryPath, "utf8"));
const now = shanghaiTimestamp();

if (role !== "all" && !registry.threads[role]) {
  console.error(JSON.stringify({ status: "blocked", reason: `unknown role: ${role}` }, null, 2));
  process.exit(1);
}

if (recordRun && role !== "all") {
  registry.threads[role].trackedRunCount = Number(registry.threads[role].trackedRunCount || 0) + 1;
  registry.threads[role].lastRunAt = now;
}

const roles = role === "all" ? Object.keys(registry.threads) : [role];
const checks = roles.map((key) => {
  const thread = registry.threads[key];
  const rules = config.roles[key];
  const days = ageDays(thread.openedAt, now);
  const reasons = [];
  if (Number(thread.trackedRunCount) >= Number(rules.maxTrackedRuns)) reasons.push(`已跟踪运行${thread.trackedRunCount}次，阈值${rules.maxTrackedRuns}次`);
  if (days >= Number(rules.maxCalendarDays)) reasons.push(`任务已使用${days}天，阈值${rules.maxCalendarDays}天`);
  const recommended = reasons.length > 0 || thread.rotationStatus === "rotation_recommended";
  if (recommended) thread.rotationStatus = "rotation_recommended";
  return { role: key, label: rules.label, threadId: thread.threadId, trackedRunCount: thread.trackedRunCount, ageDays: days, rotationRecommended: recommended, reasons };
});

if (recordRun) fs.writeFileSync(registryPath, `${JSON.stringify(registry, null, 2)}\n`, "utf8");

const holdings = readCsv(path.join(root, "data", "holdings.csv"));
const funds = readCsv(path.join(root, "data", "funds.csv"));
const transactions = readCsv(path.join(root, "data", "transactions.csv"));
const portfolio = readCsv(path.join(root, "data", "portfolio-history.csv"));
const fundNames = new Map(funds.map((fund) => [fund.fund_code, fund.fund_name]));
const totalValue = holdings.reduce((sum, holding) => sum + Number(holding.market_value || 0), 0);
const latestHoldingUpdate = holdings.map((holding) => holding.updated_at).filter(Boolean).sort().at(-1) ?? "未知";
const reportDir = path.join(root, "reports");
const latestReports = fs.existsSync(reportDir)
  ? fs.readdirSync(reportDir).filter((name) => /\.(md|docx)$/.test(name)).map((name) => ({ name, time: fs.statSync(path.join(reportDir, name)).mtimeMs })).sort((a, b) => b.time - a.time).slice(0, 6)
  : [];
const warnings = [];
if (!portfolio.length) warnings.push("组合历史为空，无法计算真实账户回撤");
const configuredCapital = JSON.parse(fs.readFileSync(path.join(root, "config", "profile.json"), "utf8")).capitalCny;
if (Number(configuredCapital) > 0 && Math.abs(totalValue - Number(configuredCapital)) / Number(configuredCapital) > 0.1) warnings.push("持仓市值与配置资金相差超过10%，不得把配置资金当真实账户总额");

const holdingLines = holdings.map((holding) => `- ${fundNames.get(holding.fund_code) ?? "未知基金"}：${Number(holding.market_value || 0).toFixed(2)}元；${holding.shares || "未知"}份；更新于${holding.updated_at || "未知"}`);
const transactionLines = transactions.slice(-8).map((transaction) => `- ${transaction.id}：${fundNames.get(transaction.fund_code) ?? "未知基金"} ${transaction.side === "sell" ? "赎回" : "申购"} ${transaction.shares ? `${transaction.shares}份` : `${transaction.amount || "未知"}元`}；状态${transaction.status || "未知"}`);
const checkLines = checks.map((check) => `- ${check.label}：已跟踪${check.trackedRunCount}次，任务日龄${check.ageDays}天，轮换${check.rotationRecommended ? "建议" : "未触发"}${check.reasons.length ? `（${check.reasons.join("；")}）` : ""}`);

const handoff = [
  "# 基金项目跨任务交接快照", "", `- 自动刷新：${now}`, "- 用途：为新任务提供最小充分入口；所有数值仍须回到原始文件核验", `- 项目自检警告：${warnings.length ? warnings.join("；") : "无"}`, `- 在持基金：${holdings.length}只；记录市值合计${totalValue.toFixed(2)}元；持仓最近更新${latestHoldingUpdate}`, "",
  "## 对话轮换状态", "", ...checkLines, "", "用户口令：`轮换基金对话`。收到后按“新建—迁移—自检—确认—归档”执行，不能先归档。", "",
  "## 当前持仓索引", "", ...holdingLines, "", "## 最近确认/申报交易", "", ...(transactionLines.length ? transactionLines : ["- 暂无记录"]), "",
  "## 当前风险警告", "", ...(warnings.length ? warnings.map((warning) => `- ${warning}`) : ["- 无"]), "",
  "## 最新报告", "", ...latestReports.map((report) => `- reports/${report.name}`), "",
  "## 新任务强制读取顺序", "", "1. `/Users/dengzhengxin/Documents/基金/AGENTS.md`", "2. 本交接快照", "3. `fund-investment/AGENTS.md`", "4. `data/holdings.csv`、`data/transactions.csv`、`data/fees.csv`、`data/funds.csv`", "5. 与请求相关的最新报告和配置", "6. 重新运行 `node fund-investment/scripts/check-project-context.mjs --root fund-investment`", "",
  "## 不从聊天继承的事项", "", "- 未写入交易账本的操作不能视为已成交。", "- 盘中估算不能视为正式净值。", "- 旧报告建议不能自动覆盖最新持仓、交易和风险否决。", ""
].join("\n");
fs.mkdirSync(path.dirname(path.join(root, config.handoffFile)), { recursive: true });
fs.writeFileSync(path.join(root, config.handoffFile), handoff, "utf8");

const anyRecommended = checks.some((check) => check.rotationRecommended);
console.log(JSON.stringify({ status: anyRecommended ? "rotation_recommended" : "context_ok", checkedAt: now, handoffFile: path.join(root, config.handoffFile), checks }, null, 2));
