import fs from 'node:fs';
import path from 'node:path';

export function parseCsv(text) {
  const rows = [];
  let row = [], field = '', quoted = false;
  const input = String(text ?? '').replace(/^\uFEFF/, '');
  for (let i = 0; i < input.length; i += 1) {
    const ch = input[i];
    if (quoted) {
      if (ch === '"' && input[i + 1] === '"') { field += '"'; i += 1; }
      else if (ch === '"') quoted = false;
      else field += ch;
    } else if (ch === '"') quoted = true;
    else if (ch === ',') { row.push(field); field = ''; }
    else if (ch === '\n') { row.push(field.replace(/\r$/, '')); rows.push(row); row = []; field = ''; }
    else field += ch;
  }
  if (field.length || row.length) { row.push(field.replace(/\r$/, '')); rows.push(row); }
  if (!rows.length) return [];
  const headers = rows[0].map((value) => value.trim());
  return rows.slice(1).filter((values) => values.some((value) => value !== '')).map((values) =>
    Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ''])));
}

function csvCell(value) {
  const text = value == null ? '' : String(value);
  return /[",\n\r]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

export function toCsv(rows, headers) {
  const keys = headers ?? (rows[0] ? Object.keys(rows[0]) : []);
  return `${keys.map(csvCell).join(',')}\n${rows.map((row) => keys.map((key) => csvCell(row[key])).join(',')).join('\n')}${rows.length ? '\n' : ''}`;
}

export function readCsv(file) {
  return parseCsv(fs.readFileSync(file, 'utf8'));
}

export function writeCsv(file, rows, headers) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, toCsv(rows, headers), 'utf8');
}

export function loadProject(root) {
  const absolute = path.resolve(root);
  const read = (name) => readCsv(path.join(absolute, 'data', name));
  const readJson = (name, fallback = {}) => {
    const file = path.join(absolute, 'config', name);
    return fs.existsSync(file) ? JSON.parse(fs.readFileSync(file, 'utf8')) : fallback;
  };
  return {
    root: absolute,
    profile: readJson('profile.json'),
    tradingConstraints: readJson('trading-constraints.json'),
    noonDecisionRules: readJson('noon-decision-rules.json'),
    reviewLoop: readJson('review-loop.json'),
    funds: read('funds.csv'),
    fees: read('fees.csv'),
    nav: read('nav-history.csv'),
    holdings: read('holdings.csv'),
    transactions: read('transactions.csv'),
    portfolio: read('portfolio-history.csv'),
    decisions: fs.existsSync(path.join(absolute, 'data', 'decision-journal.csv')) ? read('decision-journal.csv') : []
  };
}

export const number = (value, fallback = 0) => {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
};
export const bool = (value) => String(value).toLowerCase() === 'true';

export function dayDiff(later, earlier) {
  const end = new Date(`${String(later).slice(0, 10)}T00:00:00Z`);
  const start = new Date(`${String(earlier).slice(0, 10)}T00:00:00Z`);
  return Math.round((end - start) / 86400000);
}

export function mean(values) {
  return values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : 0;
}

export function standardDeviation(values) {
  if (values.length < 2) return 0;
  const average = mean(values);
  return Math.sqrt(values.reduce((sum, value) => sum + (value - average) ** 2, 0) / (values.length - 1));
}

export function maxDrawdown(values) {
  let peak = -Infinity, worst = 0;
  for (const value of values) {
    peak = Math.max(peak, value);
    if (peak > 0) worst = Math.min(worst, value / peak - 1);
  }
  return Math.abs(worst);
}

export function feeForDays(fees, fundCode, holdingDays) {
  const rows = fees.filter((row) => row.fund_code === fundCode);
  const match = rows.find((row) => {
    const min = number(row.min_days);
    const max = row.max_days === '' ? Infinity : number(row.max_days, Infinity);
    return holdingDays >= min && holdingDays < max;
  });
  if (!match || !match.source_url || !match.verified_at || !/^https?:\/\//i.test(match.source_url)) return null;
  const purchaseRate = number(match.purchase_rate);
  const redemptionRate = number(match.redemption_rate);
  const serviceRate = number(match.sales_service_rate_annual) * holdingDays / 365;
  return {
    purchaseRate,
    redemptionRate,
    serviceRate,
    totalRate: purchaseRate + redemptionRate + serviceRate,
    sourceUrl: match.source_url,
    verifiedAt: match.verified_at
  };
}

function datedNav(navRows, fundCode, asOf) {
  return navRows
    .filter((row) => row.fund_code === fundCode && !bool(row.is_estimate) && row.date <= asOf && number(row.nav) > 0)
    .sort((a, b) => a.date.localeCompare(b.date));
}

export function scoreFund(fund, navRows, fees, profile, asOf) {
  const history = datedNav(navRows, fund.fund_code, asOf);
  const vetoes = [];
  const allowed = new Set(['index', 'index-enhanced', 'qdii-index']);
  const qdii = bool(fund.is_qdii);
  const plannedAmount = profile.capitalCny * profile.maxFundWeight;
  if (!allowed.has(fund.category)) vetoes.push('基金类型不在候选范围');
  if (number(fund.aum_cny) < profile.minAumCny) vetoes.push('规模低于1亿元');
  if (number(fund.age_days) < profile.minHistoryDays) vetoes.push('成立时间不足');
  if (!['open', 'limited'].includes(fund.purchase_status)) vetoes.push('不可申购或状态未知');
  if (fund.redemption_status !== 'open') vetoes.push('不可赎回或状态未知');
  if (fund.purchase_status === 'limited' && number(fund.daily_limit_cny) < plannedAmount) vetoes.push('申购限额不足');
  if (number(fund.lock_days) > 0) vetoes.push('存在强制锁定期');
  if (!fund.fee_source || !fund.fee_verified_at) vetoes.push('费率来源未核验');
  else if (dayDiff(asOf, fund.fee_verified_at) > 90) vetoes.push('费率核验超过90天');
  if (history.length < profile.minHistoryDays) vetoes.push('净值历史不足60个交易日');

  const latest = history.at(-1);
  const staleDays = latest ? Math.max(0, dayDiff(asOf, latest.date)) : Infinity;
  if (staleDays > (qdii ? profile.maxStaleDaysQdii : profile.maxStaleDaysDomestic)) vetoes.push('净值数据过期');
  const fee = feeForDays(fees, fund.fund_code, profile.targetHoldingDays);
  if (!fee) vetoes.push('目标持有期费率未知');

  if (history.length < 21 || !latest || !fee) {
    return { fund, eligible: false, score: -Infinity, vetoes, history, latest, staleDays, fee };
  }
  const values = history.map((row) => number(row.nav));
  const last = values.at(-1);
  const ret = (days) => last / values.at(-(days + 1)) - 1;
  const r5 = ret(5), r10 = ret(10), r20 = ret(20);
  const window = values.slice(-20);
  const ma20 = mean(window);
  const trend20 = last / ma20 - 1;
  const dailyReturns = values.slice(-21).slice(1).map((value, index) => value / values.slice(-21)[index] - 1);
  const volatility20 = standardDeviation(dailyReturns) * Math.sqrt(252);
  const drawdown20 = maxDrawdown(window);
  const momentum = 0.25 * r5 + 0.35 * r10 + 0.40 * r20;
  const trackingError = number(fund.tracking_error);
  const stalenessPenalty = qdii ? Math.max(0, staleDays - 2) * 0.005 : Math.max(0, staleDays - 1) * 0.01;
  const score = 100 * (momentum + 0.10 * trend20 - 0.10 * volatility20 - 0.50 * drawdown20 - trackingError - fee.totalRate - stalenessPenalty);
  if (trend20 <= 0) vetoes.push('最新净值未站上20日均线');
  if (r20 <= 0) vetoes.push('20日动量非正');
  if (momentum - fee.totalRate <= 0) vetoes.push('费用后动量非正');
  if (score <= 0) vetoes.push('综合分数非正');
  return {
    fund, eligible: vetoes.length === 0, score, vetoes, history, latest, staleDays, fee,
    return5: r5, return10: r10, return20: r20, trend20, volatility20, drawdown20,
    momentum, ma20, confidence: staleDays <= 1 && history.length >= 120 ? '高' : '中'
  };
}

export function scoreAll(project, asOf) {
  const results = project.funds.map((fund) => scoreFund(fund, project.nav, project.fees, project.profile, asOf));
  const cheapest = new Map();
  for (const result of results.filter((item) => item.fee)) {
    const key = result.fund.base_fund || result.fund.fund_code;
    if (!cheapest.has(key) || result.fee.totalRate < cheapest.get(key).fee.totalRate) cheapest.set(key, result);
  }
  for (const result of results) {
    const key = result.fund.base_fund || result.fund.fund_code;
    if (result.fee && cheapest.get(key)?.fund.fund_code !== result.fund.fund_code) {
      result.vetoes.push('同一母基金存在更低成本份额');
      result.eligible = false;
    }
  }
  return results.sort((a, b) => b.score - a.score);
}

export function portfolioRisk(portfolioRows, profile, asOf) {
  const rows = portfolioRows.filter((row) => row.date <= asOf && number(row.total_value) > 0)
    .sort((a, b) => a.date.localeCompare(b.date));
  if (!rows.length) return { total: profile.capitalCny, cash: profile.capitalCny, drawdown: 0, riskMode: false, cooldownRemaining: 0, warning: '缺少组合历史，按初始资金计算' };
  let peak = 0, lastTrigger = -Infinity;
  rows.forEach((row, index) => {
    const total = number(row.total_value);
    peak = Math.max(peak, total);
    if (peak > 0 && 1 - total / peak >= profile.accountDrawdownTrigger) lastTrigger = index;
  });
  const latest = rows.at(-1);
  const total = number(latest.total_value);
  const drawdown = peak > 0 ? 1 - total / peak : 0;
  const elapsed = rows.length - 1 - lastTrigger;
  const cooldownRemaining = Number.isFinite(lastTrigger) ? Math.max(0, profile.cooldownTradingDays - elapsed) : 0;
  return {
    total, cash: number(latest.cash_value), drawdown,
    riskMode: drawdown >= profile.accountDrawdownTrigger || cooldownRemaining > 0,
    cooldownRemaining, warning: ''
  };
}

function latestValue(result) {
  return result.latest ? number(result.latest.nav) : 0;
}

export function planActions(project, scores, asOf) {
  const profile = project.profile;
  const risk = portfolioRisk(project.portfolio, profile, asOf);
  const fundByCode = new Map(project.funds.map((fund) => [fund.fund_code, fund]));
  const scoreByCode = new Map(scores.map((score) => [score.fund.fund_code, score]));
  const holdings = project.holdings.filter((row) => number(row.shares) > 0);
  const actions = [];
  let invested = holdings.reduce((sum, row) => sum + number(row.market_value, number(row.shares) * latestValue(scoreByCode.get(row.fund_code))), 0);

  for (const holding of holdings) {
    const result = scoreByCode.get(holding.fund_code);
    const current = result ? latestValue(result) : 0;
    const heldDays = dayDiff(asOf, holding.confirmed_date);
    const positionReturn = current > 0 && number(holding.confirmed_nav) > 0 ? current / number(holding.confirmed_nav) - 1 : 0;
    const values = result?.history?.map((row) => number(row.nav)) ?? [];
    const belowTwoDays = values.length >= 21 && values.at(-1) < mean(values.slice(-20)) && values.at(-2) < mean(values.slice(-21, -1));
    if (positionReturn <= -profile.positionReviewDrawdown || belowTwoDays || risk.riskMode) {
      const fee = feeForDays(project.fees, holding.fund_code, Math.max(0, heldDays));
      const vetoes = [];
      const fund = fundByCode.get(holding.fund_code);
      if (heldDays < 7) vetoes.push('持有不足7日，默认禁止主动赎回');
      if (!fee) vetoes.push('当前持有期赎回费未由公开产品来源核验');
      if (!fund?.fee_source || !fund?.fee_verified_at || !/^https?:\/\//i.test(fund.fee_source)) vetoes.push('基金费率来源未核验');
      if (fund?.redemption_status !== 'open') vetoes.push('赎回状态未知或不可赎回');
      if (risk.warning) vetoes.push('缺少组合历史，无法确认账户回撤');
      const amount = number(holding.market_value, number(holding.shares) * current);
      actions.push({
        fund: fund ?? { fund_code: holding.fund_code, fund_name: '未知基金', share_class: '' },
        action: '退出审查', amount, redemptionShares: null, redemptionPercentage: 0, feeRate: fee?.redemptionRate ?? null,
        expectedNavDate: asOf, confirmDate: '以销售平台确认规则为准', deadline: '平台截止时间前提交；通常15:00前按申请日未知净值处理',
        evidence: `持仓收益${formatPct(positionReturn)}；${belowTwoDays ? '连续两日低于20日均线' : risk.riskMode ? '账户风险模式' : '触发4%回撤审查'}`,
        confidence: result?.confidence ?? '低', invalidation: '最新已公布净值恢复且风险触发解除', vetoes,
        executionNote: '退出审查不等于赎回指令；必须另行满足官方费率、赎回状态和午后情景阈值'
      });
    }
  }

  if (risk.riskMode) return { risk, actions, ranked: scores.filter((item) => item.eligible).slice(0, 3) };
  const ranked = scores.filter((item) => item.eligible);
  const heldCodes = new Set(holdings.map((row) => row.fund_code));
  const themes = new Map();
  let qdiiValue = 0;
  holdings.forEach((holding) => {
    const fund = fundByCode.get(holding.fund_code);
    const value = number(holding.market_value);
    if (fund) themes.set(fund.theme, (themes.get(fund.theme) ?? 0) + value);
    if (fund && bool(fund.is_qdii)) qdiiValue += value;
  });
  let available = Math.max(0, (risk.cash || risk.total - invested) - risk.total * profile.minCashWeight);
  let positionCount = holdings.length;
  for (const result of ranked) {
    if (positionCount >= profile.maxPositions || available < profile.minOrderCny) break;
    if (heldCodes.has(result.fund.fund_code)) continue;
    const themeRoom = risk.total * profile.maxThemeWeight - (themes.get(result.fund.theme) ?? 0);
    const qdiiRoom = bool(result.fund.is_qdii) ? risk.total * profile.maxQdiiWeight - qdiiValue : Infinity;
    const limit = result.fund.purchase_status === 'limited' ? number(result.fund.daily_limit_cny) : Infinity;
    const raw = Math.min(risk.total * profile.maxFundWeight, themeRoom, qdiiRoom, available, limit);
    const amount = Math.floor(raw / 100) * 100;
    if (amount < profile.minOrderCny) continue;
    actions.push({
      fund: result.fund, action: '申购草案', amount, feeRate: result.fee.totalRate,
      expectedNavDate: asOf,
      confirmDate: bool(result.fund.is_qdii) ? '通常T+2或更晚，以产品文件为准' : '通常T+1，以产品文件为准',
      deadline: '平台截止时间前提交；通常15:00前按申请日未知净值处理',
      evidence: `综合分${result.score.toFixed(2)}；5/10/20日收益${formatPct(result.return5)}/${formatPct(result.return10)}/${formatPct(result.return20)}`,
      confidence: result.confidence, invalidation: '净值跌回20日均线下、申购状态变化或费用后动量转负', vetoes: []
    });
    available -= amount;
    positionCount += 1;
    themes.set(result.fund.theme, (themes.get(result.fund.theme) ?? 0) + amount);
    if (bool(result.fund.is_qdii)) qdiiValue += amount;
  }
  return { risk, actions, ranked: ranked.slice(0, 3) };
}

export function signalRows(scores, asOf) {
  return scores.map((item) => ({
    as_of: asOf, fund_code: item.fund.fund_code, eligible: item.eligible,
    score: Number.isFinite(item.score) ? item.score.toFixed(6) : '',
    return_5d: item.return5?.toFixed(8) ?? '', return_10d: item.return10?.toFixed(8) ?? '',
    return_20d: item.return20?.toFixed(8) ?? '', trend_20d: item.trend20?.toFixed(8) ?? '',
    volatility_20d: item.volatility20?.toFixed(8) ?? '', drawdown_20d: item.drawdown20?.toFixed(8) ?? '',
    roundtrip_cost: item.fee?.totalRate?.toFixed(8) ?? '', stale_days: Number.isFinite(item.staleDays) ? item.staleDays : '',
    vetoes: item.vetoes.join('；')
  }));
}

export function formatPct(value) {
  return Number.isFinite(value) ? `${(value * 100).toFixed(2)}%` : '未知';
}

export function renderReport(project, scores, plan, asOf, mode) {
  const title = mode === 'morning' ? '晨间风险简报' : mode === 'weekly' ? '周复盘' : '15:00前行动简报';
  const now = new Date().toISOString();
  const cutoff = project.tradingConstraints?.defaultPlatformCutoff ?? '通常为开放日15:00，以平台为准';
  const lines = [`# ${asOf} ${title}`, '', `- 生成时间：${now}`, `- 数据截止：${asOf}；评分只使用已公布净值`, `- 账户回撤：${formatPct(plan.risk.drawdown)}；风险模式：${plan.risk.riskMode ? '是' : '否'}`, '- 账户范围：只交易场外开放式基金，不做盘中ETF或股票交易', `- 未知价原则：盘中数据只决定是否提交申请，实际按正式基金净值确认；默认截止：${cutoff}`, '- 性质：研究与执行草案，不是收益保证，不会自动下单', ''];
  if (plan.risk.warning) lines.push(`> 数据提示：${plan.risk.warning}`, '');
  if (plan.risk.riskMode) lines.push(`> 风险官否决新增申购：风险资产上限40%，冷静期剩余${plan.risk.cooldownRemaining}个有记录交易日。`, '');
  lines.push('## 候选排名', '');
  if (!plan.ranked.length) lines.push('暂不行动：没有基金同时通过数据、费用、趋势和风险检查。', '');
  else {
    lines.push('| 排名 | 基金/份额 | 分数 | 5日 | 10日 | 20日 | 费用 | 数据日期 |', '|---:|---|---:|---:|---:|---:|---:|---|');
    plan.ranked.forEach((item, index) => lines.push(`| ${index + 1} | ${item.fund.fund_name} ${item.fund.share_class}（${item.fund.fund_code}） | ${item.score.toFixed(2)} | ${formatPct(item.return5)} | ${formatPct(item.return10)} | ${formatPct(item.return20)} | ${formatPct(item.fee.totalRate)} | ${item.latest.date} |`));
    lines.push('');
  }
  lines.push('## 操作草案', '');
  if (!plan.actions.length) lines.push('暂不行动：当前没有通过风险官检查且需要执行的操作。', '');
  else plan.actions.forEach((action, index) => {
    const orderSize = action.action === '申购草案'
      ? `- 申购申请金额：${Math.floor(action.amount).toLocaleString('zh-CN')}元`
      : action.action === '退出审查'
        ? '- 赎回申请：暂不提交；退出审查不等于赎回指令'
      : `- 赎回申请：${action.redemptionShares ? `${action.redemptionShares.toLocaleString('zh-CN')}份（当前规则草案为${(action.redemptionPercentage * 100).toFixed(0)}%份额）` : '份额待平台确认页核对'}`;
    lines.push(`### ${index + 1}. ${action.fund.fund_name} ${action.fund.share_class}（${action.fund.fund_code}）`, '',
      `- 操作：${action.action}`, orderSize, `- 当前净值口径的市值估算：${Math.floor(action.amount).toLocaleString('zh-CN')}元（不是成交金额）`,
      `- 预计适用净值日：${action.expectedNavDate ?? '以提交时间与平台规则为准'}；成交净值提交时未知`,
      `- 预计确认：${action.confirmDate}`, `- 预计全部费用率：${action.feeRate == null ? '未知' : formatPct(action.feeRate)}`,
      `- 信号依据：${action.evidence}`, `- 置信度：${action.confidence}`, `- 最晚操作时间：${action.deadline}`,
      ...(action.executionNote ? [`- 执行说明：${action.executionNote}`] : []),
      `- 失效条件：${action.invalidation}`, `- 风险否决项：${action.vetoes.length ? action.vetoes.join('；') : '无'}`, '');
  });
  const rejected = scores.filter((item) => !item.eligible);
  lines.push('## 数据与否决审计', '', `- 候选总数：${scores.length}；通过：${scores.length - rejected.length}；否决：${rejected.length}`);
  rejected.slice(0, 10).forEach((item) => lines.push(`- ${item.fund.fund_code} ${item.fund.fund_name}：${item.vetoes.join('；') || '未知原因'}`));
  lines.push('', '## 来源', '', '- 净值来源逐条保存在 `data/nav-history.csv`。', '- 场外交易机制保存在 `config/trading-constraints.json`；产品合同、公告和销售平台特殊规则优先。', '- 费率来源逐档保存在 `data/fees.csv`；建议执行前仍须在销售平台核对。', '- 监管费率底线：https://www.csrc.gov.cn/csrc/c101954/c7606091/content.shtml', '');
  return lines.join('\n');
}

function navOnOrBefore(rows, code, date) {
  const candidates = rows.filter((row) => row.fund_code === code && !bool(row.is_estimate) && row.date <= date).sort((a, b) => a.date.localeCompare(b.date));
  return candidates.length ? number(candidates.at(-1).nav) : 0;
}

export function runBacktest(project) {
  const dates = [...new Set(project.nav.filter((row) => !bool(row.is_estimate)).map((row) => row.date))].sort();
  const startIndex = project.profile.minHistoryDays;
  if (dates.length <= startIndex + 1) return { observations: 0, status: '不建议实盘', reason: '历史数据不足', equity: [] };
  let cash = project.profile.capitalCny;
  const positions = new Map();
  const equity = [];
  const trades = [];
  let pending = null;
  for (let index = startIndex; index < dates.length; index += 1) {
    const date = dates[index];
    if (pending?.executionDate === date) {
      const desired = new Set(pending.codes);
      for (const [code, position] of [...positions]) {
        if (desired.has(code)) continue;
        const heldDays = dayDiff(date, position.entryDate);
        if (heldDays < 7) continue;
        const nav = navOnOrBefore(project.nav, code, date);
        const fee = feeForDays(project.fees, code, heldDays);
        if (!nav || !fee) continue;
        const gross = position.units * nav;
        cash += gross * (1 - fee.redemptionRate);
        trades.push({ signalDate: pending.signalDate, executionDate: date, code, side: 'sell', fee: gross * fee.redemptionRate });
        positions.delete(code);
      }
      const marked = [...positions].reduce((sum, [code, position]) => sum + position.units * navOnOrBefore(project.nav, code, date), 0);
      const total = cash + marked;
      for (const code of pending.codes) {
        if (positions.has(code)) continue;
        const nav = navOnOrBefore(project.nav, code, date);
        const fee = feeForDays(project.fees, code, project.profile.targetHoldingDays);
        if (!nav || !fee) continue;
        const budget = Math.min(total * project.profile.maxFundWeight, cash - total * project.profile.minCashWeight);
        if (budget < project.profile.minOrderCny) continue;
        const net = budget / (1 + fee.purchaseRate);
        positions.set(code, { units: net / nav, entryDate: date });
        cash -= budget;
        trades.push({ signalDate: pending.signalDate, executionDate: date, code, side: 'buy', fee: budget - net });
      }
      pending = null;
    }
    const marked = [...positions].reduce((sum, [code, position]) => sum + position.units * navOnOrBefore(project.nav, code, date), 0);
    equity.push({ date, value: cash + marked });
    if (!pending && index % project.profile.rebalanceEveryTradingDays === 0 && index + 1 < dates.length) {
      const scores = scoreAll(project, date).filter((item) => item.eligible).slice(0, project.profile.maxPositions);
      pending = { signalDate: date, executionDate: dates[index + 1], codes: scores.map((item) => item.fund.fund_code) };
    }
  }
  const values = equity.map((row) => row.value);
  const totalReturn = values.length ? values.at(-1) / values[0] - 1 : 0;
  const drawdown = maxDrawdown(values);
  const daily = values.slice(1).map((value, index) => value / values[index] - 1);
  const hitRate = daily.length ? daily.filter((value) => value > 0).length / daily.length : 0;
  const firstDate = dates[startIndex], lastDate = dates.at(-1);
  const poolReturns = project.funds.map((fund) => {
    const first = navOnOrBefore(project.nav, fund.fund_code, firstDate);
    const last = navOnOrBefore(project.nav, fund.fund_code, lastDate);
    return first > 0 && last > 0 ? last / first - 1 : null;
  }).filter((value) => value != null);
  const poolReturn = mean(poolReturns);
  const observations = equity.length;
  const qualifies = observations >= project.profile.backtestMinObservations && totalReturn > 0 && totalReturn > poolReturn && drawdown <= project.profile.accountDrawdownTrigger;
  return {
    observations, totalReturn, poolReturn, cashReturn: 0, maxDrawdown: drawdown, hitRate,
    turnover: trades.length, totalFees: trades.reduce((sum, trade) => sum + trade.fee, 0), trades, equity,
    status: qualifies ? '可进入小额观察期' : '不建议实盘',
    reason: qualifies ? '样本、费用后收益和回撤门槛均通过' : '至少一项样本、费用后收益、基准或回撤门槛未通过'
  };
}

export function renderBacktest(result, asOf) {
  return [`# ${asOf} 回测报告`, '', `- 结论：**${result.status}**`, `- 原因：${result.reason}`,
    `- 观察期：${result.observations}个交易日`, `- 策略费用后收益：${formatPct(result.totalReturn)}`,
    `- 候选池等权收益：${formatPct(result.poolReturn)}`, `- 现金基准：${formatPct(result.cashReturn)}`,
    `- 最大回撤：${formatPct(result.maxDrawdown)}`, `- 日胜率：${formatPct(result.hitRate)}`,
    `- 交易笔数：${result.turnover ?? 0}`, `- 模拟费用：${(result.totalFees ?? 0).toFixed(2)}元`, '',
    '> 信号在当日收盘数据上生成，最早于下一净值日执行；结果不构成未来收益保证。', ''].join('\n');
}
