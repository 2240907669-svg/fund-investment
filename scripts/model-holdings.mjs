#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';
import { readCsv, writeCsv, number, loadProject, scoreAll, planActions } from '../.agents/skills/analyze-otc-fund-swings/scripts/lib/engine.mjs';

const root = path.resolve(process.argv[2] || '.');
const assumptions = JSON.parse(fs.readFileSync(path.join(root, 'config', 'model-assumptions.json'), 'utf8'));
const intake = readCsv(path.join(root, 'data', 'holding-intake.csv'));
const nav = readCsv(path.join(root, 'data', 'nav-history.csv'))
  .filter((row) => row.is_estimate !== 'true' && number(row.nav) > 0)
  .sort((a, b) => a.date.localeCompare(b.date));

function navOnOrBefore(code, date) {
  const rows = nav.filter((row) => row.fund_code === code && row.date <= date);
  return rows.at(-1) || null;
}

const active = intake.filter((row) => !row.status.startsWith('sold_'));
const holdings = active.map((row) => {
  const snapshotNav = navOnOrBefore(row.fund_code, assumptions.snapshotAsOf);
  const latestNav = navOnOrBefore(row.fund_code, assumptions.asOf);
  if (!snapshotNav || !latestNav) throw new Error(`Missing NAV for ${row.fund_code}`);
  const shares = number(row.market_value) / number(snapshotNav.nav);
  const confirmedNav = number(row.estimated_cost) / shares;
  const marketValue = shares * number(latestNav.nav);
  return {
    fund_code: row.fund_code,
    shares: shares.toFixed(4),
    confirmed_nav: confirmedNav.toFixed(6),
    confirmed_date: assumptions.modeledPurchaseDate,
    market_value: marketValue.toFixed(2),
    source: `modeled:screenshot-${assumptions.snapshotAsOf}@NAV-${snapshotNav.date};latest-NAV-${latestNav.date}`,
    updated_at: new Date().toISOString()
  };
});

const feeRows = active.flatMap((row) => [
  {
    fund_code: row.fund_code, min_days: 0, max_days: 7, purchase_rate: 0,
    redemption_rate: assumptions.redemptionFeeRule.under7Days,
    sales_service_rate_annual: 0,
    source_url: 'model://user-declared-broad-fee-rule-not-official', verified_at: assumptions.asOf
  },
  {
    fund_code: row.fund_code, min_days: 7, max_days: '', purchase_rate: 0,
    redemption_rate: assumptions.redemptionFeeRule.from7Days,
    sales_service_rate_annual: 0,
    source_url: 'model://user-declared-broad-fee-rule-not-official', verified_at: assumptions.asOf
  }
]);

writeCsv(path.join(root, 'data', 'holdings.csv'), holdings,
  ['fund_code','shares','confirmed_nav','confirmed_date','market_value','source','updated_at']);
writeCsv(path.join(root, 'data', 'fees.csv'), feeRows,
  ['fund_code','min_days','max_days','purchase_rate','redemption_rate','sales_service_rate_annual','source_url','verified_at']);

const project = loadProject(root);
const scores = scoreAll(project, assumptions.asOf);
const plan = planActions(project, scores, assumptions.asOf);
const actionRows = plan.actions.map((action) => {
  const feeRate = action.feeRate ?? 0;
  return {
    as_of: assumptions.asOf,
    fund_code: action.fund.fund_code,
    fund_name: action.fund.fund_name,
    action: action.action,
    modeled_market_value: number(action.amount).toFixed(2),
    modeled_redemption_fee_rate: feeRate.toFixed(6),
    modeled_redemption_fee: (number(action.amount) * feeRate).toFixed(2),
    modeled_net_proceeds: (number(action.amount) * (1 - feeRate)).toFixed(2),
    evidence: action.evidence,
    invalidation: action.invalidation,
    assumption: 'Purchase date 2026-07-03; broad redemption fee rule supplied by user; not official product fee verification.'
  };
});
writeCsv(path.join(root, 'reports', `${assumptions.asOf}-modeled-redemption-plan.csv`), actionRows,
  ['as_of','fund_code','fund_name','action','modeled_market_value','modeled_redemption_fee_rate','modeled_redemption_fee','modeled_net_proceeds','evidence','invalidation','assumption']);

const total = holdings.reduce((sum, row) => sum + number(row.market_value), 0);
const redemptionFee = assumptions.holdingDays < 7
  ? assumptions.redemptionFeeRule.under7Days : assumptions.redemptionFeeRule.from7Days;
process.stdout.write(`${JSON.stringify({
  positions: holdings.length,
  modeledPurchaseDate: assumptions.modeledPurchaseDate,
  holdingDays: assumptions.holdingDays,
  redemptionFee,
  currentMarketValue: Number(total.toFixed(2)),
  estimatedFullRedemptionFee: Number((total * redemptionFee).toFixed(2)),
  estimatedNetProceeds: Number((total * (1 - redemptionFee)).toFixed(2)),
  modeledRedemptionDrafts: actionRows.length,
  modeledRedemptionDraftTotal: Number(actionRows.reduce((sum, row) => sum + number(row.modeled_market_value), 0).toFixed(2)),
  modeledRedemptionDraftNet: Number(actionRows.reduce((sum, row) => sum + number(row.modeled_net_proceeds), 0).toFixed(2))
}, null, 2)}\n`);
