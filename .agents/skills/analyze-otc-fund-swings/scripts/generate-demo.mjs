#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';
import { toCsv } from './lib/engine.mjs';

let output = 'examples/demo-project';
for (let i = 2; i < process.argv.length; i += 1) if (process.argv[i] === '--output') output = process.argv[++i];
const root = path.resolve(output);
fs.mkdirSync(path.join(root, 'config'), { recursive: true });
fs.mkdirSync(path.join(root, 'data'), { recursive: true });
fs.mkdirSync(path.join(root, 'reports'), { recursive: true });
const profile = {
  timezone:'Asia/Shanghai', capitalCny:30000, targetHoldingDays:14, minHistoryDays:60, minAumCny:100000000,
  maxPositions:3, maxFundWeight:0.25, maxThemeWeight:0.40, maxQdiiWeight:0.25, minCashWeight:0.25,
  positionReviewDrawdown:0.04, accountDrawdownTrigger:0.08, riskModeMaxInvestedWeight:0.40,
  cooldownTradingDays:5, maxStaleDaysDomestic:3, maxStaleDaysQdii:5, minOrderCny:100,
  rebalanceEveryTradingDays:5, backtestMinObservations:252
};
fs.writeFileSync(path.join(root, 'config', 'profile.json'), JSON.stringify(profile, null, 2) + '\n');
const funds = [
  ['900001','演示宽基A','A','demo-wide','index','宽基','false','500000000','800','open','open','50000','0','0.003','https://example.invalid/demo-wide-a','2026-06-01'],
  ['900002','演示宽基C','C','demo-wide','index','宽基','false','500000000','800','open','open','50000','0','0.003','https://example.invalid/demo-wide-c','2026-06-01'],
  ['900003','演示科技指数C','C','demo-tech','index','科技','false','300000000','600','open','open','50000','0','0.005','https://example.invalid/demo-tech','2026-06-01'],
  ['900004','演示QDII指数C','C','demo-global','qdii-index','海外','true','400000000','900','open','open','10000','0','0.006','https://example.invalid/demo-global','2026-06-01']
].map((v) => Object.fromEntries(['fund_code','fund_name','share_class','base_fund','category','theme','is_qdii','aum_cny','age_days','purchase_status','redemption_status','daily_limit_cny','lock_days','tracking_error','fee_source','fee_verified_at'].map((k,i)=>[k,v[i]])));
fs.writeFileSync(path.join(root,'data','funds.csv'), toCsv(funds));
const fees = [];
for (const fund of funds) {
  const purchase = fund.share_class === 'A' ? 0.001 : 0;
  const service = fund.share_class === 'C' ? 0.002 : 0;
  [['0','7','0.015'],['7','30',fund.category === 'index' || fund.category === 'qdii-index' ? '0.001' : '0.01'],['30','','0']].forEach(([min,max,redemption]) => fees.push({fund_code:fund.fund_code,min_days:min,max_days:max,purchase_rate:purchase,redemption_rate:redemption,sales_service_rate_annual:service,source_url:fund.fee_source,verified_at:'2026-06-01'}));
}
fs.writeFileSync(path.join(root,'data','fees.csv'), toCsv(fees));
const nav = [];
const start = new Date('2025-01-02T00:00:00Z');
let trading = 0;
for (let day = 0; trading < 380; day += 1) {
  const date = new Date(start); date.setUTCDate(date.getUTCDate() + day);
  if ([0,6].includes(date.getUTCDay())) continue;
  const dateText = date.toISOString().slice(0,10);
  const series = [
    ['900001',1 + 0.00065*trading + 0.018*Math.sin(trading/9)],
    ['900002',1 + 0.00065*trading + 0.018*Math.sin(trading/9)],
    ['900003',1 + 0.00085*trading + 0.035*Math.sin(trading/7)],
    ['900004',1 + 0.00055*trading + 0.025*Math.sin(trading/11)]
  ];
  for (const [code,value] of series) nav.push({date:dateText,fund_code:code,nav:value.toFixed(6),source:'synthetic-demo-not-market-data',fetched_at:'2026-07-12T08:00:00+08:00',is_estimate:'false'});
  trading += 1;
}
fs.writeFileSync(path.join(root,'data','nav-history.csv'), toCsv(nav));
fs.writeFileSync(path.join(root,'data','holdings.csv'), 'fund_code,shares,confirmed_nav,confirmed_date,market_value,source,updated_at\n');
fs.writeFileSync(path.join(root,'data','transactions.csv'), 'id,submitted_at,confirmed_at,fund_code,side,amount,shares,nav,fees,status\n');
const lastDate = nav.at(-1).date;
fs.writeFileSync(path.join(root,'data','portfolio-history.csv'), `date,total_value,cash_value,source\n${lastDate},30000,30000,synthetic-demo-not-market-data\n`);
fs.writeFileSync(path.join(root,'data','signals.csv'), 'as_of,fund_code,eligible,score,return_5d,return_10d,return_20d,trend_20d,volatility_20d,drawdown_20d,roundtrip_cost,stale_days,vetoes\n');
process.stdout.write(`${root}\n`);
