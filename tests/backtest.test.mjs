import test from 'node:test';
import assert from 'node:assert/strict';
import { runBacktest } from '../.agents/skills/analyze-otc-fund-swings/scripts/lib/engine.mjs';

function project() {
  const profile={capitalCny:30000,targetHoldingDays:14,minHistoryDays:60,minAumCny:100000000,maxPositions:3,maxFundWeight:0.25,maxThemeWeight:0.4,maxQdiiWeight:0.25,minCashWeight:0.25,positionReviewDrawdown:0.04,accountDrawdownTrigger:0.08,riskModeMaxInvestedWeight:0.4,cooldownTradingDays:5,maxStaleDaysDomestic:3,maxStaleDaysQdii:5,minOrderCny:100,rebalanceEveryTradingDays:5,backtestMinObservations:252};
  const fund={fund_code:'000001',fund_name:'回测指数C',share_class:'C',base_fund:'base',category:'index',theme:'宽基',is_qdii:'false',aum_cny:'500000000',age_days:'500',purchase_status:'open',redemption_status:'open',daily_limit_cny:'50000',lock_days:'0',tracking_error:'0.001',fee_source:'https://example.invalid/fund',fee_verified_at:'2026-06-01'};
  const fees=[['0','7','0.015'],['7','30','0.001'],['30','','0']].map(([min,max,redemption])=>({fund_code:'000001',min_days:min,max_days:max,purchase_rate:'0',redemption_rate:redemption,sales_service_rate_annual:'0.002',source_url:'https://example.invalid/fee',verified_at:'2026-06-01'}));
  const nav=[]; const start=new Date('2025-01-01T00:00:00Z');
  for(let i=0;i<330;i+=1){const d=new Date(start);d.setUTCDate(d.getUTCDate()+i);nav.push({date:d.toISOString().slice(0,10),fund_code:'000001',nav:String(1+i*0.002),source:'test',fetched_at:'2026-07-12T00:00:00Z',is_estimate:'false'});}
  return {profile,funds:[fund],fees,nav,holdings:[],transactions:[],portfolio:[]};
}

test('backtest executes signals no earlier than the next NAV date',()=>{
  const result=runBacktest(project());
  assert.ok(result.observations>250);
  assert.ok(result.trades.length>0);
  assert.ok(result.trades.every((trade)=>trade.executionDate>trade.signalDate));
  assert.ok(Number.isFinite(result.totalReturn));
  assert.ok(Number.isFinite(result.maxDrawdown));
});
