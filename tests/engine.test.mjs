import test from 'node:test';
import assert from 'node:assert/strict';
import {
  feeForDays, maxDrawdown, parseCsv, planActions, portfolioRisk, scoreAll, scoreFund, standardDeviation
} from '../.agents/skills/analyze-otc-fund-swings/scripts/lib/engine.mjs';

const profile = {
  capitalCny:30000,targetHoldingDays:14,minHistoryDays:60,minAumCny:100000000,maxPositions:3,maxFundWeight:0.25,
  maxThemeWeight:0.40,maxQdiiWeight:0.25,minCashWeight:0.25,positionReviewDrawdown:0.04,accountDrawdownTrigger:0.08,
  riskModeMaxInvestedWeight:0.40,cooldownTradingDays:5,maxStaleDaysDomestic:3,maxStaleDaysQdii:5,minOrderCny:100,
  rebalanceEveryTradingDays:5,backtestMinObservations:252
};

function fund(overrides={}) {
  return {fund_code:'000001',fund_name:'测试指数C',share_class:'C',base_fund:'base-1',category:'index',theme:'宽基',is_qdii:'false',
    aum_cny:'500000000',age_days:'500',purchase_status:'open',redemption_status:'open',daily_limit_cny:'50000',lock_days:'0',
    tracking_error:'0.001',fee_source:'https://example.invalid/fund',fee_verified_at:'2026-06-01',...overrides};
}

function nav(code='000001', count=90, finalDate='2026-07-12') {
  const end = new Date(`${finalDate}T00:00:00Z`);
  return Array.from({length:count},(_,i)=>{
    const date = new Date(end); date.setUTCDate(end.getUTCDate() - (count - 1 - i));
    return {date:date.toISOString().slice(0,10),fund_code:code,nav:String(1+i*0.004),source:'test',fetched_at:'2026-07-12T00:00:00Z',is_estimate:'false'};
  });
}

function fees(code='000001', purchase='0', middle='0.001', service='0.002') {
  return [
    {fund_code:code,min_days:'0',max_days:'7',purchase_rate:purchase,redemption_rate:'0.015',sales_service_rate_annual:service,source_url:'https://example.invalid/fee',verified_at:'2026-06-01'},
    {fund_code:code,min_days:'7',max_days:'30',purchase_rate:purchase,redemption_rate:middle,sales_service_rate_annual:service,source_url:'https://example.invalid/fee',verified_at:'2026-06-01'},
    {fund_code:code,min_days:'30',max_days:'',purchase_rate:purchase,redemption_rate:'0',sales_service_rate_annual:service,source_url:'https://example.invalid/fee',verified_at:'2026-06-01'}
  ];
}

test('CSV parser preserves quoted commas', () => {
  assert.deepEqual(parseCsv('a,b\n"x,y",z\n'), [{a:'x,y',b:'z'}]);
});

test('statistics calculate volatility and drawdown', () => {
  assert.equal(standardDeviation([1,1,1]),0);
  assert.ok(Math.abs(maxDrawdown([100,110,88,90]) - 0.2) < 1e-12);
});

test('fee schedule enforces short holding penalty and 14-day costs', () => {
  assert.equal(feeForDays(fees(),'000001',6).redemptionRate,0.015);
  const day14 = feeForDays(fees(),'000001',14);
  assert.equal(day14.redemptionRate,0.001);
  assert.ok(day14.totalRate > 0.001 && day14.totalRate < 0.002);
});

test('positive fresh history can pass while stale QDII is vetoed', () => {
  const passing = scoreFund(fund(),nav(),fees(),profile,'2026-07-12');
  assert.equal(passing.eligible,true);
  assert.ok(passing.score > 0);
  const qdii = scoreFund(fund({fund_code:'000002',base_fund:'base-2',category:'qdii-index',is_qdii:'true'}),nav('000002',90,'2026-07-01'),fees('000002'),profile,'2026-07-12');
  assert.equal(qdii.eligible,false);
  assert.ok(qdii.vetoes.includes('净值数据过期'));
});

test('cheaper share class wins within one base fund', () => {
  const project = {
    profile, funds:[fund({fund_code:'000001',share_class:'A'}),fund({fund_code:'000002',share_class:'C'})],
    fees:[...fees('000001','0.005','0.001','0'),...fees('000002','0','0.001','0.002')],
    nav:[...nav('000001'),...nav('000002')],holdings:[],transactions:[],portfolio:[]
  };
  const scores = scoreAll(project,'2026-07-12');
  assert.equal(scores.find((x)=>x.fund.fund_code==='000001').eligible,false);
  assert.ok(scores.find((x)=>x.fund.fund_code==='000001').vetoes.includes('同一母基金存在更低成本份额'));
});

test('8 percent account drawdown activates risk veto and blocks buys', () => {
  const portfolio = [
    {date:'2026-07-01',total_value:'30000',cash_value:'30000'},
    {date:'2026-07-10',total_value:'27500',cash_value:'27500'}
  ];
  const risk = portfolioRisk(portfolio,profile,'2026-07-12');
  assert.equal(risk.riskMode,true);
  const project = {profile,funds:[fund()],fees:fees(),nav:nav(),holdings:[],transactions:[],portfolio};
  const plan = planActions(project,scoreAll(project,'2026-07-12'),'2026-07-12');
  assert.equal(plan.actions.filter((x)=>x.action==='申购草案').length,0);
});

test('exit trigger remains a review and cannot create an unverified full redemption', () => {
  const unsafeFund = fund({redemption_status:'unknown',fee_source:'',fee_verified_at:''});
  const unsafeFees = fees().map((row) => ({...row,source_url:'model://unverified'}));
  const project = {
    profile,
    funds:[unsafeFund],
    fees:unsafeFees,
    nav:nav(),
    holdings:[{fund_code:'000001',shares:'1000',confirmed_nav:'2',confirmed_date:'2026-07-01',market_value:'1300'}],
    transactions:[],
    portfolio:[]
  };
  const plan = planActions(project,scoreAll(project,'2026-07-12'),'2026-07-12');
  assert.equal(plan.actions.length,1);
  assert.equal(plan.actions[0].action,'退出审查');
  assert.equal(plan.actions[0].redemptionPercentage,0);
  assert.equal(plan.actions[0].redemptionShares,null);
  assert.ok(plan.actions[0].vetoes.includes('当前持有期赎回费未由公开产品来源核验'));
  assert.ok(plan.actions[0].vetoes.includes('基金费率来源未核验'));
  assert.ok(plan.actions[0].vetoes.includes('赎回状态未知或不可赎回'));
  assert.ok(plan.actions[0].vetoes.includes('缺少组合历史，无法确认账户回撤'));
});

test('position and QDII caps constrain suggested buys', () => {
  const f1=fund(), f2=fund({fund_code:'000002',base_fund:'base-2',theme:'科技'}), f3=fund({fund_code:'000003',base_fund:'base-3',category:'qdii-index',theme:'海外',is_qdii:'true'});
  const project={profile,funds:[f1,f2,f3],fees:[...fees('000001'),...fees('000002'),...fees('000003')],nav:[...nav('000001'),...nav('000002'),...nav('000003')],holdings:[],transactions:[],portfolio:[{date:'2026-07-12',total_value:'30000',cash_value:'30000'}]};
  const plan=planActions(project,scoreAll(project,'2026-07-12'),'2026-07-12');
  assert.ok(plan.actions.length<=3);
  assert.ok(plan.actions.every((x)=>x.amount<=7500));
  assert.ok(plan.actions.filter((x)=>x.fund.is_qdii==='true').reduce((s,x)=>s+x.amount,0)<=7500);
});
