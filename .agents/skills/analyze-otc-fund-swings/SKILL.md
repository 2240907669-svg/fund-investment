---
name: analyze-otc-fund-swings
description: Analyze mainland-China off-exchange public fund swing opportunities with deterministic NAV scoring, share-class fee comparison, QDII staleness handling, portfolio limits, drawdown vetoes, event and catalyst analysis, scenario trees, backtests, teaching-oriented decision explanations, morning briefs, pre-cutoff action briefs, and weekly reviews. Use for 场外基金、基金短波、基金筛选、申购赎回建议、持仓风控、新闻催化、费率核算、基金回测 or portfolio review tasks in this project. Do not use to place trades, guarantee returns, bypass data access controls, or invent missing fees or prices.
---

# Analyze OTC Fund Swings

Treat this workflow as decision support, never as a promise of profit or an authorization to trade.

## Run the workflow

1. Read `../../../AGENTS.md`, `references/policy.md`, and `references/data-contracts.md` before changing strategy inputs or producing an action brief.
2. Read `references/event-driven-and-teaching.md` whenever news, policy, earnings, product launches, geopolitical events, market shocks, or teaching explanations affect the decision.
3. Run `node scripts/sync-nav.mjs --root ../../../..` when live NAV refresh is requested. Preserve provenance and continue with cached data if every public source fails.
4. Run `node scripts/analyze.mjs --root ../../../.. --mode morning` for the morning risk brief, `--mode action` for the pre-15:00 brief, or `--mode weekly` for the weekly review.
5. Run `node scripts/backtest.mjs --root ../../../..` before recommending that a new parameter set be used with real money.
6. Separate published NAV from intraday estimates. Never present an estimate as confirmed NAV.
7. Return `暂不行动` whenever fees, redemption status, data freshness, source consistency, or risk checks are unresolved.
8. Ask the user to confirm an executed order, then record it in `data/transactions.csv`; never log in to a sales platform or place an order.

## Apply the role chain

Perform the roles in this order and retain each role's evidence:

1. **Data steward**: validate freshness, provenance, trading status, fee verification, and conflicts.
2. **Signal analyst**: compute the deterministic 5/10/20-day score without future data; for catalysts, test event timing, surprise, market breadth, turnover, and relative strength.
3. **Risk officer**: apply eligibility, concentration, QDII, cash, holding-period, stop-review, and account-drawdown rules. State the strongest opposing explanation and the evidence that would invalidate the preferred view. The risk officer can veto any action.
4. **Execution planner**: express allowed actions as amounts and dates, including fees, scenario triggers, and invalidation conditions. Do not trade an opening gap when the decision can use evidence available before the fund cutoff.
5. **Reviewer**: compare recommendations with confirmed outcomes, score evidence quality and process discipline, and identify process errors without retroactively changing signals.

## Use the bundled resources

- Read `references/policy.md` for strategy and veto rules.
- Read `references/data-contracts.md` before editing CSV/JSON inputs or interpreting output fields.
- Read `references/data-sources.md` before adding or changing a network adapter.
- Read `references/event-driven-and-teaching.md` for catalyst analysis, red-team checks, scenario trees, and the teaching contract.
- Use `scripts/lib/engine.mjs` as the single source of truth for scoring, fees, eligibility, sizing, and backtests.
- Use Node.js built-in modules only. Do not add a package dependency without explicit user approval.

## Required response shape

For each recommendation include: fund/share class, action, suggested amount, expected confirmation date, estimated total fees, signal evidence, confidence, latest submission time, invalidation condition, and all risk vetoes. Include the data timestamp and source URLs near the claim they support.

For an event-driven decision, also show the verified facts, the market-confirmation evidence, the causal chain, the strongest alternative explanation, a probability-weighted scenario tree, and observable triggers for each action. Use ranges rather than false precision when the evidence is weak.

End important briefs with one reusable lesson: define the concept in plain language, show why it matters, apply it to one current holding, and state how the next review will test it. When no action passes, explain the vetoes and what new evidence would unblock the next run. Keep facts, calculations, estimates, and judgment visibly distinct.
