# Agent improvement loop

This project borrows research architecture, not advertised returns. External systems generally test liquid stocks or crypto with simulated daily execution. They do not validate same-day execution for mainland OTC funds, where the order is submitted before the cutoff and the official NAV is unknown.

## Transferable design

1. **Specialize before synthesis.** Assign market breadth, news/catalysts, technicals, fund terms, and risk to separate evidence passes. A single allocator synthesizes them only after each pass records sources and unknowns.
2. **Scan opportunities before defending holdings.** Rank the whole required market universe by price strength, breadth, turnover, verified money flow, catalyst novelty, and crowding. Only then map leaders and laggards to holdings and eligible OTC funds.
3. **Debate independently.** The bull and bear cases are produced without seeing a preferred trade. Each must state what it explains, what it cannot explain, and one observable falsifier. Repeated wording is not independent evidence.
4. **Compare actions symmetrically.** Evaluate hold, cash, buy, and reduce using the same horizon and the same fee, confirmation-lag, concentration, and downside assumptions. Optimize fee-after value and drawdown control, not raw directional hit rate.
5. **Use layered memory.** Keep event memory in `data/catalyst-ledger.csv`, immutable decision memory in `data/decision-journal.csv`, confirmed portfolio outcomes in `data/portfolio-history.csv`, and permanent rules in versioned configuration. Retrieve only entries relevant to the current horizon and market regime.
6. **Reflect at two levels.** First test whether the information-to-price causal chain was right. Then test whether the resulting action added value versus doing nothing after real fees. A right market call with a bad OTC action is still an execution error.
7. **Fail loudly.** Missing fund status, fees, official NAV, portfolio history, or genuine money-flow data must be labeled `unknown` and can trigger a veto. Turnover and price are never relabeled as main money flow.

## Learning gate

- Preserve every original forecast and action before observing outcomes.
- Review against actual user execution when known, modeled execution under OTC rules, no action, and a broad benchmark.
- Track Brier score, fee-after value add, maximum adverse excursion, drawdown reduction, and sample size. Win rate alone rewards tiny correct calls and ignores large losses and costs.
- Treat each mistake as a hypothesis until the minimum resolved sample count in `config/review-loop.json` is met. Permanent changes require explicit evidence, no-lookahead testing, version notes, and unit tests.
- Do not promote a rule merely because it explains the latest day. Do not tune thresholds on the evaluation period.

## What is adopted from public research

- TradingAgents: specialized analysts, explicit bull/bear debate, trader synthesis, and risk management. Source: https://arxiv.org/abs/2412.20138 and https://github.com/TauricResearch/TradingAgents
- FinCon: manager-analyst hierarchy, source-specific analysis, within-period risk control, and outcome-based conceptual reflection. Source: https://arxiv.org/abs/2407.06567
- FinMem: layered memory with different horizons and retrieval priorities. Source: https://arxiv.org/abs/2311.13743
- FinAgent: multimodal market intelligence plus separate information-to-price and decision-to-outcome reflection. Source: https://arxiv.org/abs/2402.18485

Their reported results are backtests on selected assets and periods, not evidence of stable live excess returns. No reported win rate, return, or position rule is copied into this project.

## Daily failure audit

Before publication, answer these questions:

1. Which market leader would be omitted if the report only discussed current holdings?
2. Is every claimed flow backed by a named flow methodology and timestamp?
3. What is the best eligible buy, and why is it superior or inferior to cash today?
4. Did the bear case see the same evidence as the bull case and reach an independently reasoned conclusion?
5. Does the selected action still dominate after fees, unknown-price execution, confirmation lag, and concentration?
6. What single observable fact would overturn the conclusion before the platform cutoff?
