# Borrowing Base Results

## Pro forma completed. Compare the current borrowing base against the staged scenario below.

### Scenario Checks

severity | field | message
--- | --- | ---
Warning | Leverage | Pro forma leverage is 13.68x, above the 6.5x threshold.

### Availability Impact

Metric | Before | After | Change | Notes
--- | --- | --- | --- | ---
BDC Investment Amount |  | $42,000,000 |  | Scenario input
Availability | $147,088,763 | $147,088,763 | $0 | Borrowing base availability
Aggregate Adjusted Borrowing Value | $232,996,704 | $232,996,704 | $0 | Availability tab adjusted collateral value
Excess Concentration Haircut | $2,295,693 | $2,295,693 | $0 | Concentration haircut applied by the model
Net Adjusted Borrowing Value | $230,701,011 | $230,701,011 | $0 | Adjusted borrowing value net of concentrations

### Portfolio Composition

Metric | Value | Calculated From
--- | --- | ---
Total SM Balance | $60,000,000 | Scenario input
BDC Share % | 70.0% | BDC Balance / Total SM Balance
Total Net Leverage | 13.68x | (Revolver + First Out + Pari + Total SM - Cash) / EBITDA
Net Detachment | 13.68x | All debt / EBITDA
Net Attachment | -2.11x | (Revolver + First Out - Cash) / EBITDA

### Concentration Limits

Limit Type | Limit % | Applicable Limit $ | Prior Actual $ | Prior Excess $ | Pro Forma Actual $ | Pro Forma Excess $ | Delta Actual $ | Delta Excess $
--- | --- | --- | --- | --- | --- | --- | --- | ---
1)a Max Second Lien and FILO with senior lev >= 1.50x | 20.0% | $46,599,341 | $4,121,303 | $0 | $4,121,303 | $0 | $0 | $0
1)b Max Second Lien | 10.0% | $23,299,670 | $0 | $0 | $0 | $0 | $0 | $0
2) Max Non-First Lien (2L and FILO with senior lev > 1.00x) | 30.0% | $69,899,011 | $53,940,435 | $0 | $53,940,435 | $0 | $0 | $0
3) Max EBITDA < $5MM | 15.0% | $34,949,506 | $0 | $0 | $0 | $0 | $0 | $0
4) Max Obligors | 7.5% | $17,474,753 | $19,770,446 | $2,295,693 | $19,770,446 | $2,295,693 | $0 | $0
5)a Max Largest Industry | 20.0% | $46,599,341 | $46,042,972 | $0 | $46,042,972 | $0 | $0 | $0
5)b Max Second Largest Industry | 15.0% | $34,949,506 | $24,394,399 | $0 | $24,394,399 | $0 | $0 | $0
5)c Max Other Industries | 10.0% | $23,299,670 | $22,141,696 | $0 | $22,141,696 | $0 | $0 | $0
6) Fixed Rate | 10.0% | $23,299,670 | $0 | $0 | $0 | $0 | $0 | $0
7) Max Limit Industry | 10.0% | $23,299,670 | $8,077,912 | $0 | $8,077,912 | $0 | $0 | $0
8) Max DDTL and Revolver | 15.0% | $34,949,506 | $0 | $0 | $0 | $0 | $0 | $0
9)a Max Non-Sponsor/Non-Family Office | 15.0% | $34,949,506 | $12,656,118 | $0 | $12,656,118 | $0 | $0 | $0
9)b Max Div Recap in Non-Sponsor/Non-Family Office | 10.0% | $23,299,670 | $0 | $0 | $0 | $0 | $0 | $0

### Workbook Metrics

Metric | Current | Pro Forma | Notes
--- | --- | --- | ---
Weighted Average Advance Rate | 61.4% | 61.4% | Availability!L48
Credit Enhancement Test | Pass | Pass | Availability!L42
Current Advances | $122,300,000 | $122,300,000 | Availability!L51

### Commentary

BluePeak Specialty Services, LLC scenario summary:
The scenario clears the base input gate and is ready for workbook-level execution once the Excel runner uses a trusted workbook copy.
Key warning flags:
- Leverage: Pro forma leverage is 13.68x, above the 6.5x threshold.
The pro forma was run on a staged workbook copy, so the master workbook was not changed.

Eligibility status: Yes