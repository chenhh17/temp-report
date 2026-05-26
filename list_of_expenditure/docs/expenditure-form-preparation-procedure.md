# Expenditure Form Preparation Procedure

This document records the full procedure used to prepare `TSSSU_Transaction List.xlsx` for TadReamk Limited, from the initial setup through bank-statement reconciliation and final checking.

## 1. Files Used

The working Excel file was `TSSSU_Transaction List.xlsx`.

The template/reference files were:

- `TSSSU_Transaction List_Template.xlsx`
- `TadReamk Limitted_TSSSU_Transaction List.xlsx`
- `TSSSU budget review(73).xlsx`
- `docs/how-to-complete-transaction-list.md`

The bank statement source files were:

- `/Users/chenhh/projects/tadreamk_erp_agent/collected/bank_statements/period_202504_to_202510_standardized.json`
- `/Users/chenhh/projects/tadreamk_erp_agent/collected/bank_statements/six_month_april2026_standardized.json`

## 2. Preparation Rules Applied

The transaction list was prepared using these rules:

- Use `TSSSU_Transaction List.xlsx` as the only working file.
- Keep the template file unchanged.
- Use the bank statement as the main source of truth for payment dates and actual HKD amounts.
- Use the budget review and last year’s file only for structure, category wording, and expected budget alignment.
- Use one row per supporting transaction unless several expenses are clearly reimbursed through one confirmed bank transfer.
- Use HKD amount from the bank statement when the bank statement gives a final HKD debit.
- Use the actual bank debit date instead of the invoice date when the row represents a reimbursement.
- Remove items that cannot be traced to the bank statement unless separate supporting evidence confirms the reimbursement.
- After every row insertion or deletion, manually check and fix all affected formulas because `openpyxl` does not automatically update Excel formulas.

## 3. Initial Sheet Setup

The company name was set to TadReamk Limited.

The transaction list was organized into three main blocks:

- Manpower
- Equipment
- Other direct costs

The columns used were:

- Supporting Ref #
- Date
- Supplier/employee
- Category in annex B
- Subcategory in annex B
- Description
- Currency
- Invoice amount
- Exchange rate
- Amount in HKD

The subtotal rows were kept as:

- Total Manpower
- Total Equipment
- Total Other direct costs
- Grand Total

## 4. Manpower Preparation

Manpower entries were reviewed against individual salary transfers and MPF assumptions.

CHAN, Laikiu remained recorded for 12 months:

- Salary: HKD 115,200
- MPF: HKD 5,760

SUN, Xiaoyuan remained recorded for April to June:

- Salary: HKD 65,190
- MPF: HKD 3,259.50

CHEN, Chun Hang remained recorded for September to March:

- Salary: HKD 72,000
- MPF: HKD 3,600

LI, Dexiong was re-verified against all bank transfers under `LI DEXIONG`.

The bank statement showed:

- 02/01/2026: HKD 5,696.00
- 02/02/2026: HKD 9,241.60
- 02/03/2026: HKD 7,052.80
- 30/03/2026: HKD 1,170.00
- 30/03/2026: HKD 3,822.00

The salary total used was HKD 26,982.40. The MPF amount was recalculated as HKD 1,349.12.

Small separate transfers of HKD 10.00 and HKD 65.60 were identified during the review. They were not added into the salary line because the main salary total was already bank-verified and the user confirmed Samuel’s expenses were correct under the current treatment.

LI, Peiqi was re-verified against individual bank transfers for the claimed period.

The bank statement showed:

- 02/01/2026: HKD 3,300.00
- 02/02/2026: HKD 2,925.00
- 02/03/2026: HKD 2,250.00

The salary total was updated to HKD 8,475.00. The MPF amount was recalculated as HKD 423.75.

## 5. Equipment Preparation

Equipment entries were reviewed against bank payments.

B1 Jumbo Computer Supplies was confirmed as:

- Date: 13/03/2026
- Amount: HKD 52,094.00

B2 Apple Shop was updated to match the bank statement date:

- Date changed from 30/04/2025 to 28/05/2025
- Amount: HKD 12,999.00

B3 Farm Computer Limited / TaoBao was reviewed. Although the description contains two purchases, the reimbursement appeared as a combined bank transfer, so the current sheet keeps it as one consolidated equipment line:

- Date: 02/10/2025
- Amount: HKD 18,313.43

## 6. Other Direct Costs Preparation

Other direct costs were heavily reconciled against bank statement records.

C1 Notion was first removed when it could not be traced directly. It was later re-added after the user provided supporting email evidence showing that it was part of a bulk reimbursement to SUN XIAO YUAN.

Final C1 treatment:

- Date: 14/05/2025
- Amount: HKD 3,525.38
- Basis: Included in HKD 4,391.68 reimbursement to SUN XIAO YUAN

C2 Telecom was kept because it was supported by repeated monthly reimbursements to CHAN, LAI KIU.

C3 WEBEYE CLOUD TECHNOL was corrected to the bank date:

- Date: 06/12/2025
- Amount: HKD 8,966.00

C4 MDC GITHUB, INC. was kept as a bank-confirmed card transaction:

- Date: 18/12/2025
- Amount: HKD 3,742.71

C5 Alibaba Cloud was kept as a bank-confirmed card transaction:

- Date: 23/12/2025
- Amount: HKD 6,326.96

C6 Apple Developer was first removed when it could not be traced directly. It was later re-added after the same SUN XIAO YUAN reimbursement email confirmed the Apple Developer payment.

Final C6 treatment:

- Date: 14/05/2025
- Amount: HKD 788.00

C7 Github was corrected to the bank statement date:

- Date: 10/07/2025
- Amount: HKD 391.47

C8 Amazon/TaoBao was corrected to the bank statement date:

- Date: 10/07/2025
- Amount: HKD 3,903.62

C9 Perplexity and C10 Amazon were kept after matching to reimbursement evidence:

- C9: HKD 1,600.72
- C10: HKD 640.30

C11 Notion Labs was re-searched and matched to SUN XIAO YUAN reimbursement:

- Date changed from 29/08/2025 to 15/09/2025
- Amount: HKD 669.06

C12 Cursor was reduced to only the verified March transaction:

- Kept: Cursor March, 30/03/2026, HKD 159.69
- Removed: Cursor January, no bank statement match
- Removed: Cursor February, no bank statement match

C13 Anthropic was split into two independent records:

- 13/03/2026: HKD 1,564.70
- 30/03/2026: HKD 1,564.70

C14 Kie.ai was reviewed against bank and reimbursement records:

- Kept: 23/12/2025, HKD 405.62
- Kept: 24/12/2025, HKD 405.31
- Kept: 12/01/2026, HKD 406.10
- Removed: February Kie.ai, no bank statement match

C15 Alibaba Cloud was used to close the budget gap using bank-confirmed charges:

- 07/01/2026: HKD 3,625.39
- 24/02/2026: HKD 6,738.50
- 04/03/2026: HKD 3,634.50

C16 Loyal Insurance Consultants was kept:

- Date: 17/07/2025
- Amount: HKD 6,922.19

C17 Global Consulting Services Limited was kept:

- Period: 1/4/2025-31/3/2026
- Amount: HKD 8,000.00

C18 Join Us CPA Limited was corrected to the May 2025 bank payment:

- Date: 22/05/2025
- Amount: HKD 11,000.00

C19 Hong Kong Trade Development Council was corrected to the bank date:

- Date: 16/09/2025
- Amount: HKD 10,850.00

C20 TRAN HOANG GIA LINH was kept as marketing and promotion expense:

- Period: 01/04/2025-08/08/2025
- Amount: HKD 26,000.00
- Supplier name was corrected to `TRAN HOANG GIA LINH`

## 7. Formula Maintenance

After each deletion or insertion, formulas were manually checked and fixed.

The final formula structure is:

- Manpower rows use `=H(row)*I(row)`
- Equipment rows use `=H(row)*I(row)`
- Other direct cost rows use `=H(row)*I(row)`
- Total Manpower: `=SUM(J12:J22)`
- Total Equipment: `=SUM(J25:J34)`
- Total Other direct costs: `=SUM(J38:J63)`
- Grand Total: `=SUM(J24,J37,J65)`

Rows were shifted several times during the process, especially in the Other direct costs section. The subtotal and grand total formulas were updated each time to point to the new final rows.

## 8. Final Audit Performed

The final audit checked:

- All populated rows had supporting reference numbers.
- Dates were consistent with the relevant bank evidence.
- Supplier names were in the correct column.
- HKD rows used exchange rate 1.
- Formula rows referenced the correct row numbers.
- Subtotal formulas covered the correct ranges.
- Grand total formula referenced the final subtotal rows.
- Known untraceable items had been removed.
- Bank-confirmed Alibaba Cloud charges had been re-added.
- Floating point display issues on Kie.ai rows were cleaned.

## 9. Final Totals

Final category totals:

- Manpower: HKD 302,239.77
- Equipment: HKD 83,406.43
- Other direct costs: HKD 113,834.92

Final grand total:

- HKD 499,481.12

Budget:

- HKD 500,000.00

Remaining gap:

- HKD 518.88

## 10. Final Status

The transaction list is bank-reconciled to the extent possible based on available statements and user-provided supporting evidence.

The remaining gap of HKD 518.88 is approximately 0.1% of the HKD 500,000 budget. The sheet was considered ready for submission after the final audit.
