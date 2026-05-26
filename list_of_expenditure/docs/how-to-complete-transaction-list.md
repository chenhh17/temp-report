# How to complete the TSSSU transaction list (using last year as reference)

## Files

| File | Role |
|------|------|
| `TSSSU_Transaction List_Template.xlsx` | **Do not edit.** Keep as the blank structure and examples. |
| `TSSSU_Transaction List.xlsx` | **Working file** for the current period: add all real transactions here. |
| `TadReamk Limitted_TSSSU_Transaction List.xlsx` | **Last year’s submitted list:** use only as a pattern for wording, categories, and how amounts are shown. |
| `TSSSU budget review(73).xlsx` | Budget review; use to align totals and line items with what you claimed, not as a substitute for the transaction list. |

## Sheet and header

- Open the **working** file → sheet **Transaction list**.
- Set **Company name** (cell next to the label) to your legal entity, as last year did for **TadReamk Limited**.

## Column mapping (last year → this year)

Last year’s workbook uses a slightly different header row. When you copy the *meaning* of a row forward, map fields like this:

| Last year (`TadReamk Limitted_…`) | This year (`TSSSU_Transaction List` / template) |
|----------------------------------|---------------------------------------------------|
| Supporting Ref # | Supporting Ref # (sequential: 1, 2, 3, …) |
| Date | Date (single payment date **or** keep the same text range if Excel allows; prefer one clear date per invoice when possible) |
| Supplier | **Supplier/employee** (payee: shop, vendor, or person name) |
| (no separate currency column) | **Currency** (e.g. HKD, USD, RMB—match the invoice) |
| Invoice amount | Invoice amount (numeric; use the amount **in the invoice currency**) |
| Exchange rate | Exchange rate (use **1** for HKD; otherwise the rate used to get HKD, as last year did for RMB rows) |
| Amount in HKD | Amount in HKD (final reported HKD) |
| Category in annex B | Category in annex B (e.g. Manpower, Equipment, Other direct costs—**same Annex B wording** as last year where still valid) |
| Description in annex B | Split across **Subcategory in annex B** + **Description** (see below) |
| Summary | **Not present** this year—omit. Do not rely on the odd numbers that appeared in some Summary cells last year. |

### Subcategory and description (this year)

This year separates **Subcategory in annex B** and **Description**. Follow the **Examples** block in the template:

- **Manpower:** use subcategories such as **Salary** vs **MPF** (template shows `Manpower` / `Salary` and salary text in **Description**).
- **Equipment / other costs:** use subcategories that match your Annex B breakdown (template example: `Equipment` / `Computer`).

From last year, rows such as “Full stack developer (1)” behave like **Salary** lines; “MPF to …” behave like **MPF** lines. Put the full annex wording you need in **Description** if it must match the funder’s schedule.

## Filling rules (aligned with last year’s examples)

1. **One row per supporting item** (one ref number), same as last year (refs 1–18 each had one line).

2. **Manpower:** **Supplier/employee** is the person paid (e.g. last year: employee name); **Amount in HKD** is the salary or MPF in HKD; **Category in annex B** = `Manpower`.

3. **Equipment:** vendor in **Supplier/employee**; **Category in annex B** = `Equipment` (e.g. laptop purchases).

4. **Other direct costs:** insurance, marketing, exhibitions, subscriptions, etc.—use **Other direct costs** (or the exact Annex B label you use today) and subcategories that match the current annex.

5. **Foreign currency (e.g. RMB):** Last year stored invoice amounts as text sometimes (e.g. “RMB2700”). This year prefer: **Currency** = RMB, **Invoice amount** = numeric (e.g. 2700), **Exchange rate** = the rate used, **Amount in HKD** = calculated HKD (as last year: e.g. 2700 × rate).

6. **Period text in Date:** Last year used text like `01/04/2024-31/03/2025` for some manpower lines. If your current template expects real dates, either use payroll **payment dates** or confirm with your admin whether a text range is still acceptable—**be consistent** across the sheet.

7. **Totals:** The template includes category blocks (e.g. **Total Manpower**) with formulas that must tie to **Annex B** (“Must equal to annex B”). After entering rows, check those totals and fix missing rows or wrong categories.

## Quality checks before submission

- [ ] **Company name** filled.
- [ ] Every data row has **Supporting Ref #**, **Date**, **Supplier/employee**, **Category in annex B**, **Amount in HKD**.
- [ ] **Currency** / **Invoice amount** / **Exchange rate** / **Amount in HKD** are consistent for non-HKD lines.
- [ ] No edits saved into `TSSSU_Transaction List_Template.xlsx`.
- [ ] Subtotals in the sheet match **Annex B** and, where applicable, the **budget review** workbook.
