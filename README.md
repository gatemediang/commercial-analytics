# Commercial Analytics & Financial Modelling Capstone (v2)

An Excel-first commercial analytics project built for a Data Analyst (Commercial) role in UK social care. It takes messy procurement, fleet, workforce, agency and finance extracts, cleans and reconciles them, and turns them into an interactive dashboard, a driver-based 2026 financial model and a savings tracker, with VBA to automate the monthly cycle. All data is synthetic.

## Headline results (2025 synthetic baseline)

| Measure | Value |
|---|---|
| Revenue / Opex / EBITDA | £86.5m / £60.4m / £26.1m (30.2% margin) |
| Procurement spend | £13.9m across 1,800 invoices, 24 suppliers |
| Off-contract spend | £11.5m (82%) |
| Data-quality exceptions | 6 of 1,800 invoices (0.3%), all logged and treated |
| 2026 EBITDA, Base / Conservative / Aggressive | £28.1m / £25.6m / £30.5m |
| Savings plan delivery | 81% of the £782k category target forecast |

The £5,662 gap between procurement spend and the finance ledger is one invoice with no site code, which the ledger could not post. One point of payroll inflation is worth £381k of EBITDA, 2.7 times a point of procurement saving.

## Repository

| Path | Contents |
|---|---|
| `Commercial_Analytics_Capstone_v2.xlsx` | The workbook (23 sheets, ~41,000 live formulas, 0 errors) |
| `Technical_Report.docx` / `.pdf` | 676-word technical report |
| `Dataset/` | Seven source CSVs plus three split procurement files for the consolidation macro |
| `VBA/` | Six `.bas` modules |
| `PowerQuery_SQL/` | M query and SQL checks mirroring the Excel logic |

## Workbook map

| Sheet | Purpose |
|---|---|
| README | Navigation with hyperlinks, colour key, five-minute demo |
| Dashboard | Region, category, month and scenario drop-downs; 15 KPI cards; 4 self-writing insight sentences; 8 charts |
| Financial_Model | Driver table (4 scenarios), 2026 forecast, scenario comparison, EBITDA bridge, monthly phasing, budget comparison, two sensitivity grids, tie-out checks |
| Savings_Tracker | Target vs realised YTD, run-rate forecast, RAG status, cross-check to the model |
| Site_Performance | Budget vs actual Opex and margin by site, ranked, RAG thresholds as inputs |
| Ops_Analysis | Fleet cost per mile, overtime %, absence per FTE, agency rates and share of staff cost |
| Supplier_Analysis | Spend, rank, Pareto top 10, contract compliance |
| Pivot_Analysis / Actual_Pivot | Formula cross-tabs and a genuine PivotTable used to validate them |
| Data_Quality | Exception counts and treatments, reconciliations with PASS/FAIL, completeness, live exception log |
| Cleaned_Procurement | Formula transform layer (`tblCleanProc`) |
| *_Raw, Sites, Category_Targets, Supplier_Lookup | Source tables, untouched except the supplier master rebuild |
| Dash_Calc | Filtered measures and chart series behind the Dashboard |
| Formula_Audit, VBA_Code, Methodology | Documentation |

Colour key: blue text on yellow is an input, green is a link to another sheet, black is a calculation.

## What changed from v1

| Issue found in v1 | Fix |
|---|---|
| `Cleaned_Procurement` headers said Contracted / DataQualityFlag but the formulas pulled VAT and NetSpend, so contracted spend showed £0 for every category | Columns rebuilt; contracted spend now reads the real flag |
| Blank spend produced `""`, so `Potential_Saving` returned `#VALUE!` and the Wales total on Pivot_Analysis errored | NetSpend_Clean is always numeric |
| Supplier lookup had 13 lower-case duplicates and mapped CloudLink to "medical supplies" | Rebuilt as a 24-row master |
| Category and supplier were never standardised in the clean layer | Case-insensitive INDEX/MATCH to the masters |
| Data-quality flags were pre-filled in the raw CSV rather than calculated | Flags now computed by formula, VBA and M independently |
| Savings tracker hard-coded realised savings as 85% of target, so every line read "Watch" | Realised YTD inputs, run-rate forecast and thresholds as inputs |
| Dashboard had four static KPIs and three charts, and no fleet, workforce, agency or site analysis existed | New interactive dashboard and four analysis sheets |
| The cleansing macro overwrote raw supplier and category text, and used COUNTIF per row | Array and Dictionary validation with a log; source text left alone |
| M query used `Text.Proper`, turning "IT & Telecoms" into "It & Telecoms" | Join to master lists instead |

## Using the dashboard

Open the Dashboard and pick from the four yellow drop-downs. Region, category and month filter the 2025 actuals; the scenario drives the 2026 model everywhere. No macros are needed for any of this.

## Installing the macros

1. Open the workbook in Excel desktop and save a copy as `.xlsm`.
2. Alt+F11, then File > Import File, and import all six modules from `VBA/` (they share a helper, so import all of them).
3. Alt+F8 > `AddDashboardButtons` once to draw the buttons.
4. Try `ConsolidateCSVFiles` on `Dataset/Sample_Procurement_Files` and `CleanseProcurementData` on a copy before live use.

| Macro | Job |
|---|---|
| CleanseProcurementData | Validates tblProcurement, writes flags and a DQ_Log sheet |
| ConsolidateCSVFiles | Appends same-header CSVs with a SourceFile column, reports mismatches |
| RefreshCommercialModel | Refresh, recalc, reconciliation status stamped on the Dashboard |
| ExportDashboardToPDF / ExportRegionalReportPack | One PDF, or one per region |
| RunAllScenarios | Logs all four scenarios to Scenario_Log |
| ResetDashboardFilters / AddDashboardButtons | Dashboard controls |

## Five-minute interview walkthrough

1. Dashboard: set Region to Wales and Month to Oct-2025, and read the insight sentences aloud.
2. Switch the scenario to Conservative and watch the bridge and scenario chart move.
3. Data_Quality: show the reused invoice ID and why it was kept, then the ledger reconciliation.
4. Financial_Model: the sensitivity grids and the payroll point.
5. Run `ExportRegionalReportPack`.

## GDPR and limitations

No names, addresses, NHS or payroll numbers or dates of birth. Workforce data is aggregated by site, month and role, and exported PDFs contain aggregates only. Figures are synthetic and do not represent CareTech. The model stops at EBITDA and ignores volume growth. The VBA was written for Excel desktop on Windows and has not been executed in this build environment, so test it on a copy first.
