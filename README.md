# Commercial Analytics & Financial Modelling Capstone

## 1. Project overview

This project is an end-to-end **commercial data analytics and financial modelling solution built in Microsoft Excel**. It simulates the type of analytical workflow found in commercial, finance, procurement, workforce and operational reporting teams.

The project starts with fragmented operational datasets and finishes with:

- cleansed and validated commercial data;
- category, supplier and regional spend analysis;
- fleet, workforce and agency-cost analysis;
- a driver-based financial forecast;
- procurement savings modelling;
- an executive Excel dashboard;
- repeatable VBA automation for cleansing, CSV consolidation, refresh and PDF reporting; and
- supporting Power Query/M and SQL examples.

The dataset is **100% synthetic**. It contains no real employee, service-user, customer or supplier personal data and is intended for portfolio, interview and learning use.

---

## 2. Business problem

A commercial analytics team receives information from several operational and finance sources. The files do not necessarily use consistent category names, may contain duplicate records or missing fields, and must be converted into reliable management information.

The analyst needs to answer questions such as:

1. How much are we spending, and where?
2. Which procurement categories represent the largest opportunities?
3. Which suppliers and sites drive expenditure?
4. How are fleet, workforce and agency costs changing?
5. Where are actual costs different from budget or expected performance?
6. What savings could be achieved under different assumptions?
7. What would those savings and cost changes do to EBITDA?
8. Can the recurring workflow be automated rather than performed manually every month?

The project therefore follows the analytical pipeline:

**Source data → Data quality → Transformation → Analysis → Financial model → Dashboard → Automation**

---

## 3. Solution architecture

```text
CSV source datasets
       │
       ├── Procurement
       ├── Fleet
       ├── Workforce
       ├── Agency
       ├── Finance
       ├── Sites
       └── Category targets
              │
              ▼
       Excel Tables / Lookups
              │
              ▼
       Cleaned_Procurement
              │
       ┌──────┴─────────┐
       ▼                ▼
Commercial Analysis   Financial Model
       │                │
       └──────┬─────────┘
              ▼
        Savings Tracker
              │
              ▼
          Dashboard
              │
              ▼
        VBA Automation
```

The design intentionally separates **raw data**, **transformation**, **analysis**, **modelling** and **presentation** so that the workbook is easier to audit and maintain.

---

## 4. Repository structure

| File | Purpose |
|---|---|
| `README.md` | Full project documentation and operating instructions. |
| `Commercial_Analytics_Financial_Modelling_Capstone.xlsx` | Main Excel workbook containing raw data, formulas, modelling, dashboard and analysis. The supplied workbook may retain its original filename; the project is deliberately generic. |
| `01_Procurement_Spend.csv` | Synthetic invoice-level procurement spend. |
| `02_Fleet.csv` | Synthetic monthly fleet utilisation and cost data. |
| `03_Workforce.csv` | Synthetic monthly workforce/payroll data. |
| `04_Agency.csv` | Synthetic agency hours, rates and spend. |
| `05_Finance.csv` | Synthetic monthly revenue and operating-cost data. |
| `06_Sites.csv` | Site master/lookup table. |
| `07_Category_Targets.csv` | Category-level savings assumptions and owners. |
| `Module_CleanseData.bas` | VBA data cleansing and quality-control macro. |
| `Module_ConsolidateCSVs.bas` | VBA folder-based CSV consolidation macro. |
| `Module_RefreshDashboard.bas` | VBA workbook refresh/calculation macro. |
| `Module_ExportReport.bas` | VBA dashboard-to-PDF reporting macro. |
| `PowerQuery_Clean_Procurement.m` | Power Query/M example for procurement cleansing and quality flags. |
| `SQL_Commercial_Checks.sql` | Optional SQL checks for commercial data validation and aggregation. |
| `Technical_Report_*.docx` / `.pdf` | Short technical summary supplied with the original workbook. The README is the authoritative implementation guide. |

> **Note:** If your local copy of the workbook still has its original filename, do not change worksheet names or Excel Table names unless you also update formulas and VBA references.

---

## 5. Dataset documentation

### 5.1 Procurement

`01_Procurement_Spend.csv` contains approximately 1,800 invoice-level records.

Key fields include:

- `InvoiceDate`
- `SiteID`
- `SiteName`
- `Region`
- `Supplier`
- `Category`
- `InvoiceID`
- `Quantity`
- `UnitCost`
- `GrossSpend`
- `PaymentTermsDays`
- `SourceSystem`
- contract/status and derived fields in the workbook

The procurement data deliberately contains quality exceptions, including inconsistent category casing and records suitable for duplicate/missing-value testing.

### 5.2 Fleet

`02_Fleet.csv` contains monthly fleet activity by site and vehicle.

Important measures include:

- miles travelled;
- fuel litres;
- fuel cost;
- maintenance cost;
- lease cost; and
- total fleet cost.

This supports cost-per-mile and fleet-efficiency analysis.

### 5.3 Workforce

`03_Workforce.csv` contains monthly workforce measures by site and role.

Important measures include:

- FTE;
- annual salary;
- overtime hours;
- overtime cost;
- absence days;
- base payroll; and
- total payroll.

No employee names or personal identifiers are included.

### 5.4 Agency

`04_Agency.csv` contains monthly agency labour activity by site, agency and role.

The principal measures are:

- hours;
- hourly rate; and
- agency spend.

This allows agency-cost and rate-inflation scenarios to be modelled.

### 5.5 Finance

`05_Finance.csv` provides the financial baseline by site and month.

Measures include:

- revenue;
- payroll;
- procurement;
- fleet;
- agency;
- other operating expenditure;
- budget Opex; and
- actual Opex.

### 5.6 Reference tables

`06_Sites.csv` provides site IDs, names, regions and towns.

`07_Category_Targets.csv` provides savings targets and ownership by procurement category.

The workbook also contains a `Supplier_Lookup` table used for supplier/category and contracted-status enrichment.

---

## 6. Excel workbook guide

The workbook contains the following important worksheets.

### `README`

A compact in-workbook overview. This GitHub README is more detailed and should be treated as the main documentation.

### `Procurement_Raw`

The source procurement table. It is stored as Excel Table **`tblProcurement`**.

Do not overwrite the table headers if you intend to run the supplied VBA code.

### `Fleet_Raw`

Fleet source table: **`tblFleet`**.

### `Workforce_Raw`

Workforce source table: **`tblWorkforce`**.

### `Agency_Raw`

Agency source table: **`tblAgency`**.

### `Finance_Raw`

Finance source table: **`tblFinance`**.

### `Sites`

Site master table: **`tblSites`**.

### `Category_Targets`

Savings assumptions: **`tblTargets`**.

### `Supplier_Lookup`

Supplier/category/contract-status lookup: **`tblSuppliers`**.

### `Cleaned_Procurement`

The transformed procurement layer. It is stored as **`tblCleanProc`** and is used by downstream analysis.

This layer separates source records from analytical calculations and contains derived fields such as cleaned spend and data-quality information.

### `Pivot_Analysis`

Formula-driven commercial analysis using functions including:

- `SUMIFS`
- `COUNTIFS`
- `IFERROR`
- `INDEX` + `MATCH`

It summarises spend by category, calculates percentage contribution, applies savings targets and estimates potential savings.

### `Actual_Pivot`

A PivotTable-style analysis of procurement spend by category and region. It provides a familiar management-reporting view for validating the formula-based analysis.

### `Financial_Model`

The core financial modelling sheet.

It uses 2025 actuals as the baseline and applies scenario drivers for 2026:

| Driver | Conservative | Base | Aggressive |
|---|---:|---:|---:|
| Revenue growth | 1% | 3% | 5% |
| Procurement savings | 2% | 4% | 7% |
| Agency rate inflation | 6% | 5% | 3% |
| Fleet efficiency | 1% | 3% | 5% |
| Payroll inflation | 3.5% | 2.5% | 2% |
| Other Opex inflation | 4% | 3% | 2% |

The selected scenario is controlled by the scenario cell on the model sheet.

The model calculates:

- 2026 revenue;
- procurement cost;
- agency cost;
- fleet cost;
- payroll;
- other Opex;
- total Opex;
- EBITDA; and
- EBITDA margin.

The scenario-selection formulas use `MATCH` and `CHOOSE`, which demonstrates a practical approach to scenario-driven modelling without requiring external software.

### `Savings_Tracker`

Calculates category-level:

**2025 spend × target savings % = target saving**

It then models realised savings and assigns a status:

- **On Target**
- **Watch**
- **Off Track**

### `Dashboard`

The executive reporting layer.

It surfaces:

- total revenue;
- actual Opex;
- EBITDA;
- potential savings;
- largest spend category;
- highest savings opportunity; and
- EBITDA margin.

Charts are included to make the output suitable for management discussion rather than presenting raw tables alone.

### `Formula_Audit`

A compact catalogue explaining the advanced formulas used in the model and where they are appropriate.

### `VBA_Code`

Documents the purpose of the supplied VBA modules and the import/run process.

### `Methodology`

Documents data provenance, quality assumptions, GDPR-first design, financial-model assumptions and limitations.

---

## 7. Advanced Excel techniques demonstrated

### Lookup and retrieval

The model uses lookup logic to enrich and connect datasets. The live model uses `INDEX` + `MATCH` where compatibility and explicit matching are useful; lookup alternatives such as `XLOOKUP` can be substituted in Microsoft 365.

### Conditional aggregation

`SUMIFS` is used for category and savings calculations, while `COUNTIFS` supports record-level validation and counting.

Example:

```excel
=SUMIFS(Cleaned_Procurement!$O:$O,
        Cleaned_Procurement!$F:$F,A5)
```

### Error handling

`IFERROR` prevents missing lookup keys from producing user-facing errors in management reports.

Example:

```excel
=IFERROR(INDEX(Category_Targets!$B$2:$B$9,
MATCH(A4,Category_Targets!$A$2:$A$9,0)),0)
```

### Scenario modelling

The selected scenario is converted into the relevant driver using `MATCH` + `CHOOSE`.

Example:

```excel
=CHOOSE(MATCH($B$3,$B$5:$D$5,0),B6,C6,D6)
```

### Financial modelling

The model follows a driver-based structure rather than hard-coding a forecast. This makes sensitivity analysis possible and creates a clear audit trail from assumption → forecast → profitability.

### Data validation and conditional formatting

The workbook uses Excel's presentation and control features to highlight quality issues, model status and commercial opportunities.

### Pivot analysis

Pivot analysis provides a second method of validating the formula-driven calculations and demonstrates practical management-information reporting.

---

## 8. VBA automation

The VBA layer demonstrates how repetitive Excel operations can be converted into controlled, repeatable processes.

### Module 1 — `Module_CleanseData.bas`

**Macro:** `CleanseProcurementData`

Purpose:

1. Reads `tblProcurement`.
2. Trims supplier, site and category values.
3. Standardises known category variants.
4. Checks missing supplier/site/invoice/spend values.
5. Checks duplicate invoice IDs.
6. Writes a `DataQualityFlag` for each record.

This is the most important macro for demonstrating data-quality automation.

### Module 2 — `Module_ConsolidateCSVs.bas`

**Macro:** `ConsolidateCSVFiles`

Purpose:

1. Prompts the analyst to select a folder.
2. Finds CSV files in that folder.
3. Opens each CSV.
4. Appends its data into a consolidated worksheet.
5. Closes the source workbook.
6. Produces a repeatable consolidation process instead of manual copy/paste.

This is particularly relevant to recurring supplier-spend or operational-reporting workflows.

### Module 3 — `Module_RefreshDashboard.bas`

**Macro:** `RefreshCommercialModel`

Purpose:

1. Recalculates the workbook.
2. Refreshes workbook connections/queries where present.
3. Refreshes PivotTables where supported.
4. Records a refresh timestamp.

This provides a simple operational control for recurring reporting.

### Module 4 — `Module_ExportReport.bas`

**Macro:** `ExportDashboardToPDF`

Purpose:

1. Uses the `Dashboard` worksheet as the reporting output.
2. Opens a Save As dialog.
3. Exports the dashboard to PDF.
4. Provides a timestamped management-reporting output when required.

> VBA cannot be executed by Excel Online. Use **Microsoft Excel Desktop for Windows**.

---

## 9. Installing and running the VBA code

### Prerequisites

- Microsoft Excel Desktop.
- Windows is recommended for the supplied VBA workflow.
- Macros enabled for this trusted local project.
- A working copy of the workbook.

### Step 1 — Save as XLSM

Open the workbook in Excel Desktop and select:

**File → Save As → Excel Macro-Enabled Workbook (`*.xlsm`)**

Never test unfamiliar VBA code against your only copy of a workbook.

### Step 2 — Open the VBA editor

Press:

**Alt + F11**

Then select:

**File → Import File...**

Import each `.bas` file.

### Step 3 — Run cleansing

Run:

```text
CleanseProcurementData
```

Review the resulting data-quality flags.

### Step 4 — Test CSV consolidation

Place the CSV files you want to consolidate in a test folder, then run:

```text
ConsolidateCSVFiles
```

Check the generated consolidated worksheet before using the macro on production data.

### Step 5 — Refresh

Run:

```text
RefreshCommercialModel
```

Check that formulas, PivotTables and report outputs have updated.

### Step 6 — Export the report

Run:

```text
ExportDashboardToPDF
```

The macro saves the dashboard as a PDF for distribution or management review.

---

## 10. Power Query / M

`PowerQuery_Clean_Procurement.m` demonstrates a simple Power Query pattern:

1. Load procurement data.
2. Trim text fields.
3. Standardise data types.
4. Create a quality flag.
5. Produce a clean analytical table.

To use it in Excel:

1. Open **Data → Get Data → From Other Sources → Blank Query**.
2. Open **Advanced Editor**.
3. Paste the M code.
4. Change the source path if required.
5. Select **Done**.
6. Review the query preview.
7. Select **Close & Load**.

For a real project, replace the local file path with a controlled source such as SharePoint, OneDrive, a database or an approved enterprise data platform.

---

## 11. SQL checks

`SQL_Commercial_Checks.sql` is optional because the core project is Excel-based.

It demonstrates how the same analytical thinking could be transferred to a relational database.

Example tasks include:

- total procurement spend;
- spend by category;
- supplier concentration;
- duplicate invoice detection; and
- regional aggregation.

The SQL is intentionally simple enough to be understood during an interview while demonstrating that the candidate can move beyond spreadsheet-only analysis.

---

## 12. Recommended workflow for a fresh run

For someone using this project from scratch:

```text
1. Open the workbook
       ↓
2. Inspect the raw tables
       ↓
3. Run / review data-quality checks
       ↓
4. Inspect Cleaned_Procurement
       ↓
5. Validate Pivot_Analysis against Actual_Pivot
       ↓
6. Review category savings opportunities
       ↓
7. Open Financial_Model
       ↓
8. Change the scenario
       ↓
9. Observe EBITDA / margin sensitivity
       ↓
10. Review Dashboard
       ↓
11. Run VBA refresh
       ↓
12. Export the dashboard to PDF
```

This sequence mirrors a practical recurring commercial-reporting process.

---

## 13. Data-quality test cases

The dataset is intentionally imperfect so that the project demonstrates analytical judgement rather than only successful calculations.

Test for:

- blank supplier values;
- blank site IDs;
- blank spend values;
- blank invoice IDs;
- duplicate invoice IDs;
- inconsistent category casing;
- supplier/category mismatches;
- unusual spend values; and
- differences between source totals and analytical totals.

A professional analyst should **not silently overwrite questionable source data**. The preferred pattern is to retain the source value, create a quality flag, investigate the exception and document the treatment.

---

## 14. Financial-model controls

The model should be reviewed using the following controls:

### Revenue reconciliation

2026 revenue should equal:

```text
2025 Revenue × (1 + Selected Revenue Growth)
```

### Procurement forecast

```text
2025 Procurement × (1 − Selected Procurement Savings)
```

### Agency forecast

```text
2025 Agency × (1 + Selected Agency Rate Inflation)
```

### Fleet forecast

```text
2025 Fleet × (1 − Selected Fleet Efficiency)
```

### Payroll forecast

```text
2025 Payroll × (1 + Selected Payroll Inflation)
```

### EBITDA

```text
Forecast Revenue − Forecast Total Opex
```

### EBITDA margin

```text
EBITDA ÷ Revenue
```

These checks make the model auditable and make it easier to explain during an interview.

---

## 15. GDPR and responsible-data design

The project deliberately follows a **GDPR-first portfolio design**.

It does not use:

- names;
- addresses;
- telephone numbers;
- email addresses;
- NHS numbers;
- employee identifiers; or
- other direct personal identifiers.

Operational entities are represented using synthetic IDs such as `SiteID` and `VehicleID`.

In a real regulated-sector environment, the analyst would additionally need to follow the organisation's approved access controls, retention rules, lawful basis, data minimisation requirements, information-security procedures and policies for handling special-category or otherwise sensitive data.

---

## 16. Validation checklist

Before presenting this project, verify:

- [ ] Raw CSV totals reconcile to the workbook tables.
- [ ] Procurement categories are standardised.
- [ ] Duplicate invoices are flagged.
- [ ] Missing values are visible rather than silently deleted.
- [ ] `Pivot_Analysis` and `Actual_Pivot` are directionally consistent.
- [ ] Savings calculations reconcile to category spend × target %.
- [ ] Changing the scenario changes the forecast.
- [ ] EBITDA equals revenue less total forecast Opex.
- [ ] EBITDA margin equals EBITDA / revenue.
- [ ] Dashboard KPIs update after calculation/refresh.
- [ ] VBA is tested on a copy of the workbook.
- [ ] No personal/confidential information has been introduced.

---

## 17. Interview demonstration script

A concise five-minute walkthrough can follow this sequence:

**Minute 1 — Problem**

Explain that the analyst receives fragmented commercial data and needs reliable recurring management information.

**Minute 2 — Data quality**

Open `Procurement_Raw`, demonstrate an intentional quality issue, then show how `CleanseProcurementData` flags the exception.

**Minute 3 — Commercial insight**

Open `Pivot_Analysis` and show category spend, savings targets and potential savings. Validate the result against `Actual_Pivot`.

**Minute 4 — Financial modelling**

Open `Financial_Model` and change the scenario from Base to Conservative or Aggressive. Explain how the selected drivers flow through to Opex, EBITDA and margin.

**Minute 5 — Automation and reporting**

Run the refresh macro and export the `Dashboard` to PDF. Explain how this converts a manual monthly workflow into a repeatable process.

---

## 18. Portfolio positioning

This project is designed to demonstrate the following capabilities to employers:

| Capability | Evidence in project |
|---|---|
| Advanced Excel | Tables, formulas, lookups, SUMIFS, COUNTIFS, error handling, PivotTables, charts and validation |
| Financial modelling | Driver-based forecast, scenarios, savings model, EBITDA and margin |
| VBA | Cleansing, consolidation, refresh and PDF-report automation |
| Data quality | Missing-value, duplicate and standardisation controls |
| Commercial analysis | Procurement, supplier, regional, fleet, workforce and agency analysis |
| Reporting | Executive dashboard and PDF output |
| Power Query | M-based cleansing example |
| SQL | Commercial validation and aggregation examples |
| Communication | Dashboard, methodology, audit trail and technical documentation |
| GDPR awareness | Synthetic IDs, data minimisation and no personal data |

The project is intentionally domain-neutral enough to be presented for **commercial analyst, finance analyst, data analyst, procurement analyst, operations analyst and reporting analyst** applications.

---

## 19. Limitations and honest claims

This is a portfolio simulation, not a production enterprise reporting system. The datasets, assumptions and financial values are synthetic. The model does not reproduce any employer's actual contracts, accounting policies, procurement rules, KPI definitions or reporting architecture.

The project should therefore be presented as evidence of **transferable analytical capability**, not as evidence of access to confidential organisational data.

Likewise, Power BI is intentionally not the primary deliverable because this project is designed to showcase advanced Excel, VBA and financial modelling. A separate Power BI portfolio project can be used to demonstrate deeper DAX, semantic modelling and interactive BI capability.

---

## 20. License / portfolio use

Unless a separate licence is added, treat the supplied synthetic datasets and project materials as portfolio-learning material. Do not add confidential employer data, personal data or proprietary business information to this repository.

---

## 21. Quick start

If you only have five minutes:

1. Open the Excel workbook.
2. Go to **Dashboard**.
3. Review **Pivot_Analysis**.
4. Open **Financial_Model** and change the scenario.
5. Review **Savings_Tracker**.
6. Open **Formula_Audit**.
7. Open **VBA_Code**.
8. Import the `.bas` modules into an `.xlsm` copy.
9. Run the cleansing macro.
10. Export the dashboard to PDF.

That sequence demonstrates the complete project from raw data through commercial insight, financial modelling and automation.

