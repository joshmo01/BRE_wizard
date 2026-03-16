---
name: loan-portfolio-dashboard
description: Build a loan portfolio Excel dashboard with synthetic data and formula-driven summary tables. Use this skill whenever the user asks to create, build, or generate a loan portfolio dashboard, loan book report, banking portfolio summary, or NBFC portfolio analysis in Excel. Also trigger when the user wants to analyze loan data across segments, products, credit ratings, DPD buckets, or NPA classifications — even if they don't use the word "dashboard". Always use this skill for any request involving loan portfolio reporting, credit risk summaries, or portfolio concentration analysis.
---

# Loan Portfolio Dashboard Skill

Generates a complete, formula-driven Excel loan portfolio dashboard with synthetic data in one command.

## Output

A single file: `loan_portfolio_dashboard.xlsx` with two sheets:

| Sheet | Contents |
|-------|----------|
| **Data** | 1,000 synthetic loan records with realistic distributions |
| **Dashboard** | 6 formula-driven summary sections, no charts |

## How to Run

Execute the bundled script from the current working directory:

```bash
python "C:/Users/joshm/.claude/skills/loan-portfolio-dashboard/scripts/generate_dashboard.py"
```

If the user wants to save the output to a specific folder, pass the path:

```bash
python "C:/Users/joshm/.claude/skills/loan-portfolio-dashboard/scripts/generate_dashboard.py" --output "C:/path/to/output"
```

## Data Sheet Columns

| Col | Field | Description |
|-----|-------|-------------|
| A | Loan_ID | L-1001 to L-2000 |
| B | Segment | Retail / SME / Corporate / Agri |
| C | Product | Personal / Home / Auto / Business |
| D | Age_Group | 18-30 / 31-45 / 46-60 / 60+ |
| E | Credit_Rating | AAA → C (7 grades) |
| F | Disbursement_Amt | Original loan amount (₹) |
| G | Outstanding_Principal | Current balance (₹) |
| H | Interest_Rate | Annual rate (decimal, e.g. 0.09) |
| I | Status | Active / Delinquent / NPA |
| J | DPD | Days Past Due |
| K | Monthly_Profit | Approx. monthly income (₹) |

## Dashboard Sections

### 1 — Executive KPIs
10 KPIs in a 2-column grid:
- Total Loans, Total Disbursed, Total Outstanding
- NPA Count, NPA Rate (%), Delinquent Count
- Portfolio Yield (weighted avg interest rate)
- Total Monthly Profit, Avg Loan Size, Active Loans

### 2 — NPA by Segment
Rows: Retail, SME, Corporate, Agri
Columns: Total Loans | NPA Count | NPA % | Outstanding (NPA) | Delinquent Count

### 3 — Portfolio Concentration by Product
Rows: Personal, Home, Auto, Business
Columns: Count | Outstanding (₹) | % of Portfolio | NPA Count | Avg Interest Rate

### 4 — Age Group vs DPD Analysis
Rows: 18-30, 31-45, 46-60, 60+
Columns: Count | Avg DPD | Delinquent Count | NPA Count | NPA %

### 5 — Interest Rate Band Analysis
Rows: Low (< 8%) | Medium (8–12%) | High (≥ 12%)
Columns: Count | Avg Rate | Total Outstanding | Total Monthly Profit | NPA Count

### 6 — Profitability by Credit Rating
Rows: AAA, AA, A, BBB, BB, B, C
Columns: Count | Avg Interest Rate | Avg Monthly Profit | Total Monthly Profit | NPA Count

## Script 2 — Loan Applications (`generate_loan_applications.py`)

Generates 3,000 synthetic Indian loan application records with a formula-driven Summary sheet.

```bash
python "C:/Users/joshm/.claude/skills/loan-portfolio-dashboard/scripts/generate_loan_applications.py"
# custom record count or output folder:
python generate_loan_applications.py --records 5000 --output "C:/path/to/folder"
```

Output file: `loan_applications.xlsx`

### Application Fields

| Field | Description |
|-------|-------------|
| App_ID | APP-0001 onwards |
| Age | 21–65 |
| Gender | Male / Female |
| City / City_Tier / State | 35 Indian cities across Tier 1 / 2 / 3 |
| Employment_Type | Salaried / Self-Employed / Business Owner |
| Employer_Category | Government / PSU / Private / MNC (Salaried only) |
| Monthly_Income | Based on employment type |
| Existing_EMI | 0–40% of income |
| CIBIL_Score | 600–900 |
| Loan_Product | Personal / Home / Auto / Business |
| Loan_Amount_Requested | Product-specific ranges |
| Loan_Tenure_Months | Product-specific options |
| Lead_Source | Online / Branch / DSA / Referral / Walk-in |
| Application_Date | Last 12 months |
| FOIR | Existing EMI / Monthly Income |

### Summary Sheet Sections

1. Executive KPIs (total applications, avg CIBIL, avg income, total loan ask, avg FOIR, etc.)
2. Applications by Loan Product
3. Applications by Employment Type
4. Applications by City Tier
5. CIBIL Band Analysis (600–649 / 650–699 / 700–749 / 750–799 / 800–900)
6. Lead Source Analysis

---

## Script 3 — Approved Loans (`generate_approved_loans.py`)

Reads `loan_applications.xlsx`, applies BRE eligibility rules, and enriches approved records.

```bash
python "C:/Users/joshm/.claude/skills/loan-portfolio-dashboard/scripts/generate_approved_loans.py"
# custom input/output:
python generate_approved_loans.py --input "C:/path/loan_applications.xlsx" --output "C:/path/to/folder"
```

Output file: `approved_loans.xlsx` (~137 loans, ~4.6% approval rate from 3,000 applications)

### BRE Eligibility Rules Applied

| Rule | Condition |
|------|-----------|
| Loan Type | Personal Loan only |
| CIBIL Score | > 750 |
| FOIR (Salaried) | < 20% |
| FOIR (Self-Employed / Business Owner) | < 15% |
| Loan Amount | < Rs 15,00,000 |
| City Tier | Tier 1 or Tier 2 only |
| Age | 25–50 years |

### Enriched Fields Added

Loan_ID, Approval_Date, Disbursement_Date, Sanctioned_Amount (90–100% of requested),
Interest_Rate (risk-based 9.5%–14.5% by CIBIL), EMI (reducing balance), Processing_Fee (1–2%)

### Summary Sheet Sections

1. Executive KPIs (approved count, total sanctioned, avg interest rate, avg EMI, avg CIBIL, etc.)
2. By Employment Type
3. By City Tier
4. CIBIL Band Analysis (751–900)
5. Interest Rate Bands
6. Lead Source Analysis

---

## Script 4 — Repayment Schedule (`generate_repayment_schedule.py`)

Reads `approved_loans.xlsx`, simulates monthly repayment for all loans applying 3 BRE rule sets.

```bash
python "C:/Users/joshm/.claude/skills/loan-portfolio-dashboard/scripts/generate_repayment_schedule.py"
# custom input/output:
python generate_repayment_schedule.py --input "C:/path/approved_loans.xlsx" --output "C:/path/to/folder"
```

Output file: `loan_repayment_schedule.xlsx`

### BRE Rule Sets Applied

| Rule Set | Logic |
|----------|-------|
| Penal Interest | 2% p.a. calculated daily on outstanding principal when DPD > 0 |
| Default Penalty | Rs 1,000 flat fee when DPD > 90 |
| Prepayment Charges | 2% of outstanding principal (full prepayment) or prepaid amount (partial) |

### Payment Profiles Simulated

| Profile | % of Loans | Behaviour |
|---------|-----------|-----------|
| Pristine | 45% | Always on time, DPD = 0 |
| Occasional Delay | 25% | 70% on-time, 20% DPD 15, 10% DPD 45 |
| Chronic Delay | 10% | Mix of DPD 0/20/45/75/95 |
| Partial Prepayment | 10% | On-time with one mid-tenure partial prepayment |
| Full Prepayment | 10% | On-time with full closure at random month |

### Output Columns (21)

Schedule_ID, Loan_ID, EMI_Number, Due_Date, Payment_Date, DPD, DPD_Bucket,
Opening_Balance, Scheduled_EMI, Principal_Component, Interest_Component,
Penal_Interest, Prepayment_Amount, Prepayment_Charge, Prepayment_Type,
Default_Penalty, Loan_Status, Closing_Balance, Payment_Status, Total_Amount_Due, Total_Amount_Paid

### DPD Buckets

Current | DPD 1-30 | DPD 31-60 | DPD 61-90 | DPD 90+

### Summary Sheet Sections

1. Executive KPIs (total EMIs, total collected, total penal interest, total prepayments, NPA count, etc.)
2. Payment Status Distribution
3. DPD Bucket Analysis
4. Loan Status Summary
5. Prepayment Analysis

---

## Using the User's Own Data

If the user provides a CSV file instead of synthetic data, replace the data-generation block with:

```python
import pandas as pd
df = pd.read_csv("their_file.csv")
```

The column names must match exactly (case-sensitive). The dashboard formulas reference the Data sheet by column letter — column order must be preserved.

## Dependencies

```bash
pip install openpyxl numpy pandas
```
