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
| Status | Converted / Not Converted (driven by CIBIL + FOIR logic) |

### Summary Sheet Sections

1. Executive KPIs (10 values: conversion rate, avg CIBIL, avg income, total loan ask, avg FOIR, etc.)
2. Conversion by Loan Product
3. Conversion by Employment Type
4. Conversion by City Tier
5. CIBIL Band Analysis (600–649 / 650–699 / 700–749 / 750–799 / 800–900)
6. Lead Source Analysis

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
