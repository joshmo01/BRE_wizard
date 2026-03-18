---
name: loan-portfolio-dashboard
description: Build a loan portfolio Excel dashboard with synthetic data and formula-driven summary tables. Use this skill whenever the user asks to create, build, or generate a loan portfolio dashboard, loan book report, banking portfolio summary, or NBFC portfolio analysis in Excel. Also trigger when the user wants to analyze loan data across segments, products, credit ratings, DPD buckets, or NPA classifications — even if they don't use the word "dashboard". Always use this skill for any request involving loan portfolio reporting, credit risk summaries, or portfolio concentration analysis.
---

# Loan Portfolio Dashboard Skill

Generates a complete, formula-driven Excel loan portfolio dashboard with synthetic data in one command.

## Output

A single file: `loan_portfolio_dashboard.xlsx` with **6 sheets** — generated as the **final step** after all pipeline scripts have run:

| Sheet | Contents |
|-------|----------|
| **Data** | 1,000 synthetic loan records with realistic distributions |
| **Dashboard** | 6 formula-driven analytical sections + Executive KPIs with Actual vs Target RAG status |
| **Portfolio Setup** | Initial assumptions from Step 0 interview (product mix, employment mix, city tiers, loan amount ranges, record count) |
| **Pricing Summary** | Spread logic table + avg rate pivot by CIBIL band × Employment Type (from `Loan Pricing.xlsx`) |
| **Eligibility Rules** | Full BRE rule table R001–R008 with values as configured (from `Loan Eligibility Rules.xlsx`) |
| **Repayment Config** | All 6 sections of repayment assumptions — profiles, DPD distributions, prepayment & penal parameters (from `Repayment Assumptions.xlsx`) |

## How to Run

> **MANDATORY STOP — TARGET KPI INTERVIEW**
> Before running this script, you MUST ask the user for their target values for each of the
> 10 Executive KPIs. Do NOT run the script until the user has confirmed their targets.
> See "Target KPI Interview" section below for the exact questions to ask.

Once targets are confirmed, pass them as CLI args:

```bash
python "C:/Users/joshm/.claude/skills/loan-portfolio-dashboard/scripts/generate_dashboard.py" \
  --output "C:/path/to/output" \
  --tgt-total-loans       1000 \
  --tgt-total-disbursed   50000000 \
  --tgt-total-outstanding 30000000 \
  --tgt-npa-count         50 \
  --tgt-npa-rate          0.05 \
  --tgt-delinquent-count  80 \
  --tgt-portfolio-yield   0.10 \
  --tgt-monthly-profit    250000 \
  --tgt-avg-loan-size     30000 \
  --tgt-active-loans      870
```

Omit any `--tgt-*` arg to default that KPI's target to the dashboard's actual value (will always show ▲ On Track in green).

---

## Target KPI Interview

Ask the user the following before generating the dashboard. Present all questions at once:

> I need your **target values** for the 10 Executive KPIs so the dashboard can show
> Actual vs Target with RAG (green/red) status.
>
> Please provide targets for any or all of the following (skip any you don't have a target for):
>
> | # | KPI | Direction | Example |
> |---|-----|-----------|---------|
> | 1 | Total Loans (count) | Higher is better | e.g. 1000 |
> | 2 | Total Disbursed (Rs) | Higher is better | e.g. 5,00,00,000 |
> | 3 | Total Outstanding (Rs) | Higher is better | e.g. 3,00,00,000 |
> | 4 | NPA Count | **Lower is better** | e.g. 50 |
> | 5 | NPA Rate (%) | **Lower is better** | e.g. 5% |
> | 6 | Delinquent Count | **Lower is better** | e.g. 80 |
> | 7 | Portfolio Yield (%) | Higher is better | e.g. 10% |
> | 8 | Total Monthly Profit (Rs) | Higher is better | e.g. 2,50,000 |
> | 9 | Avg Loan Size (Rs) | Higher is better | e.g. 30,000 |
> | 10 | Active Loans (count) | Higher is better | e.g. 870 |

After collecting answers, show a confirmation table and wait for explicit user confirmation before running.

### Mapping targets to CLI args

| KPI | CLI Arg | Notes |
|-----|---------|-------|
| Total Loans | `--tgt-total-loans` | Integer count |
| Total Disbursed | `--tgt-total-disbursed` | Rs amount |
| Total Outstanding | `--tgt-total-outstanding` | Rs amount |
| NPA Count | `--tgt-npa-count` | Integer count |
| NPA Rate | `--tgt-npa-rate` | Decimal: 5% → 0.05 |
| Delinquent Count | `--tgt-delinquent-count` | Integer count |
| Portfolio Yield | `--tgt-portfolio-yield` | Decimal: 10% → 0.10 |
| Total Monthly Profit | `--tgt-monthly-profit` | Rs amount |
| Avg Loan Size | `--tgt-avg-loan-size` | Rs amount |
| Active Loans | `--tgt-active-loans` | Integer count |

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

## Script 2b — Loan Pricing Table (`build_loan_pricing.py`)

Generates a complete pricing table with all combinations of risk dimensions and suggested spreads.

> **MANDATORY PAUSE — DO NOT PROCEED PAST THIS STEP WITHOUT EXPLICIT USER CONFIRMATION.**
> After running this script you MUST stop, tell the user the file is ready for review, and wait
> for them to confirm before running any subsequent script. No exceptions.

```bash
python "C:/Users/joshm/.claude/skills/loan-portfolio-dashboard/scripts/build_loan_pricing.py"
# custom output folder:
python build_loan_pricing.py --output "C:/path/to/folder"
```

Output file: `Loan Pricing.xlsx` (two sheets)

### Pricing Dimensions

| Dimension | Values |
|-----------|--------|
| Employment Type + Category | Salaried (Govt/PSU/MNC/Private), Self-Employed, Business Owner |
| CIBIL Score Band | 751-775, 776-800, 801-825, 826-850, 851-900 |
| Age Group | 25-35, 36-45, 46-50 |
| Loan Amount | Up to 3L, 3L-7L, 7L-15L |
| Tenure | 12-24 months, 25-36 months, 37-60 months |

**Total combinations: 810 rows**

### Spread Logic (additive over Base Rate of 9.00%)

| Component | Range |
|-----------|-------|
| CIBIL spread | 0.50% (851-900) → 4.00% (751-775) |
| Employment spread | 0.00% (Govt/PSU) → 1.00% (Business Owner) |
| Loan amount spread | 0.00% (7L-15L) → +0.50% (Up to 3L) |
| Tenure spread | -0.25% (short) → +0.25% (long) |
| Age spread | 0.00% (25-45) → +0.25% (46-50) |

**Final Rate range: 9.25% — 15.00%** (colour-coded green/yellow/red in Excel)

Sheet 2 (`Rate Summary`) shows a pivot of avg rate by CIBIL band × Employment Category.

> **Required message to user after running this script:**
> "Loan Pricing table saved to `<output_path>/Loan Pricing.xlsx`. Please open it, review/adjust
> the `Final Rate (%)` column (and spread columns if needed) to match your institution's pricing
> policy, then save and close the file. Let me know when you're done and I'll continue."
>
> **Then STOP. Do not run Step 3 or any further script until the user explicitly confirms.**

---

## Script 2c — Loan Eligibility Rules (`build_loan_eligibility_rules.py`)

Generates an editable Excel BRE configuration file with all loan eligibility rules.

> **MANDATORY PAUSE — DO NOT PROCEED PAST THIS STEP WITHOUT EXPLICIT USER CONFIRMATION.**
> After running this script you MUST stop, tell the user the file is ready for review, and wait
> for them to confirm before running `generate_approved_loans.py`. No exceptions.

```bash
python "C:/Users/joshm/.claude/skills/loan-portfolio-dashboard/scripts/build_loan_eligibility_rules.py"
# custom output folder:
python build_loan_eligibility_rules.py --output "C:/path/to/folder"
```

Output file: `Loan Eligibility Rules.xlsx` (two sheets)

### Rules Defined (R001–R008)

| Rule ID | Rule Name | Default Value | Applies To |
|---------|-----------|---------------|------------|
| R001 | Loan Product | Personal | All |
| R002 | CIBIL Score Minimum | > 750 | All |
| R003 | FOIR Cap – Salaried | < 20% | Salaried |
| R004 | FOIR Cap – Non-Salaried | < 15% | Self-Employed, Business Owner |
| R005 | Max Loan Amount (Rs) | < 15,00,000 | All |
| R006 | City Tier | Tier 1, Tier 2, Tier 3 | All |
| R007 | Minimum Age | >= 25 | All |
| R008 | Maximum Age | <= 50 | All |

### What the user can edit
- **Value column** (amber) — change any threshold or list
- **Enabled column** — set `No` to deactivate a rule entirely
- Operator and Field columns must not be changed

> **Required message to user after running this script:**
> "Loan Eligibility Rules saved to `<output_path>/Loan Eligibility Rules.xlsx`. Please open it,
> review/adjust the Value column (amber) for each rule to match your credit policy, set Enabled=No
> for any rule you want to skip, then save and close. Let me know when you're done."
>
> **Then STOP. Do not run Step 3 until the user explicitly confirms.**

---

## Script 3 — Approved Loans (`generate_approved_loans.py`)

Reads `loan_applications.xlsx`, applies BRE eligibility rules, and enriches approved records.

```bash
python "C:/Users/joshm/.claude/skills/loan-portfolio-dashboard/scripts/generate_approved_loans.py"
# custom input/pricing/rules/output:
python generate_approved_loans.py --input   "C:/path/loan_applications.xlsx" \
                                   --pricing "C:/path/Loan Pricing.xlsx" \
                                   --rules   "C:/path/Loan Eligibility Rules.xlsx" \
                                   --output  "C:/path/to/folder"
```

Output file: `approved_loans.xlsx`

**Eligibility rules are loaded from `Loan Eligibility Rules.xlsx`** (R001–R008).
If the file is missing, built-in defaults are used and a WARNING is printed.

**Interest rates are looked up from `Loan Pricing.xlsx`**.
If missing, falls back to a CIBIL-based formula (9.5%–14.5%).

### BRE Rules Applied (read from Loan Eligibility Rules.xlsx)

| Rule ID | Rule | Default |
|---------|------|---------|
| R001 | Loan Product | Personal only |
| R002 | CIBIL Score | > 750 |
| R003 | FOIR – Salaried | < 20% |
| R004 | FOIR – Non-Salaried | < 15% |
| R005 | Max Loan Amount | < Rs 15,00,000 |
| R006 | City Tier | Tier 1, Tier 2, Tier 3 |
| R007 | Min Age | >= 25 |
| R008 | Max Age | <= 50 |

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

## Script 6 — Lifecycle Dashboard (`generate_lifecycle_dashboard.py`)

Reads all three data files and builds a single unified Excel dashboard covering the full loan lifecycle.
**Run this as the final step after the repayment schedule has been generated.**

```bash
python "C:/Users/joshm/.claude/skills/loan-portfolio-dashboard/scripts/generate_lifecycle_dashboard.py"
# custom inputs/output:
python generate_lifecycle_dashboard.py --apps   "C:/path/loan_applications.xlsx" \
                                        --loans  "C:/path/approved_loans.xlsx" \
                                        --sched  "C:/path/loan_repayment_schedule.xlsx" \
                                        --output "C:/path/to/folder"
```

Output file: `loan_lifecycle_dashboard.xlsx` (4 sheets: Dashboard + 3 data sheets)

### Dashboard Sections

| # | Section | Source |
|---|---------|--------|
| 1 | Lifecycle KPIs | Cross-file — approval rate, collection efficiency, totals, portfolio yield |
| 2 | Applications by Lead Source | Applications sheet |
| 3 | Approved Loans by Employment Type | Approved_Loans sheet |
| 4 | Approved Loans by City Tier | Approved_Loans sheet |
| 5 | CIBIL Band Analysis (751-900) | Approved_Loans sheet |
| 6 | Repayment Performance by Payment Status | Repayment_Schedule sheet |

---

## Step 0 — Portfolio Interview

**Run this interview BEFORE generating any data.** Ask each question, wait for the answer, then
present a confirmation summary table. Only proceed to Step 1 once the user confirms.

### Questions to ask (one block, not one-by-one)

> I'll need a few details to tailor the synthetic portfolio to your requirements:
>
> **1. Loan Products** — What % mix would you like?
> (Defaults: Personal 40%, Home 25%, Auto 20%, Business 15%)
>
> **2. Customer Type (Employment)** — What % mix?
> (Defaults: Salaried 50%, Self-Employed 30%, Business Owner 20%)
>
> **3. City Tier** — What % mix across Tier 1 / Tier 2 / Tier 3?
> (Default: Tier 1 heavy — approx 57% / 32% / 11%)
>
> **4. Loan Amounts** — Any changes to the default ranges?
> | Product  | Default Range (Rs)         |
> |----------|----------------------------|
> | Personal | 50,000 – 15,00,000         |
> | Home     | 10,00,000 – 1,00,00,000    |
> | Auto     | 3,00,000 – 20,00,000       |
> | Business | 5,00,000 – 50,00,000       |
>
> **5. Record Count** — How many applications? (Default: 3,000)

### Confirmation summary to show before proceeding

After the user responds, display:

```
Portfolio Distribution Confirmed
─────────────────────────────────────────────
Products    : Personal X%  Home X%  Auto X%  Business X%
Employment  : Salaried X%  SEP X%  Business Owner X%
City Tiers  : Tier 1 X%  Tier 2 X%  Tier 3 X%
Loan Ranges : Personal Rs A–B | Home Rs C–D | Auto Rs E–F | Business Rs G–H
Records     : X,000 applications
─────────────────────────────────────────────
Shall I proceed with Step 1?
```

Wait for explicit user confirmation before running any script.

### Mapping answers to CLI args

```bash
python "C:/Users/joshm/.claude/skills/loan-portfolio-dashboard/scripts/generate_loan_applications.py" \
  --records       5000 \
  --product-weights "50,20,20,10" \
  --emp-weights     "60,25,15" \
  --tier-weights    "40,35,25" \
  --loan-amt-personal "50000,1500000" \
  --loan-amt-home     "1000000,10000000" \
  --loan-amt-auto     "300000,2000000" \
  --loan-amt-business "500000,5000000" \
  --output "C:/Users/joshm/OneDrive/Documents/BRE"
```

- `--product-weights` : comma-separated weights for Personal, Home, Auto, Business (need not sum to 100 — used as relative weights)
- `--emp-weights`     : Salaried, Self-Employed, Business Owner
- `--tier-weights`    : must sum to 100 — converted to per-city probabilities inside the script
- `--loan-amt-*`      : Min,Max in Rs for each product (no spaces around comma)
- Omit any arg to keep its default

---

## End-to-End Workflow

```
Step 0   [INTERVIEW]                        →  Confirm portfolio distribution (products, employment,
                                               city tier, loan amounts, record count)
         *** MANDATORY STOP — Wait for user to confirm before proceeding ***

Step 1   generate_loan_applications.py     →  loan_applications.xlsx            (parameterised by interview)

Step 2a  build_loan_pricing.py             →  Loan Pricing.xlsx                 (810 rate combinations)

  *** MANDATORY STOP — PRICING ***
  Tell the user the file path. Ask them to review/adjust Final Rate (%) and confirm.
  DO NOT continue until the user explicitly says they are done reviewing.

Step 2b  build_loan_eligibility_rules.py   →  Loan Eligibility Rules.xlsx       (R001–R008)

  *** MANDATORY STOP — ELIGIBILITY RULES ***
  Tell the user the file path. Ask them to review/adjust the Value column (amber)
  and confirm. DO NOT run Step 3 until the user explicitly says they are done reviewing.

Step 3   generate_approved_loans.py        →  approved_loans.xlsx               (rules + rates applied)

Step 3b  build_repayment_assumptions.py    →  Repayment Assumptions.xlsx        (6-section config)

  *** MANDATORY STOP — REPAYMENT ASSUMPTIONS ***
  Tell the user the file path. Ask them to review/adjust all 6 sections (amber cells)
  and confirm % columns sum to 100 where required.
  DO NOT run Step 4 until the user explicitly says they are done reviewing.

Step 4   generate_repayment_schedule.py    →  loan_repayment_schedule.xlsx      (assumptions applied)
Step 5   generate_repayment_dashboard.py   →  repayment_dashboard.xlsx          (repayment analysis)
Step 6   generate_lifecycle_dashboard.py   →  loan_lifecycle_dashboard.xlsx     (full lifecycle view)

Step 7   generate_dashboard.py            →  loan_portfolio_dashboard.xlsx     (FINAL COMPREHENSIVE OUTPUT)

  *** MANDATORY STOP — TARGET KPI INTERVIEW ***
  Before running Step 7, ask the user for Target KPI values (see Target KPI Interview section).
  DO NOT run the script until targets are confirmed.

  Pass all pipeline config files + interview answers + targets as CLI args:
    --pricing     "C:/path/Loan Pricing.xlsx"
    --rules       "C:/path/Loan Eligibility Rules.xlsx"
    --assumptions "C:/path/Repayment Assumptions.xlsx"
    --records N --product-weights "..." --emp-weights "..." --tier-weights "..."
    --tgt-npa-rate 0.02 --tgt-total-outstanding 79000000 ...

Step 8   generate_kpi_suggestions.py     →  KPI_Gap_Suggestions.xlsx          (GAP ANALYSIS + SUGGESTIONS)

  Run AFTER Step 7. Automatically reads Target vs Actual from the dashboard.
  No interview needed — suggestions are auto-generated based on KPI gaps.

  Pass pipeline file paths as CLI args:
    --dashboard   "C:/path/loan_portfolio_dashboard.xlsx"
    --pricing     "C:/path/Loan Pricing.xlsx"
    --rules       "C:/path/Loan Eligibility Rules.xlsx"
    --approved    "C:/path/approved_loans.xlsx"
    --applications "C:/path/loan_applications.xlsx"
    --output      "C:/path/to/folder"
```

---

## Script 8 — KPI Gap Analysis & Suggestions (`generate_kpi_suggestions.py`)

Compares Target KPIs vs Actual from the dashboard and generates actionable suggestions across three levers:
1. **Funnel / Application Intake** — expand customer segments, add products, widen eligibility
2. **Loan Pricing Matrix** — adjust spreads, base rate, or band-level pricing
3. **Rule Engine / Eligibility + Collection Policy** — tighten/relax BRE rules, strengthen collections

Run this as the **final step** after `generate_dashboard.py` has produced the dashboard with Target KPIs.

```bash
python "C:/Users/joshm/.claude/skills/loan-portfolio-dashboard/scripts/generate_kpi_suggestions.py"
# custom paths:
python generate_kpi_suggestions.py --dashboard "C:/path/loan_portfolio_dashboard.xlsx" \
                                    --pricing   "C:/path/Loan Pricing.xlsx" \
                                    --rules     "C:/path/Loan Eligibility Rules.xlsx" \
                                    --approved  "C:/path/approved_loans.xlsx" \
                                    --applications "C:/path/loan_applications.xlsx" \
                                    --output    "C:/path/to/folder"
```

Output file: `KPI_Gap_Suggestions.xlsx` (3 sheets)

### Sheet 1 — Gap Summary
All 10 KPIs with Actual vs Target, Gap %, and On Track / Off Target status.

### Sheet 2 — Suggestions (grouped by lever)

| Lever | Example Scenarios |
|-------|-------------------|
| **Funnel / Application Intake** | Disbursement below target → Expand to SENP segment, add Home/Auto products, widen age range, add digital lead channels |
| **Loan Pricing Matrix** | Yield below target → Increase CIBIL spreads for lower bands, raise tenure/amount spreads, adjust employment-category pricing |
| **Rule Engine / Eligibility** | NPA above target → Tighten CIBIL floor for high-NPA segments, cap loan amounts for risky ratings, exclude worst-performing segments |
| **Collection Policy** | Delinquency above target → Automated early-stage reminders, risk-based collection intensity, field visits at DPD 30 for high-risk |

Each suggestion includes: KPI affected, gap description, priority (High/Medium/Low), current value, suggested change, and an **Action column (amber)** for the user to fill in their decision.

### Sheet 3 — Data Analysis
Segment-level drill-down supporting the suggestions:
- NPA by Segment (Retail, SME, Corporate, Agri)
- NPA by Credit Rating (AAA → C)
- NPA by Product (Personal, Home, Auto, Business)

All tables are RAG-coded: red if NPA > 10%, yellow if NPA > 5%.

### Suggestion Logic

| KPI Gap | Lever | Suggestion Type |
|---------|-------|-----------------|
| Disbursement / Total Loans < Target | Funnel | Add SENP segment, expand products, add lead channels |
| Disbursement / Total Loans < Target | Rules | Relax CIBIL floor (750→700), increase FOIR cap, widen age range |
| Portfolio Yield < Target | Pricing | Increase CIBIL spreads for lower bands, add tenure/amount spreads |
| Portfolio Yield < Target | Pricing | Increase employment spread for Private sector |
| NPA Rate > Target | Rules | Tighten CIBIL for worst segment, exclude worst rating, cap loan amounts |
| NPA / Delinquency > Target | Collection | Early-stage reminders, risk-based intensity, field visits |
| NPA > Target | Pricing | Add segment-level NPA spread (+0.50% where NPA > 5%) |
| Monthly Profit < Target | Pricing | Review NIM, reprice book |
| Monthly Profit < Target | Funnel | Grow outstanding balance |
| Avg Loan Size < Target | Rules | Raise max loan cap, add higher-ticket products |

> **No mandatory pause needed** — this step is informational. Present the output path and tell
> the user to review the Suggestions sheet and fill in the Action column.

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
