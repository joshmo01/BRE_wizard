# Loan Portfolio Dashboard

A Claude Code skill that generates a complete, formula-driven Excel loan portfolio dashboard with synthetic data in one command.

## What It Produces

A single Excel file `loan_portfolio_dashboard.xlsx` with two sheets:

| Sheet | Contents |
|-------|----------|
| **Data** | 1,000 synthetic loan records with realistic distributions |
| **Dashboard** | 6 formula-driven summary sections (no charts) |

## Dashboard Sections

| # | Section | Description |
|---|---------|-------------|
| 1 | Executive KPIs | 10 KPIs: Total Loans, Disbursed, Outstanding, NPA Rate, Portfolio Yield, Monthly Profit, etc. |
| 2 | NPA by Segment | Retail / SME / Corporate / Agri — NPA count, %, outstanding |
| 3 | Portfolio Concentration | By product (Personal / Home / Auto / Business) — outstanding %, NPA count |
| 4 | Age Group vs DPD | 18-30 / 31-45 / 46-60 / 60+ — Avg DPD, delinquent & NPA counts |
| 5 | Interest Rate Bands | Low (<8%) / Medium (8-12%) / High (>=12%) — outstanding, profit, NPA |
| 6 | Profitability by Rating | AAA to C — avg interest rate, avg & total monthly profit, NPA count |

All values are live Excel formulas (SUMIF, COUNTIFS, AVERAGEIF, SUMPRODUCT) — change any data row and the dashboard recalculates automatically.

## Data Fields

| Column | Field | Description |
|--------|-------|-------------|
| A | Loan_ID | Unique ID (L-1001 to L-2000) |
| B | Segment | Retail / SME / Corporate / Agri |
| C | Product | Personal / Home / Auto / Business |
| D | Age_Group | 18-30 / 31-45 / 46-60 / 60+ |
| E | Credit_Rating | AAA, AA, A, BBB, BB, B, C |
| F | Disbursement_Amt | Original loan amount |
| G | Outstanding_Principal | Current balance |
| H | Interest_Rate | Annual rate (decimal) |
| I | Status | Active / Delinquent / NPA |
| J | DPD | Days Past Due |
| K | Monthly_Profit | Approx. monthly income |

## Usage

### Run the script directly

```bash
pip install openpyxl numpy pandas

python scripts/generate_dashboard.py
# or with a custom output folder:
python scripts/generate_dashboard.py --output /path/to/folder
```

### Use as a Claude Code skill

Place the folder at `~/.claude/skills/loan-portfolio-dashboard/`, then in any Claude Code session say:

```
build me a loan portfolio dashboard
```

Claude will run the script and save the Excel file automatically.

### Using your own data

Replace the data-generation block in `generate_dashboard.py` with:

```python
import pandas as pd
df = pd.read_csv("your_file.csv")
```

Column names and order must match exactly (case-sensitive).

## Dependencies

```
openpyxl
numpy
pandas
```

## File Structure

```
loan-portfolio-dashboard/
├── README.md
├── SKILL.md                        <- Claude Code skill definition
└── scripts/
    └── generate_dashboard.py       <- Self-contained dashboard generator
```
