# Investment Portfolio Project 2026

## Project Overview

This is a course project for Investment (15% of final grade), consisting of two tasks:
- **Task 1**: Markowitz portfolio optimization across 10 Chinese industry indices (2015-2025)
- **Task 2**: Fama-French three-factor and five-factor model testing via two-pass regression

## Directory Structure

```
investment_hw/
├── README.md                  ← This file
├── task1.py                   ← Markowitz optimization script
├── task2.py                   ← Fama-French regression script
├── 数据文件/                    ← Raw data (from CSMAR and RESSET)
│   ├── 20150101-20171231.xlsx  ← Industry index data 2015-2017
│   ├── 20180101-20211231.xlsx  ← Industry index data 2018-2021
│   ├── 20220101-20251231.xlsx  ← Industry index data 2022-2025
│   ├── SHIBOR.xlsx             ← Overnight SHIBOR rates
│   ├── 日频三因子.xlsx           ← Fama-French 3 factors (daily)
│   └── 日频五因子.xlsx           ← Fama-French 5 factors (daily)
├── instructions/               ← Project requirements
│   ├── original_hw_requirements.txt
│   └── project_instructions.txt
└── output/                     ← Generated outputs (submit these)
    ├── task1_report.docx       ← [SUBMIT] Task 1 English Word report
    ├── task1_report_cn.docx    ← Task 1 Chinese Word report
    ├── efficient_frontier.png  ← Efficient frontier + CML plot (English)
    ├── efficient_frontier_cn.png
    ├── portfolio_weights.png   ← Tangency portfolio weights (English)
    ├── portfolio_weights_cn.png
    ├── task2_results.xlsx      ← [SUBMIT] Task 2 Excel file
    ├── task2_analysis.docx     ← [SUBMIT] Task 2 English analysis
    └── task2_analysis_cn.docx  ← Task 2 Chinese analysis
```

## What to Submit

| Deliverable | File | Format |
|---|---|---|
| Task 1 | `output/task1_report.docx` (or `_cn.docx`) | Word, ≤5 pages |
| Task 2 (data & tables) | `output/task2_results.xlsx` | Excel, 5 sheets |
| Task 2 (analysis) | `output/task2_analysis.docx` (or `_cn.docx`) | Word |

## Data Sources

| Data | Source | Period | Notes |
|---|---|---|---|
| 10 SW industry indices | CSMAR (国泰安) | 2015-01-05 to 2025-12-31 | 2,674 trading days, returns in % |
| SHIBOR overnight | RESSET (锐思) | Same period | Annualized %, converted via /36000 |
| FF3 factors | CSMAR | Same period | MarkettypeID=P9706, decimal form |
| FF5 factors | CSMAR | Same period | P9706, Portfolios=1 (2x3 sort) |

## 10 Industry Indices

| Code | Chinese Name | English Name |
|---|---|---|
| 801010 | 农林牧渔 | Agriculture |
| 801050 | 有色金属 | Non-Ferrous Metals |
| 801080 | 电子 | Electronics |
| 801110 | 家用电器 | Home Appliances |
| 801120 | 食品饮料 | Food & Beverage |
| 801150 | 医药生物 | Pharma & Bio |
| 801160 | 公用事业 | Utilities |
| 801180 | 房地产 | Real Estate |
| 801780 | 银行 | Banking |
| 801890 | 机械设备 | Machinery |

## Key Parameters

| Parameter | Value |
|---|---|
| Risk aversion (A) | 3 |
| Utility function | U = E(r) - ½Aσ² |
| Short-selling | Allowed (unconstrained weights) |
| Risk-free rate | SHIBOR overnight, r_daily = r_annual% / 36000 |
| Trading days/year | 252 |
| FF factor market type | P9706 (All A-shares, free-float weighted) |
| FF5 portfolio sort | Portfolios = 1 (2x3 sorting method) |

## Key Results Summary

### Task 1: Markowitz Portfolio Optimization

| Portfolio | Ann. Return | Ann. Std Dev | Sharpe Ratio |
|---|---|---|---|
| Global Min-Variance | 3.51% | 18.80% | — |
| Tangency (optimal risky) | 143.26% | 152.24% | 0.93 |
| Complete (y*=0.2041) | 30.31% | 31.07% | — |

The tangency portfolio has extreme weights (ranging from -361% to +405%) due to unconstrained short-selling. The investor allocates 20.4% to the tangency portfolio and 79.6% to the risk-free asset.

### Task 2: Fama-French Factor Analysis

| Model | Avg R² (first pass) | Significant alphas (5%) |
|---|---|---|
| FF3 | 0.648 | 食品饮料(+), 公用事业(-), 房地产(-) |
| FF5 | 0.655 | 房地产(-) |

Fama-MacBeth second-pass: No gamma coefficients are significant at the 5% level in either model. This is partly due to having only 10 cross-sectional observations (industry indices), which limits statistical power.

## How to Reproduce

```bash
# Requires conda environment with: pandas numpy scipy matplotlib openpyxl
# statsmodels xlsxwriter python-docx

PYTHON=/path/to/python

# Task 1
$PYTHON task1.py

# Task 2
$PYTHON task2.py
```

## Unit Conversion Notes (Important!)

The raw data files use different units — getting these wrong will produce incorrect results:

| Data | Raw Unit | Conversion to Decimal |
|---|---|---|
| Industry returns (Idxtrd08) | Percentage (1.5 = 1.5%) | Divide by 100 |
| SHIBOR (WghAvgPr) | Annualized % (3.64 = 3.64%/yr) | Divide by 36,000 |
| FF factors (RiskPremium1, SMB1, etc.) | Decimal (0.035 = 3.5%) | No conversion needed |
