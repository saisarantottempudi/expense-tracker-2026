<div align="center">

# Multi-Currency Expense Tracker 2026

### A fully automated Excel-based personal finance tracker built with Python

[![Python](https://img.shields.io/badge/Python-3.8+-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![Excel](https://img.shields.io/badge/Excel-Compatible-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://www.microsoft.com/excel)
[![License](https://img.shields.io/badge/License-MIT-yellow?style=for-the-badge)](LICENSE)
[![Formulas](https://img.shields.io/badge/Formulas-15%2C300+-orange?style=for-the-badge)]()
[![Sheets](https://img.shields.io/badge/Sheets-18-blue?style=for-the-badge)]()
[![Google Sheets](https://img.shields.io/badge/Google%20Sheets-Compatible-34A853?style=for-the-badge&logo=google-sheets&logoColor=white)](https://sheets.google.com)

Manage your finances across GBP and INR with 14 accounts, 14 categories, and zero manual calculations.

[Features](#features) | [Quick Start](#quick-start) | [Sheets Overview](#sheets-overview) | [How It Works](#how-it-works)

</div>

---

## The Problem

Managing finances across two countries means juggling multiple bank accounts, credit cards, and currencies. Existing apps either do not support multi-currency properly, lock features behind subscriptions, or do not give you full control over your data.

## The Solution

A single Python script generates a fully formatted Excel workbook with 15,300+ interlinked formulas across 18 sheets. Enter data in 3 input sheets and everything else auto-calculates.

---

## Features

### Core Tracking
- Multi-Currency: GBP and INR with auto-detection based on payment method
- 14 Payment Accounts: Banks, credit cards, digital wallets, travel cards
- 14 Expense Categories: Fully customisable
- Income Tracking: Multiple employers with pay cycle date ranges

### Smart Features
- Splitwise Integration: Track shared expenses, auto-split balances
- Credit Card Payments: Transfers restore available credit limits
- Oyster Top-ups: Cross-account tracking between bank and travel card
- Recurring Expense Flag: Mark subscriptions and bills for forecasting

### Auto-Generated Views
- Financial Dashboard: Single-page snapshot of your entire financial position
- Monthly Overview: Budget limits per category with red/green conditional formatting
- Daily Cashflow: Running balance for every day of the year with trend line chart
- Weekly Itemised: Auto-populated Monday to Sunday grid from your entries

### Planning and Analysis
- Debt Tracker: Loans, EMIs, Buy Now Pay Later, repayment schedules
- Savings Goals: Set targets with deadlines, track progress percentage
- Recurring Forecast: Project fixed monthly costs across all 12 months
- Year-on-Year Comparison: Compare spending patterns across years
- Currency Conversion Log: Track exchange rates and fees for international transfers

---

## Supported Accounts

### GBP Accounts
| No | Account | Type |
|----|---------|------|
| 1 | Lloyds Bank | Bank Account |
| 2 | HSBC Bank | Bank Account |
| 3 | Halifax Bank | Bank Account |
| 4 | Barclays Bank | Bank Account |
| 5 | Capitol One Credit | Credit Card |
| 6 | Ocean Credit | Credit Card |
| 7 | Zable Credit Card | Credit Card |
| 8 | Oyster | Travel Card |
| 9 | Splitwise | Expense Sharing |
| 10 | Klarna Pay | Pay Later |
| 11 | Klarna Pay Later | Credit |

### INR Accounts
| No | Account | Type |
|----|---------|------|
| 12 | ICICI Bank | Bank Account |
| 13 | ICICI Amazon Pay VISA Credit Card | Credit Card |
| 14 | ICICI Coral Rupay Credit Card | Credit Card |

New accounts can be added via the Accounts sheet. The system is designed to be extensible.

---

## Expense Categories

Rent, Transport, Travel, Shopping, Electronics, Bills, Entertainments, Food And Drinks, Groceries, Saving, Donation, Investments, Subscriptions, Conversions

---

## Quick Start

### Prerequisites

- Python 3.8 or higher
- openpyxl library

### Installation

```bash
git clone https://github.com/saisarantottempudi/expense-tracker-2026.git
cd expense-tracker-2026
pip install -r requirements.txt
python build_tracker.py
```

### First Time Setup

1. Open Expense_Tracker_2026.xlsx in Excel or Google Sheets
2. Settings: Update the exchange rate and year
3. Accounts: Verify your accounts and enter card last 4 digits
4. Balances: Enter opening balances or credit limits for each account
5. Monthly Overview: Set your budget limits per category
6. Cashflow: Enter starting balance on row 1
7. Start logging in Expenses and Income. Everything else updates automatically.

---

## Sheets Overview

The workbook contains 18 sheets organised into four groups:

### Config Sheets
- **Settings**: Global configuration including year, exchange rates, and budget alert thresholds
- **Accounts** : Master list of all bank accounts and cards with type, currency, and status

### Input Sheets (where you enter data)
- **Expenses**: Log all spending with date, category, price, transaction gateway, Splitwise flag, recurring flag, and notes
- **Income**: Log all earnings with company, amount, bank account credited, and pay cycle dates
- **Transfers**: Record credit card payments, Oyster top-ups, and Splitwise settlements

### Auto-Calculated Sheets
- **Dashboard**: Single-page financial snapshot showing this month, YTD, and account balances
- **Balances**: Real-time balance for all 14 accounts with credits, debits, and current balance
- **Monthly Overview GBP**: 12 months x 14 categories with budget limits, conditional formatting, bar chart, and pie chart
- **Monthly Overview INR**: Same as above for INR currency
- **Cashflow GBP**: Daily running balance for every day of 2026 with income, expenses, and line chart
- **Cashflow INR**: Same as above for INR currency
- **Itemised GBP**: Weekly Monday-to-Sunday grid auto-populated from Expenses and Income sheets
- **Itemised INR**: Same as above for INR currency

### Planning Sheets
- **Debt Tracker**: Track loans, EMIs, and BNPL with interest rates, monthly payments, and remaining balance
- **Savings Goals**: Set financial targets with deadlines and track progress with on-track indicators
- **Recurring Forecast**: Project fixed monthly expenses across all 12 months with annual totals
- **Year-on-Year**: Compare current year vs previous year spending per category with change percentages
- **Currency Conversions**: Log GBP-INR transfers with exchange rates, fees, and auto-calculated received amounts

---

## How It Works

### Data Flow

All data enters through three input sheets: Expenses, Income, and Transfers. These feed into the Balances sheet which tracks all 14 accounts. From there, the data flows into Monthly Overviews (budget tracking with charts), Cashflow sheets (daily running balance with trend charts), and Itemised sheets (weekly grid breakdown). The Dashboard pulls from everything to give you a single-page snapshot.

```
Settings (Config)
       |
       v
Expenses + Income + Transfers (Input)
       |
       v
Balances (All 14 accounts)
       |
       +---> Monthly Overview GBP/INR (Budget + Charts)
       +---> Cashflow GBP/INR (Daily Balance + Line Chart)
       +---> Itemised GBP/INR (Weekly Grid)
       +---> Dashboard (Snapshot)
```

### Splitwise Logic

When you buy something and split it with a friend:

1. Enter the expense normally in the Expenses sheet
2. Set Splitwise column to Yes
3. Enter Your Share amount (the rest is what your friend owes)
4. The full price debits your paying account (e.g. Lloyds Bank)
5. Only Your Share counts toward your Monthly Overview spending
6. The difference (Price minus Your Share) credits your Splitwise balance

When your friend pays you back, use the Transfers sheet: set To as your bank account and From as Splitwise. When you pay your friend what you owe, set To as Splitwise and From as your bank account.

### Credit Card Payments

Use the Transfers sheet to pay credit cards from bank accounts. Set To as the credit card and From as the bank account. This debits the bank balance and restores the credit card available limit. It is not counted as an expense, keeping your spending totals accurate.

### Oyster Top-ups

Same as credit card payments. Use Transfers with To as Oyster and From as the bank card you used to top up.

---

## Colour Coding Guide

- Yellow cells: Input cells where you enter your data
- Blue text: Manual inputs like opening balances and credit limits
- Green text: Auto-calculated formulas pulling from other sheets
- Bold text: Calculated totals and summaries
- Green background: Under budget or positive amount
- Red background: Over budget or negative amount

---

## Customisation

### Adding New Accounts
1. Add the account name to the DEFAULT_GATEWAYS list in build_tracker.py
2. If it is an INR account, also add it to INR_GATEWAYS_LIST
3. If it is a credit card, also add it to CREDIT_CARDS
4. Re-run python build_tracker.py

### Adding New Categories
Edit the CATEGORIES list in build_tracker.py and re-run the script.

### Changing the Year
Update YEAR = 2026 at the top of build_tracker.py to your desired year.

### Adding New Income Sources
Edit the COMPANIES list in build_tracker.py and re-run the script.

---

## Project Structure

```
expense-tracker-2026/
|-- build_tracker.py              Main Python script that generates the Excel file
|-- Expense_Tracker_2026.xlsx     Generated output ready to use
|-- README.md                     Documentation (this file)
|-- LICENSE                       MIT License
|-- requirements.txt              Python dependencies
|-- .gitignore                    Git ignore rules
```

---

## Roadmap

- Google Sheets native integration using Apps Script
- Web dashboard with real-time charts
- Bank statement CSV import for auto-populating expenses
- Automated currency rate fetching via API
- Multi-year support in a single workbook
- PDF monthly report generation
- Telegram or WhatsApp bot for quick expense entry on the go

---

## Contributing

Contributions are welcome. Here is how:

1. Fork the repository
2. Create your feature branch: git checkout -b feature/amazing-feature
3. Commit your changes: git commit -m "feat: add amazing feature"
4. Push to the branch: git push origin feature/amazing-feature
5. Open a Pull Request

---

## License

This project is licensed under the MIT License. See the LICENSE file for details.

---

## Acknowledgements

- Built with openpyxl for Excel generation
- Inspired by the need for better multi-currency personal finance tools
- Designed for the Indian diaspora managing finances across borders

---

<div align="center">

Star this repo if you find it useful!

Built with Python and openpyxl

[Report Bug](https://github.com/saisarantottempudi/expense-tracker-2026/issues) | [Request Feature](https://github.com/saisarantottempudi/expense-tracker-2026/issues)

</div>
