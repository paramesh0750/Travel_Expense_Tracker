# ✈️ Travel Expense Tracker

An Excel-based travel expense tracker that records trip budgets and day-to-day spending, then summarises everything with PivotTables, charts, slicers and a KPI dashboard. No macros or external dependencies: just open the workbook in Microsoft Excel.

## Features

- **Trip management**: store destination, travel dates and budget for each trip
- **Expense logging**: record every expense by date, category, payment mode and amount
- **PivotTable analysis**: spending summarised by category and by destination
- **Interactive charts**: pie chart (category split) and column chart (destination-wise spend)
- **Slicers**: filter by Date, Destination, Category and Payment Mode
- **Dashboard**: live KPIs for budget, spend and transaction statistics

## Workbook Structure

| Sheet | Purpose |
|---|---|
| **Trip Details** | Master list of trips: Trip ID, Destination, Start Date, End Date, Budget |
| **Expense Data** | Transaction log: Date, Trip ID, Destination, Category, Description, Payment Mode, Amount |
| **Pivot** | PivotTables, charts and slicers for category-wise and destination-wise analysis |
| **Dashboard** | Summary KPIs and a chart |

### Dashboard KPIs

| Metric | Formula |
|---|---|
| Total Budget | `=SUM('Trip Details'!E2:E4)` |
| Total Expense | `=SUM('Expense Data'!G:G)` |
| Remaining Budget | `=B4-B5` |
| Highest Expense | `=MAX('Expense Data'!G:G)` |
| Lowest Expense | `=MIN('Expense Data'!G:G)` |
| Average Expense | `=AVERAGE('Expense Data'!G:G)` |
| Transactions | `=COUNT('Expense Data'!G:G)` |

## Sample Data

The workbook ships with three sample trips:

| Trip ID | Destination | Dates | Budget (₹) | Spent (₹) |
|---|---|---|---|---|
| T001 | Goa | 1–5 Aug 2026 | 25,000 | 23,050 |
| T002 | Mysore | 15–17 Aug 2026 | 12,000 | 5,300 |
| T003 | Ooty | 25–29 Aug 2026 | 18,000 | 10,250 |
| | **Total** | | **55,000** | **38,600** |

**Expense categories:** Transport, Hotel, Food, Entertainment, Shopping, Fuel
**Payment modes:** Cash, Card, UPI

### Spending by category

| Category | Amount (₹) |
|---|---|
| Hotel | 16,000 |
| Transport | 10,600 |
| Shopping | 5,100 |
| Entertainment | 3,700 |
| Food | 1,800 |
| Fuel | 1,400 |

## Getting Started

1. Clone or download this repository:
   ```bash
   git clone https://github.com/paramesh0750/Travel_Expense_Tracker.git
   cd Travel_Expense_Tracker
   ```
2. Open `Travel_Expense_Tracker.xlsx` in **Microsoft Excel** 
3. Explore the **Dashboard** and **Pivot** sheets.

## How to Use

### Add a new trip
1. Go to **Trip Details** and add a new row with a unique Trip ID (e.g. `T004`), destination, dates and budget.
2. Extend the budget range in the Dashboard formula (`E2:E4` → `E2:E5`), or convert the table to an Excel Table for automatic expansion.

### Log an expense
1. Go to **Expense Data** and add a row at the bottom.
2. Fill in the date, Trip ID, destination, category, description, payment mode and amount.

### Refresh the analysis
- Go to **Data → Refresh All** (or right-click a PivotTable → **Refresh**) so the PivotTables, charts and slicers pick up new rows.
- Dashboard KPIs use whole-column formulas and update automatically.

### Filter the data
Use the slicers on the **Pivot** sheet to filter by date, destination, category or payment mode.

## Tech Stack

- Microsoft Excel
- PivotTables and PivotCharts
- Slicers
- Excel formulas (`SUM`, `MAX`, `MIN`, `AVERAGE`, `COUNT`)
- Data validation

## Ideas for Future Improvements

- Per-trip budget vs. actual comparison with over-budget alerts (conditional formatting)
- Multi-currency support
- Daily spending trend chart
- Per-person expense splitting for group trips

## Requirements

- Microsoft Excel


## Author

**Paramesh D**
GitHub: [@paramesh0750](https://github.com/paramesh0750)
