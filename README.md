# Portfolio Tracker (Python)

A command-line based **portfolio management and analytics tool** designed for tracking, updating, and visualizing stock holdings.
The system automates price retrieval, portfolio performance calculation, and exports a dynamic Excel dashboard with charts and gain/loss summaries.

---

## 🚀 Key Features

* **Automated Price Fetching** – Retrieves live market data using Yahoo Finance API (`yfinance`).
* **Portfolio Analytics** – Calculates total portfolio value, gain/loss (RM), and percentage change.
* **Data Persistence** – Saves and loads portfolio data through CSV files for easy reuse.
* **Excel Dashboard Export** – Generates a professional Excel report with charts, summary tables, and conditional formatting using `openpyxl`.
* **Interactive CLI** – Add, edit, or remove stocks directly through a user-friendly terminal menu.
* **Visual Feedback** – Colored output for gain/loss representation via `colorama`.

---

## 🧩 Tech Stack

* **Language:** Python
* **Libraries:** `yfinance`, `tabulate`, `openpyxl`, `colorama`, `csv`, `os`, `time`

---

## ⚙️ How It Works

1. **Data Loading**

   * On startup, the program loads existing portfolio data from `portfolio_tracker.csv` (if available).

2. **User Interaction (Main Menu)**

   * View portfolio summary
   * Add or remove stocks
   * Edit stock information
   * Refresh live prices
   * Export Excel dashboard

3. **Excel Export**

   * Creates an interactive Excel file (`portfolio_dashboard.xlsx`) featuring:

     * Portfolio summary table
     * Conditional formatting for gain/loss
     * Pie chart (value distribution)
     * Bar chart (gain/loss by stock)

---

## 📁 File Structure

```
portfolio_tracker/
│
├── portfolio_tracker.py      # Main application script
├── portfolio_tracker.csv     # Saved portfolio data (auto-created)
├── portfolio_dashboard.xlsx  # Exported Excel dashboard (auto-generated)
└── README.md                 # Project documentation
```

---

## 💡 Example Workflow

1. Run the program:

   ```bash
   python portfolio_tracker.py
   ```
2. Add your stock holdings (e.g., NESTLE, 4707.KL).
3. Fetch live prices automatically or enter manually if unavailable.
4. View a summarized portfolio table with live performance data.
5. Export to Excel for reporting or record-keeping.

---

## 📊 Output Preview (Console)

```
📊 Portfolio Summary:
╒═════════════╤════════╤════════╤════════════╤═════════════════╤══════════╤══════════════╤══════════════╕
│ Stocks      │ Ticker │ Shares │ Buy Price  │ Current Price   │ Value    │ Gain/Loss    │ Gain/Loss %  │
╞═════════════╪════════╪════════╪════════════╪═════════════════╪══════════╪══════════════╪══════════════╡
│ NESTLE BHD  │ 4707.KL│  100   │ 123.000    │ 130.000         │ 13000.00 │ +700.00      │ +5.69%       │
╘═════════════╧════════╧════════╧════════════╧═════════════════╧══════════╧══════════════╧══════════════╛
💰 Total Portfolio Value: RM 13,000.00  
📈 Total Gain/Loss: RM +700.00  
```

---

## 🧠 Purpose

This project demonstrates:

* **Practical financial programming** using Python
* **Integration of data automation, analytics, and reporting**
* **Application of Python to real-world finance workflows**

Designed to showcase technical competence in **data-driven investment management tools** for resumes and portfolios.
