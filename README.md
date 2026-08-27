# 📊 Python Excel Automation — Sales Report Generator

> An automated Python data processing tool that converts raw sales Excel files into presentation-ready, chart-rich reports in seconds.

![Python](https://img.shields.io/badge/Python-3.10+-blue?logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-2.x-green?logo=pandas)
![License](https://img.shields.io/badge/License-MIT-yellow)

---

## 🎯 Problem

Businesses spend **hours every week** manually preparing sales reports in Excel:
- Cleaning messy data
- Calculating totals and averages
- Finding top-performing products
- Creating charts and visualisations
- Formatting everything for management

## 💡 Solution

This tool **automates the entire process**. Drop in your raw sales Excel file and get a complete, presentation-ready report with:
- KPI summary dashboard
- Top products & categories analysis
- Monthly revenue trends with charts
- City and sales rep performance breakdown
- Clean, formatted Excel output with embedded charts

## ✨ Features

| Feature | Description |
|---------|-------------|
| 🧹 **Auto Data Cleaning** | Handles missing values, normalises columns, parses dates |
| 📈 **Revenue Analytics** | Total revenue, average order value, monthly trends |
| 🏆 **Top Products** | Ranks products by revenue with horizontal bar charts |
| 🥧 **Category Breakdown** | Pie chart showing revenue share by category |
| 🌆 **City Analysis** | Revenue breakdown by city |
| 📊 **Chart Generation** | Professional PNG charts embedded in the report |
| 📄 **Multi-Sheet Report** | Summary, Top Products, Monthly Revenue, Raw Data tabs |
| 🎨 **Professional Styling** | Branded colours, borders, number formatting |

## 🛠️ Technologies

- **Python 3.10+**
- **Pandas** — data manipulation & analysis
- **OpenPyXL** — Excel file generation with charts
- **Matplotlib** — high-quality chart rendering

## 📁 Project Structure

```
python-excel-automation/
├── excel_reporter.py        # Main report generator
├── generate_sample_data.py  # Creates realistic sample data
├── requirements.txt         # Dependencies
├── sample_data/
│   └── sales_data.xlsx      # Sample input (auto-generated)
└── reports/
    ├── sales_report.xlsx    # Generated report
    └── charts/              # PNG chart images
```

## 🚀 Installation

```bash
# Clone the repository
git clone https://github.com/itu-itis22-gunayy21/python-excel-automation.git
cd python-excel-automation

# Create virtual environment (optional but recommended)
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows

# Install dependencies
pip install -r requirements.txt
```

## ▶️ How It Works

### 1. Generate sample data (first time only)
```bash
python generate_sample_data.py
```
This creates `sample_data/sales_data.xlsx` with 500 realistic sales records.

### 2. Run the report generator
```bash
python excel_reporter.py                          # uses default sample data
python excel_reporter.py your_sales_file.xlsx     # use your own data
```

### 3. Check the output
```
reports/
├── sales_report.xlsx   ← Your professional report
└── charts/
    ├── monthly_revenue.png
    ├── top_products.png
    ├── category_pie.png
    └── city_revenue.png
```

## 📸 Screenshots

### Generated Report — Summary Tab
The summary sheet contains key KPIs at the top followed by embedded charts:

| KPI | Value |
|-----|-------|
| Total Revenue | ₺2,456,789.00 |
| Total Orders | 500 |
| Average Order Value | ₺4,913.58 |

### Charts
The system generates 4 professional charts:
- **Monthly Revenue** — Bar chart with trend
- **Top 10 Products** — Horizontal bar chart
- **Category Breakdown** — Pie chart with percentages
- **City Revenue** — Comparative bar chart

## 🔧 Customisation

You can easily adapt this for your business:

```python
# Change input/output paths
input_path = "your_data.xlsx"
output_path = "reports/custom_report.xlsx"

# The system auto-detects columns with these keywords:
# order, date, product, category, quantity, unit, total, city, sales, payment, customer, discount
```

## 📝 License

MIT License — free for personal and commercial use.

---

**Built by Yasin Günay** — Computer Engineering, Istanbul Technical University

*Need a custom Excel automation solution? Contact me*

