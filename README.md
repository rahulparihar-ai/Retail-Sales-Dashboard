# 📊 Retail Sales Analytics Dashboard
### BEE Lab Project | JECRC Foundation | B.Tech First Year

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![PowerBI](https://img.shields.io/badge/Power_BI-F2C811?style=for-the-badge&logo=powerbi&logoColor=black)
![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)
![HTML5](https://img.shields.io/badge/HTML5-E34F26?style=for-the-badge&logo=html5&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)

> **An end-to-end Business Intelligence project** featuring an interactive Excel Dashboard, a Power BI-style HTML Dashboard with live filters & 3-month sales forecast, and a complete professional report — built entirely with Python automation.

---

## 🌐 Live Demo

**👉 Open `bi-dashboard/Retail_Sales_BI_Dashboard_PREMIUM.html` in any browser**

Features:
- 🎨 Dark glassmorphism theme
- 📊 6 interactive charts (Line, Bar, Donut, Region, Sub-Category, Forecast)
- 🔍 Live slicers — Year, Category, Region, Segment
- 📈 3-month forecast with 95% confidence band
- 📄 One-click PDF export

---

## 📁 Project Structure

```
BEE-Retail-Sales-Dashboard/
│
├── 📂 dataset/
│   └── retail_sales_final.csv          # 300 orders | 2020-2022 | 14 columns
│
├── 📂 excel-dashboard/
│   └── Store_Sales_Dashboard_FINAL.xlsx # 6 sheets | KPI Cards | 3 PivotCharts
│
├── 📂 bi-dashboard/
│   ├── Retail_Sales_BI_Dashboard_PREMIUM.html  # ⭐ Interactive BI Dashboard
│   ├── DAX_Measures.dax                        # 16 Power BI DAX measures
│   └── PowerBI_Setup_Guide.md                  # Step-by-step Power BI guide
│
├── 📂 report/
│   ├── BEE_Lab_Report_RAHUL_PARIHAR.docx       # Full project report (Word)
│   ├── BEE_Lab_Presentation_RAHUL_PARIHAR.pptx # 9-slide presentation
│   ├── Report_Rahul_Parihar.pdf                # Report PDF
│   ├── Dashboard_Rahul_Parihar.pdf             # Dashboard PDF
│   └── Presentation_Rahul_Parihar.pdf          # Presentation PDF
│
├── 📂 scripts/                                  # Python automation scripts
│   ├── analyze_data.py                          # Dataset statistics extractor
│   ├── build_excel_final.py                     # Excel dashboard generator
│   ├── build_word_final.py                      # Word report generator
│   ├── build_pptx_final.py                      # PowerPoint generator
│   ├── build_premium_dashboard.py               # HTML dashboard generator
│   └── build_final_submission.py               # Final ZIP builder
│
└── README.md
```

---

## 📊 Dataset Summary

| Metric | Value |
|---|---|
| **Total Records** | 300 orders |
| **Date Range** | January 2020 – December 2022 |
| **Total Sales** | $3,34,395 |
| **Total Profit** | $51,612 |
| **Profit Margin** | 15.4% |
| **YoY Growth (2021)** | +61.9% |
| **Categories** | Technology · Furniture · Office Supplies |
| **Regions** | East · West · Central · South |

---

## 🔑 Key Findings

- 🥇 **Technology** dominates at 75% of revenue ($2,50,980) — Laptops & Phones lead
- 📈 **2021** was the best year with **+61.9% YoY growth**
- 📉 **2022** saw a **-6.8% correction** — warrants pricing strategy review
- 🌎 **West region** leads all regions with $1,01,195 in sales
- ⚠️ **Tables & Bookcases** generate losses due to heavy discounting (15–20%)
- 🔮 **Forecast**: Sales projected to stabilize in Jan–Mar 2023

---

## 🛠️ Tech Stack

| Tool | Purpose |
|---|---|
| **Python 3.x** | Data processing & file generation |
| **openpyxl** | Excel dashboard automation |
| **python-docx** | Word report generation |
| **python-pptx** | PowerPoint generation |
| **playwright** | HTML → PDF conversion |
| **Chart.js** | Interactive charts in HTML dashboard |
| **Microsoft Excel** | Final dashboard view |
| **Microsoft Power BI** | DAX measures & data modeling |

---

## 🚀 How to Run Scripts

### Prerequisites
```bash
pip install openpyxl python-docx python-pptx playwright
python -m playwright install chromium
```

### Generate All Files
```bash
# 1. Analyze dataset
python scripts/analyze_data.py

# 2. Build Excel Dashboard
python scripts/build_excel_final.py

# 3. Build Word Report
python scripts/build_word_final.py

# 4. Build PowerPoint
python scripts/build_pptx_final.py

# 5. Build Premium HTML Dashboard
python scripts/build_premium_dashboard.py

# 6. Build Final Submission ZIP
python scripts/build_final_submission.py
```

---

## 📈 Dashboard Screenshots

> Open `bi-dashboard/Retail_Sales_BI_Dashboard_PREMIUM.html` in Chrome/Edge for best experience.

**Dashboard Features:**
- ✅ 10 KPI Cards (Total Sales, Profit, Margin, Orders, Units, 3-Year Comparison)
- ✅ Monthly Sales Trend — 2020 vs 2021 vs 2022
- ✅ Sales & Profit by Category
- ✅ Sales by Segment (Donut) & Region (Bar)
- ✅ Top Sub-Categories by Profit
- ✅ 3-Month Forecast with 95% Confidence Interval
- ✅ Live Slicers — filter by Year, Category, Region, Segment
- ✅ Export PDF button

---

## 📋 DAX Measures (Power BI)

```dax
-- Total Sales
Total Sales = SUM(SalesData[Sales])

-- Profit Margin
Profit Margin % = DIVIDE([Total Profit], [Total Sales], 0)

-- Year-to-Date Sales
YTD Sales = TOTALYTD([Total Sales], DateTable[Date])

-- Year-over-Year Growth
YoY Sales Growth % = 
DIVIDE([Total Sales] - [Sales Previous Year], [Sales Previous Year], 0)
```
> See `bi-dashboard/DAX_Measures.dax` for all 16 measures.

---

## 👨‍💻 Author

**Rahul Parihar**
- 🎓 B.Tech First Year | Roll No: 268
- 🏛️ JECRC Foundation, Jaipur
- 📘 Subject: Business & Economic Environment (BEE) Lab
- 👨‍🏫 Faculty: Prof. Ram Singh
- 📅 April 2026

---

## 📄 License

This project is open-source under the [MIT License](LICENSE).

---

⭐ **If this helped you, please star the repository!**
