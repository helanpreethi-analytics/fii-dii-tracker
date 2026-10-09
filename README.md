# 📊 FII & DII Flow Tracker (Jan–Jun 2026)

An Excel-based analytics workbook that tracks **Foreign Institutional Investor (FII)** and **Domestic Institutional Investor (DII)** activity across 10 Indian market sectors, compares it with **Nifty 50** movement, and produces a live dashboard with flow signals, sector rotation, and a scenario-based forecast.

> ⚠️ **Data note:** The flow dataset is **synthetic** (generated for practice and portfolio purposes) and was deliberately created with data-quality issues to demonstrate a full cleaning workflow. It should **not** be used for real investment decisions.

---

## 📌 Project Overview

FII and DII flows are among the most closely watched indicators in the Indian equity market. This project answers questions such as:

- Are FIIs and DIIs buying or selling this week, and are they moving in the **same or opposite direction**?
- Which **sectors** are attracting or losing institutional money?
- How **concentrated** is FII activity in a few sectors?
- Is there any **correlation** between FII flows and Nifty 50 daily returns?
- What could flows look like next under **Bull, Base, and Bear** scenarios?

---

## 🗂️ Repository Structure

```
fii-dii-tracker/
│
├── FII_DII_Tracker_27.xlsx     # Main workbook (all analysis lives here)
├── README.md                   # Project documentation
└── screenshots/                # Dashboard and sheet previews (add your own)
```

---

## 📑 Workbook Sheets

| Sheet | Purpose |
|---|---|
| **Dashboard** | KPI cards, daily FII / DII / Nifty table and chart. The front page of the project. |
| **FII_DII_raw_synthetic_Jan-Jun20** | Cleaned raw data stored as an Excel Table (`FII_DII_raw_synthetic_Jan_Jun2026`). About 878 rows of sector-level flows. |
| **Flow_Model** | Calculation engine: daily totals, rolling averages, divergence, sector split, concentration, correlation, signals, and the forecast. |
| **Sector_Rotation** | Weekly FII net flow by sector (27 weeks × 10 sectors) with conditional-format heatmap. |
| **Sector_PE_Snapshot** | Approximate sector P/E ratios as of 25 Jun 2026, for valuation context. |
| **Nifty 50 Historical Data** | Daily Nifty 50 price, open, high, low, volume, and change %. |
| **Data_Audit** | Log of every cleaning step plus a validation test table. |
| **SectorMap** | Lookup table that standardises inconsistent sector names. |

---

## 🧾 Data Dictionary (Raw Data Table)

| Column | Description |
|---|---|
| `Date` | Trading date (1 Jan – 30 Jun 2026) |
| `Entity_Type` | `FII` or `DII` |
| `Sector` | One of 10 sectors (see below) |
| `Gross_Buy_Cr` | Gross purchases in ₹ Crore |
| `Gross_Sell_Cr` | Gross sales in ₹ Crore |
| `Net_Flow_Cr` | Net flow = Buy − Sell (₹ Crore) |
| `Remarks` | Notes from source (for example, unit flags) |
| `Week_Num` | ISO-style week number, calculated with `WEEKNUM(Date, 2)` |

**Sectors covered:** Pharmaceuticals, Energy, Information Technology, Realty, Metals, FMCG, Banking, Telecom, Infrastructure, Automobile.

---

## 🧮 Key Metrics & Logic

| Metric | Formula / Logic |
|---|---|
| **Daily FII / DII Net** | `SUMIFS` on the raw table by date and entity type |
| **5-Day Rolling Average** | Rolling mean of FII and DII net flows |
| **Divergence Index** | `ABS(FII Net) − ABS(DII Net)`, showing which side is dominating |
| **Sector Concentration Ratio** | (Top 2 sectors' absolute FII flow) ÷ (Total absolute FII flow) |
| **FII–Nifty Correlation** | `CORREL` of daily FII net flow vs Nifty daily change % |
| **FII Threshold** | `1.5 × average absolute daily FII net flow` |
| **Signal Flag** | See the signal table below |

### 🚦 Signal Classification

| Condition | Signal |
|---|---|
| FII net ≤ −threshold | **Strong FII Selling** |
| FII net ≥ +threshold | **Strong FII Buying** |
| FII or DII net = 0 | **Neutral** |
| FII and DII move in opposite directions | **Divergent Flow** |
| FII and DII move in the same direction | **Aligned Flow** |

### 🔮 Scenario Forecast

A dropdown (cell `AD2` in `Flow_Model`) lets you pick a scenario. The forecast scales the latest 5-day average flow by the scenario multipliers.

| Scenario | FII Multiplier | DII Multiplier |
|---|---|---|
| Bull | 1.5 | 1.2 |
| Base | 1.0 | 1.0 |
| Bear | −1.3 | 0.7 |

---

## 📈 Dashboard Snapshot

The dashboard shows five headline KPIs, based on the latest week of data:

| KPI | Value |
|---|---|
| This Week FII Net (₹ Cr) | ≈ 3,652.78 |
| This Week DII Net (₹ Cr) | ≈ 652.85 |
| Current Signal | Divergent Flow |
| Sector Concentration | ≈ 0.66 |
| FII–Nifty Correlation | ≈ −0.02 (essentially none) |

Below the KPIs, a table and bar chart plot daily FII net, DII net, and Nifty close.

<!-- Add your screenshots to the screenshots/ folder and uncomment the lines below -->
<!--
![Dashboard](screenshots/dashboard.png)
![Sector Rotation Heatmap](screenshots/sector_rotation.png)
-->

---

## 🧹 Data Cleaning Summary

The raw dataset arrived messy. Every fix is documented in the **Data_Audit** sheet:

| Issue | Rows Affected | Fix |
|---|---|---|
| Duplicate header row | 1 | Removed |
| Mixed date formats (DD-MM-YYYY and MM/DD/YYYY) | 899 | Normalised |
| Inconsistent sector names (e.g. "IT Sector", "BFSI", "Auto") | 63 | Mapped using `SectorMap` |
| Values reported in Lakh instead of Crore | 13 | Divided by 100 |
| Missing `Net_Flow_Cr` | 25 | Derived as Buy − Sell |
| Corrupted rows (`#N/A` / `ERR`) | 3 | Dropped |
| Duplicate entries | 18 | Kept first occurrence per Date + Entity + Sector |
| Day/month swapped on ambiguous dates | 53 | Re-corrected in a second pass (dataset is Jan–Jun only) |

### ✅ Validation Tests (all passed)

Daily SUMIFS vs manual sum · 5-day rolling average · Divergence Index · Sector Concentration Ratio · Date bug check (no dates after 30 Jun) · Refresh test (changes ripple downstream) · Scenario toggle multipliers.

---

## 🛠️ Excel Features & Skills Demonstrated

- Structured **Excel Tables** and structured references
- `SUMIFS`, `SUMPRODUCT`, `LARGE`, `CORREL`, `WEEKNUM`, `INDEX/MATCH`, `IFERROR`
- Dynamic array and lookup functions: `XLOOKUP`, `UNIQUE`, `SORT`
- **Data validation** dropdown for scenario selection
- **Conditional formatting** heatmap for sector rotation
- Dashboard **KPI cards and charts**
- Data cleaning, standardisation, and an **audit trail**
- Model validation and testing

---

## 🚀 How to Use

1. Download `FII_DII_Tracker_27.xlsx` from this repository.
2. Open it in **Microsoft Excel 365 or Excel 2021+** (the file uses `XLOOKUP`, `UNIQUE`, and `SORT`, which older versions do not support).
3. Start on the **Dashboard** sheet.
4. Go to **Flow_Model** and change the **Selected_Scenario** dropdown (`AD2`) to Bull, Base, or Bear to see the forecast update.
5. Open **Sector_Rotation** to see which sectors saw the heaviest FII buying or selling each week.
6. Check **Data_Audit** to see how the data was cleaned and validated.

> 💡 GitHub cannot preview Excel files in the browser. Click the file, then **Download raw file** to open it locally.

---

## ⚠️ Limitations

- Flow data is **synthetic**, so insights are illustrative, not real market findings.
- The sector P/E values are **approximate** (Telecom and Infrastructure are estimated from constituents, since no standalone NSE index exists).
- Data covers only **Jan–Jun 2026** (about six months).
- The forecast is a simple multiplier model, not a statistical prediction.
- Correlation does not imply causation.

---

## 🔭 Future Improvements

- Replace synthetic data with real NSE / NSDL FII-DII data
- Add a Power BI or Tableau version of the dashboard
- Automate data refresh with Python or Power Query
- Add statistical forecasting (for example, moving-average or ARIMA models)
- Add sector valuation vs flow analysis using the P/E snapshot

---

## 👤 Author

**helanpreethi-analytics**
GitHub: [@helanpreethi-analytics](https://github.com/helanpreethi-analytics)

---

## 📄 License

This project is for educational and portfolio purposes.
