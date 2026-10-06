# Call_Centre_Report_2023

# 📊 Call Centre Analytics Dashboard (2023)

An interactive, single-page Excel dashboard built to analyze call center operations, sales revenue, customer ratings, and representative metrics for 2023.

---

## 📌 Dashboard Preview.

---

## 📂 Workbook Architecture

The workbook is structured into 4 clean layers:

| Sheet Name | Layer Type | Description |
| :--- | :--- | :--- |
| **`Data`** | **Data Layer** | Raw tabular call logs containing caller IDs, revenue, call duration, satisfaction ratings, and geographic data. |
| **`Assets`** | **Resource Layer** | Stores representative lookup IDs (`R01`–`R05`) alongside their high-resolution caller photographs. |
| **`Pivot`** | **Calculation Layer** | Houses Pivot Tables, total aggregations, ranking formulas, and dynamic named ranges (`CallerImage`). |
| **`Dashboard`** | **Presentation Layer** | The executive user interface featuring interactive Slicers, KPI cards, charts, and dynamic caller photos. |

---

## 🚀 Key Dashboard Features & KPIs

* **Executive KPI Summary Cards:** Instant view of Total Calls (1,000), Total Revenue ($96,623), Call Duration, Average Rating (3.9/5), and Happy Callers.
* **Dynamic Representative Lookup:** Selecting a representative (e.g., `R02`) automatically triggers a linked picture lookup to update the caller's photo, call rank, revenue rank, and call percentage.
* **Demographic Breakdown:** Stacked bar chart highlighting female vs. male caller ratios across key cities (*Cincinnati*, *Cleveland*, *Columbus*).
* **Call Volume & Satisfaction Trends:** 
  * Monthly call volume trend line (Jan – Dec 2023).
  * Day-of-week volume distribution (Monday – Sunday).
  * Customer satisfaction rating spread (1 to 5 stars).

---

## 🛠️ Technical Implementation

* **Dynamic Image Lookup (Classic Excel Compatible):** Uses an `INDEX + MATCH` formula inside a Named Range (`CallerImage`) connected to a **Linked Picture** object to dynamically switch representative images without relying on `XLOOKUP`.
* **Data Modeling:** Multi-table aggregations via Pivot Tables, calculated fields, and conditional formatting rules for matrix visualization.
* **User Interactivity:** Interactive Slicers mapped to the Pivot engine to seamlessly filter performance metrics.

---

## 📂 File Structure

```text
├── Call_Centre_Analytics_Dashboard_2023.xlsx   # Main Excel Workbook
├── README.md                                   # Project Documentation
└── dashboard_preview.png                        # Dashboard Screenshot Preview
