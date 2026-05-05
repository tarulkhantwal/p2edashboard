# 📊 Jira Dashboard Visualizer (v5)

A Streamlit-based interactive dashboard for analyzing Jira issue data from Excel/CSV exports.
Built for flexible reporting across **quarters, half-year, and full-year views** with rich visual insights.

---

## 🚀 Features

* 📂 **Multi-sheet Excel support**

  * Automatically reads multiple sheets (Q1, Q2, etc.)
  * Supports:

    * Quarterly (Q1, Q2…)
    * Half-year (H1, H2)
    * Full-year (FY)

* 📊 **Interactive visualizations**

  * Overall volume (Bug vs RFH with % split)
  * Valid vs Invalid (bar + pie)
  * Root cause breakdown
  * Component distribution
  * Version distribution
  * Resolution type analysis
  * Type of request

* 🎯 **Dynamic filtering**

  * Project
  * Priority
  * Status
  * Type of Request
  * Validity

* 📈 **Smart time aggregation**

  * Quarter → Monthly view
  * Half-year / Full-year → Quarterly view

* 📤 **Export capabilities**

  * Filtered dataset download
  * Custom export builder

---

## 🧱 Tech Stack

* Python
* Streamlit
* Pandas
* Plotly
* OpenPyXL

---

## 📁 Project Structure

```
.
├── jira_dashboard_v5.py
├── JIRA_Report_2026.xlsx
├── requirements.txt
└── README.md
```

---

## ⚙️ Installation

### 1. Clone the repo

```
git clone https://github.com/your-username/jira-dashboard.git
cd jira-dashboard
```

### 2. Install dependencies

```
pip install -r requirements.txt
```

---

## ▶️ Run the App

```
streamlit run jira_dashboard_v5.py
```

---

## 📊 Data Format

* Input file: `JIRA_Report_2026.xlsx`
* Supports:

  * Multiple sheets (Q1_2026, Q2_2026, etc.)
* Key columns expected:

  * `Resolution Date`
  * `Type of Request` (Bug / RFH)
  * `Validity` (Valid / Invalid)
  * `Priority`
  * `Status`
  * `Component/s`
  * `Version`
  * `Root Cause Resolution`

> Column names are automatically trimmed to handle spaces like `Validity `

---

## 🧠 Key Logic Highlights

* **Period Selector**

  * Automatically detects sheets and builds:

    * Q1, Q2, H1, H2, FY views

* **Overall Volume Chart**

  * Stacked bars (Bug vs RFH)
  * Shows:

    ```
    Count (Percentage %)
    ```
  * Adapts granularity:

    * Quarter → Monthly
    * H1/FY → Quarterly

* **Validity Handling**

  * Uses clean `Valid / Invalid` values
  * Displayed as:

    * Bar chart
    * Pie chart

---

## 🛠️ Customization

You can easily:

* Add new filters in sidebar
* Modify color schemes (Adobe palette included)
* Add new charts using Plotly
* Extend period logic for new sheet formats

---

## ⚠️ Known Considerations

* Ensure consistent column naming across sheets
* Large Excel files may impact load time
* Duplicate Streamlit charts require unique `key=` values (already handled in v5)

---

## 🌐 Deployment Options

* Streamlit Community Cloud (recommended)
* Azure / AWS / GCP
* Docker container
* Power Pages (for public embedding)

---

## 👨‍💻 Author

Built for internal Jira analytics and visualization.

---

## 📄 License

MIT License (or update as per your org policy)
