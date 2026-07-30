# 🏛️ NCRP Daily Transaction Analyzer

An automated web app to analyze daily NCRP (National Cyber Crime Reporting Portal) transaction data — built to make daily reporting faster and easier.

🔗 **Live App:** [ncrp-analyzer.streamlit.app](https://ncrp-analyzer.streamlit.app)

---

## 📌 What it does

- Upload any daily NCRP Excel file
- Instantly get a full breakdown across **8 analysis tabs**
- Separates **API (SBI API)** vs **CCC** transactions automatically
- Download a full **Excel report** with 15 sheets covering every angle of the data

---

## 📊 Analysis Tabs

| Tab | What it shows |
|---|---|
| 📊 Summary | Layer distribution, amount analysis, key metrics |
| 🔢 ACK & Accounts | Unique ACK numbers + Unique account numbers with full details |
| 🧅 Layers | Layer 0, 1-3, 4-5, 6-10, 10+ breakdown with API per layer |
| ⚡ Actions | All action types with count and % |
| 🗺️ States | State-wise transaction breakdown |
| 🏦 Banks | Money transfer TO bank breakdown |
| 👮 Officers | Action taken by officer breakdown |
| 🎯 API Analysis | Full API deep dive — actions, layers, states, banks, officers |

---

## 📋 Key Metrics Tracked

- Total Transactions
- Unique ACK Numbers
- Unique Account Numbers
- API (SBI API tagged) vs CCC split
- Layer-wise breakdown (Layer 0, 1-3, 4-5, 6-10, 10+)
- Disputed Amount < ₹500
- Transaction Amount < ₹500
- Transaction Amount ≥ ₹20,000

---

## 🎯 API Analysis

API transactions are identified by `Action Taken by = "SBI API"`. The API tab shows:
- Total API vs CCC count
- API action breakdown (Money Transfer TO, Put on Hold, Others etc.)
- API by Layer
- API by State
- API by Bank
- API by Officer
- Separate downloadable API Excel report

---

## 📂 Expected Input Format

Upload an Excel file with a sheet named `ViewTransactions` containing these columns:

| Column | Description |
|---|---|
| Acknowledgement No. | Unique complaint ID |
| Transaction Amount | Amount in ₹ |
| Disputed Amount | Disputed amount in ₹ |
| Layers | Transaction layer number |
| Action | Action taken on the transaction |
| Action Taken by | Officer/system who took action |
| State | State of the complaint |
| Account Number | Account involved |
| Money transfer TO Bank | Destination bank |

---

## 🚀 Run Locally

```bash
git clone https://github.com/Aryan1404xxx/NCRP-analyzer.git
cd NCRP-analyzer
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

---

## 📦 Tech Stack

- **Streamlit** — web app framework
- **Pandas** — data processing
- **Plotly** — interactive charts
- **OpenPyXL** — Excel report generation

---

## 👨‍💻 Developer

**Aryan Sinha**  
Computer Science Student  
Built during SBI Internship — to automate daily NCRP transaction reporting.
