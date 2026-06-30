# 119 vs Mgmt C Month-End Equity Reconciliation Web App

This is a full-stack replacement for the Excel VBA macro `Combined_119_MgmtC_MonthEnd_Reconciliation_Approval`.

## What it does

- Uploads two raw Excel files:
  - RAW 119 Report
  - RAW Mgmt C Report
- Identifies the correct main sheets automatically.
- Validates required headers and returns clear missing-column errors.
- Builds adjusted 119 and adjusted Mgmt C data in memory.
- Reconciles by `Trade ID + Broker` using a `0.01` tolerance.
- Detects:
  - Brokerage differences
  - Missing trades from either side
  - Broker mismatches
  - Duplicate Trade ID + Broker keys
  - Missing broker values
  - GS discount involvement flags
- Displays results in browser tabs.
- Exports a final workbook with these sheets:
  - Summary
  - 119 vs Mgmt C
  - Mgmt C vs 119
  - All Exceptions
  - GS Discount
  - Issues
  - Approval

## Project structure

```text
equity-reconciliation-app/
  backend/
    app/main.py
    app/services/reconciliation.py
    requirements.txt
    run_dev.py
  frontend/
    src/App.jsx
    src/styles.css
    package.json
    index.html
```

## Run locally

### 1. Backend

```bash
cd backend
python -m venv .venv
source .venv/bin/activate   # Windows PowerShell: .venv\Scripts\Activate.ps1
pip install -r requirements.txt
python run_dev.py
```

The API runs at:

```text
http://127.0.0.1:8001
```

### 2. Frontend

Open a second terminal:

```bash
cd frontend
npm install
npm run dev
```

Open the Vite URL shown in the terminal, usually:

```text
http://localhost:5173
```

If Vite uses `5174`, that is also fine; the backend CORS allows both ports.

## Important implementation notes

The backend keeps all core reconciliation logic server-side. It does not depend on Excel formulas for processing.

### 119 logic

- Main sheet excludes `BrokerList`, `Pivot`, and `Adjusted_119_InMemory`.
- Required headers include `Deal Nb` and `Desk`.
- `ID = "2" + BW value`.
- `Broker` is mapped from the Dealer code in column H using the embedded dictionary from the VBA macro.
- `Bro = Net Bro + Amount Z + Amount X` when product column C is `EQEXT` or `EQEXI`.
- Otherwise, `Bro = Net Bro + Amount Z + Amount W`.
- Aggregates by `Broker + ID` and keeps only `Sum of Bro` outside `-0.5` to `0.5`.

### Mgmt C logic

- Main sheet excludes `TRS`, `Pivot`, `Brokerage Data`, and `Adjusted_MgmtC_InMemory`.
- Required headers include `Trade ID` and `Brokerage`.
- Broker is derived from column D using sequential fill-down behavior.
- Trade ID cleanup removes hyphen suffixes and normalizes numeric IDs.
- Aggregates by `Broker + Trade ID` and keeps only `Total` outside `-0.5` to `0.5`.

### Issues tab logic

Issues are built from All Exceptions where GS Involvement Flag is blank, then:

1. Brokerage Difference rows are de-duplicated by `Trade ID + Broker`.
2. Rows with `Difference 119 - Mgmt C` between `-1` and `1` are removed, while blank differences are preserved.
3. Rows with `Fees` between `-1` and `1` are removed only when difference is not blank.
4. Final Issues keep rows where `Discount %` is zero or `Status` contains `missing from 119`.
5. `Comments` is added for reviewer notes.
6. The reporting period label is added above the Issues table.

## Export filename

The workbook filename is:

```text
119 vs Mgmt C Thru [latest date].xlsx
```

The latest date is taken from a `Thru` date in the 119 filename when present; otherwise it uses the latest trade date from the 119 report; otherwise today's date.
