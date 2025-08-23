# Loan Interest Rate Processor

A small desktop tool (Tkinter + pandas) that ingests a workbook, evaluates loan tranches, and writes a cleaned, multi-sheet Excel file.

---

## ✨ What it does

- Opens an input Excel file that **must** contain sheets named:
  - `Report`
  - `Criteria`
- Cleans and annotates the `Report`:
  - Drops the **last 5 rows** (assumed footers/totals).
  - Adds `Accepted`, `Rejected`, and `Reason for Rejection` columns.
  - Marks tranches **Accepted** if `Tranche Type` ∈  
    `['Term Loan', 'Term Loan A', 'Term Loan B', 'Term Loan C']`; otherwise **Rejected** and the rejection reason is the original `Tranche Type`.
- Builds subsets:
  - `Accepted` — all accepted rows
  - `Rejected` — all rejected rows
  - `SWAPS` — accepted rows with only:
    - `Tranche Active Date`
    - `Tranche Maturity Date`
    - `Tranche Amount (m)`
    - `Tranche Currency`
    - `Base Rate & Margin (bps)`
- Prompts you for a save location and writes an output workbook containing sheets:
  - `Criteria` (copied through)
  - `Workings` (full annotated report)
  - `Accepted`
  - `SWAPS`
  - `Rejected`

Success and error dialogs are shown in-app.

---

## 🧰 Requirements

- **Python** 3.9+ (tested with modern 3.x)
- **Packages**
  - `pandas`
  - `openpyxl`
  - (Tkinter ships with most CPython installations)


## 🚀 Run the app

Save your script as `app.py` (or similar) and run:

```bash
python app.py
```

A simple window titled **Loan Interest Rate Processor** opens.

1. Click **Browse**.
2. Select your input Excel file (`.xlsx` or `.xls`).
3. Choose where to save the processed workbook when prompted.

---

## 📥 Input workbook contract (Loan Connector)

Your file must have:

- A sheet named **`Report`** with (at minimum) these columns **spelled exactly**:

  - `Tranche Type`
  - `Tranche Active Date`
  - `Tranche Maturity Date`
  - `Tranche Amount (m)`
  - `Tranche Currency`
  - `Base Rate & Margin (bps)`

- A sheet named **`Criteria`** (content passed through verbatim to the output).

> ⚠️ Column names are case- and space-sensitive. If names differ, the script will raise an error.

---

## 🔍 Business rules (quick reference)

- **Accepted tranches**: `Tranche Type` in  
  `Term Loan`, `Term Loan A`, `Term Loan B`, `Term Loan C`
- **Rejected tranches**: everything else
- **Reason for Rejection**: set to the non-accepted `Tranche Type`
- **Footer removal**: the bottom 5 rows of `Report` are dropped before processing

---

## 🧾 Output workbook

You’ll be prompted for a path, and the file will be saved as `.xlsx` with these sheets:

- `Criteria` — original `Criteria` content
- `Workings` — full `Report` after cleaning & new columns
- `Accepted` — accepted rows only
- `SWAPS` — accepted rows with swap-relevant columns only
- `Rejected` — rejected rows only

---

## 🧱 Limitations

- Assumes consistent sheet names and column headers.
- Assumes last 5 rows of `Report` are non-data footers.
- Drag-and-drop is not wired; only the file dialog is functional.


---

## 🩺 Troubleshooting

- **“An error occurred: …” on run**
  - Check that `Report` and `Criteria` sheets exist.
  - Verify required column names match exactly.
  - Ensure the file isn’t open in Excel (Windows file locks can block writing).
- **Dates look wrong**
  - convert to proper dates in Excel if needed

