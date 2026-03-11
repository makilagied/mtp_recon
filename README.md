# DSE MTP vs Bank Statement Reconciliation

Reconcile DSE MTP trades against bank statements: match rows where the **DSE control number** appears in the bank narration/details and **amount paid** equals the **credit** amount.

## Setup

```bash
cd d:\mtp_recon
pip install -r requirements.txt
```

## Run the UI

```bash
streamlit run app.py
```

Then open the URL shown in the terminal (usually http://localhost:8501).

## Usage

1. **Upload files**
   - **MTP trades**: Excel file with at least a control-number column and an amount-paid column.
   - **Bank statement**: Excel file with a narration/description/details column and a credit column.

2. **Map columns**
   - Choose which MTP column is the **DSE control number** and which is **amount paid**.
   - Choose which bank column is **Narration/Description/Details** and which is **Credit**.

3. **Reconcile**
   - Click **Reconcile**. A row is matched when:
     - The control number appears **exactly** (as a substring) in the bank narration text.
     - The amount paid equals the bank credit (small float tolerance allowed).

4. **Results**
   - **Matched**: MTP rows paired with one bank row.
   - **Unmatched MTP**: MTP rows with no matching bank row.
   - **Unmatched bank**: Bank rows not used in any match.
   - **Multi-matches**: MTP rows that matched more than one bank row (for manual review).

5. **Export**
   - Use **Download results (Excel)** to get all four result sheets in one file.

## Logic (script only)

Use `reconcile.reconcile()` from `reconcile.py` with your DataFrames and column names; it returns a `ReconciliationResult` with `matched`, `unmatched_mtp`, `unmatched_bank`, and `multi_matches` DataFrames.

---

## Developer

**Erick D. Makilagi** · [GitHub @makilagied](https://github.com/makilagied)  
Software engineer focused on FinTech systems, automation, and scalable web applications.

[☕ Buy me a coffee](https://snippe.me/pay/makilagied)
