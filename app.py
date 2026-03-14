"""
DSE MTP vs Bank Statement Reconciliation — Streamlit UI.
Upload MTP trades and bank statement Excel files, map columns, run reconciliation.
"""

import io
import re
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from reconcile import reconcile, ReconciliationResult

# Excel highlighting: green = matched; white = unmatched (bank); yellow = unmatched (MTP control)
FILL_MATCHED = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
FILL_WHITE = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
FILL_UNMATCHED_MTP = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")


def _safe_for_display(df: pd.DataFrame) -> pd.DataFrame:
    """Convert columns that can overflow PyArrow (e.g. large ints) to string for display."""
    if df is None or df.empty:
        return df
    out = df.copy()
    for col in out.columns:
        try:
            dtype = out[col].dtype
            if pd.api.types.is_integer_dtype(dtype):
                out[col] = out[col].astype("object").astype(str)
            elif dtype == object:
                # object columns may contain Python ints too large for C long
                out[col] = out[col].astype(str)
        except (TypeError, ValueError, OverflowError):
            out[col] = out[col].astype(str)
    return out


def _apply_row_highlights(ws, ncols: int, status_col_idx: int):
    """Apply green fill to matched rows, white to unmatched (for bank statement)."""
    for row_idx in range(2, ws.max_row + 1):
        cell = ws.cell(row=row_idx, column=status_col_idx)
        fill = FILL_MATCHED if (cell.value == "Matched") else FILL_WHITE
        for col_idx in range(1, ncols + 1):
            ws.cell(row=row_idx, column=col_idx).fill = fill


def _excel_with_highlighted_rows(df: pd.DataFrame, status_column: str = "Recon status") -> bytes:
    """Write dataframe to Excel and fill rows: green if matched, white if unmatched (bank)."""
    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl")
    buf.seek(0)
    wb = load_workbook(buf)
    ws = wb.active
    try:
        status_col_idx = list(df.columns).index(status_column) + 1  # 1-based
    except ValueError:
        status_col_idx = df.shape[1]
    _apply_row_highlights(ws, df.shape[1], status_col_idx)
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


def _excel_with_control_column_highlight(
    df: pd.DataFrame,
    status_column: str = "Recon status",
    control_column: str | None = None,
) -> bytes:
    """Write dataframe to Excel; only the control number column is filled: green if matched, white if unmatched."""
    if control_column is None or control_column not in df.columns:
        control_column = df.columns[0]
    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl")
    buf.seek(0)
    wb = load_workbook(buf)
    ws = wb.active
    try:
        status_col_idx = list(df.columns).index(status_column) + 1
    except ValueError:
        status_col_idx = len(df.columns)
    control_col_idx = list(df.columns).index(control_column) + 1
    for row_idx in range(2, ws.max_row + 1):
        status_cell = ws.cell(row=row_idx, column=status_col_idx)
        fill = FILL_MATCHED if (status_cell.value == "Matched") else FILL_UNMATCHED_MTP
        ws.cell(row=row_idx, column=control_col_idx).fill = fill
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


def _sanitize_sheet_name(name: str, max_len: int = 31) -> str:
    """Excel sheet names: max 31 chars, no : \\ / ? * [ ]."""
    name = re.sub(r'[:\\/?*\[\]]', "_", str(name))
    return name[:max_len] if len(name) > max_len else name or "Sheet"

st.set_page_config(page_title="DSE MTP Reconciliation", layout="wide")
st.title("DSE MTP Trades vs Bank Statement Reconciliation")
st.caption(
    "Match MTP trades to bank statement rows where the DSE control number appears "
    "in the narration/details and the amount paid equals the credit amount."
)

# Session state for uploaded data and result
if "result" not in st.session_state:
    st.session_state.result = None

# --- File upload ---
st.header("1. Upload files")
col1, col2 = st.columns(2)

with col1:
    mtp_files = st.file_uploader(
        "MTP trades (Excel) — upload one or more",
        type=["xlsx", "xls"],
        key="mtp_upload",
        accept_multiple_files=True,
        help="Each file's first sheet is used. Map control number and amount per file.",
    )
with col2:
    bank_files = st.file_uploader(
        "Bank statement(s) (Excel) — upload one or more",
        type=["xlsx", "xls"],
        key="bank_upload",
        accept_multiple_files=True,
        help="Upload multiple files to reconcile against all statements together. Each file's first sheet is used.",
    )


def load_excel(uploaded_file, sheet_index: int = 0) -> pd.DataFrame | None:
    if uploaded_file is None:
        return None
    try:
        return pd.read_excel(uploaded_file, sheet_name=sheet_index)
    except Exception as e:
        st.error(f"Could not read {getattr(uploaded_file, 'name', 'file')}: {e}")
        return None


def _idx(cols, *prefer):
    for p in prefer:
        if p in cols:
            return cols.index(p)
    return 0


# Store each MTP file as (name, df)
if "mtp_files_data" not in st.session_state:
    st.session_state.mtp_files_data = []

if mtp_files:
    st.session_state.mtp_files_data = []
    for i, f in enumerate(mtp_files):
        name = getattr(f, "name", f"MTP_{i+1}")
        df = load_excel(f)
        if df is not None and not df.empty:
            st.session_state.mtp_files_data.append((name, df))
else:
    st.session_state.mtp_files_data = []
    st.session_state.result = None

# Store each bank statement as (name, df) so we can map columns per file
if "bank_files_data" not in st.session_state:
    st.session_state.bank_files_data = []

if bank_files:
    # (Re)build list of (name, df) for current set of uploads
    st.session_state.bank_files_data = []
    for i, f in enumerate(bank_files):
        name = getattr(f, "name", f"Statement {i+1}")
        df = load_excel(f)
        if df is not None and not df.empty:
            st.session_state.bank_files_data.append((name, df))
else:
    st.session_state.bank_files_data = []
    st.session_state.result = None

mtp_files_data = st.session_state.mtp_files_data
bank_files_data = st.session_state.bank_files_data

# --- Column mapping ---
st.header("2. Map columns")

if mtp_files_data and bank_files_data:
    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("MTP file(s) — map per file")
        st.caption("Each file can have different column names. Choose control number and amount for each.")
        mtp_control_choices = []
        mtp_amount_choices = []
        for i, (name, df) in enumerate(mtp_files_data):
            mcols = list(df.columns)
            with st.expander(f"**{name}** ({len(df)} rows)", expanded=(i == 0)):
                mi = _idx(mcols, "dse_control_number", "control_number", "reference", "Reference")
                ai = _idx(mcols, "amount_paid", "amount", "credit", "Amount")
                ctrl = st.selectbox(
                    "DSE control number",
                    options=mcols,
                    index=mi,
                    key=f"mtp_ctrl_{i}",
                )
                amt = st.selectbox(
                    "Amount paid",
                    options=mcols,
                    index=ai,
                    key=f"mtp_amt_{i}",
                )
                mtp_control_choices.append(ctrl)
                mtp_amount_choices.append(amt)
                st.dataframe(_safe_for_display(df.head(5)), width="stretch", height=120)
    with col_b:
        st.subheader("Bank statement(s) — map per file")
        st.caption("Each file can have different column names. Choose narration and credit for each.")
        bank_narration_choices = []
        bank_credit_choices = []
        for i, (name, df) in enumerate(bank_files_data):
            bcols = list(df.columns)
            with st.expander(f"**{name}** ({len(df)} rows)", expanded=(i == 0)):
                ni = _idx(bcols, "Narration", "Description", "Details", "Particulars", "Remarks")
                ci = _idx(bcols, "Credit", "credit", "Amount", "credit_amount", "Credit Amount")
                narr = st.selectbox(
                    "Narration / Description / Details",
                    options=bcols,
                    index=ni,
                    key=f"bn_narr_{i}",
                )
                cred = st.selectbox(
                    "Credit amount",
                    options=bcols,
                    index=ci,
                    key=f"bn_cred_{i}",
                )
                bank_narration_choices.append(narr)
                bank_credit_choices.append(cred)
                st.dataframe(_safe_for_display(df.head(5)), width="stretch", height=120)

    # Build combined MTP dataframe with standardized columns
    mtp_combined_frames = []
    for i, (name, df) in enumerate(mtp_files_data):
        part = df.copy()
        part = part.rename(columns={
            mtp_control_choices[i]: "_recon_control",
            mtp_amount_choices[i]: "_recon_amount",
        })
        part["_mtp_source"] = name
        mtp_combined_frames.append(part)
    mtp_df = pd.concat(mtp_combined_frames, ignore_index=True)

    # Build combined bank dataframe with standardized narration/credit columns
    combined_frames = []
    for i, (name, df) in enumerate(bank_files_data):
        part = df.copy()
        narr_col = bank_narration_choices[i]
        cred_col = bank_credit_choices[i]
        part = part.rename(columns={
            narr_col: "_recon_narration",
            cred_col: "_recon_credit",
        })
        part["_bank_source"] = name
        combined_frames.append(part)
    bank_df = pd.concat(combined_frames, ignore_index=True)

    st.header("3. Run reconciliation")
    if st.button("Reconcile", type="primary"):
        try:
            res = reconcile(
                mtp_df=mtp_df,
                mtp_control_col="_recon_control",
                mtp_amount_col="_recon_amount",
                bank_df=bank_df,
                bank_narration_col="_recon_narration",
                bank_credit_col="_recon_credit",
            )
            st.session_state.result = res
            st.success("Reconciliation completed.")
        except KeyError as e:
            st.error(str(e))
        except Exception as e:
            st.error(f"Reconciliation failed: {e}")

    st.markdown("[☕ Buy me a coffee](https://snippe.me/pay/makilagied)")

    # --- Results ---
    res: ReconciliationResult | None = st.session_state.result
    if res is not None:
        st.header("4. Results")

        n_matched = len(res.matched)
        n_unmatched_mtp = len(res.unmatched_mtp)
        n_unmatched_bank = len(res.unmatched_bank)
        n_multi = len(res.multi_matches)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Matched", n_matched)
        m2.metric("Unmatched MTP", n_unmatched_mtp, help="MTP rows with no matching bank entry")
        m3.metric("Unmatched bank", n_unmatched_bank, help="Bank rows not matched to any MTP")
        m4.metric("Multi-matches", n_multi, help="MTP rows that matched more than one bank row — review manually")

        tab1, tab2, tab3, tab4 = st.tabs(
            ["Matched", "Unmatched MTP", "Unmatched bank", "Multi-matches"]
        )
        with tab1:
            if not res.matched.empty:
                st.dataframe(_safe_for_display(res.matched), width="stretch")
            else:
                st.info("No matched rows.")
        with tab2:
            if not res.unmatched_mtp.empty:
                st.dataframe(_safe_for_display(res.unmatched_mtp), width="stretch")
            else:
                st.info("All MTP rows were matched.")
        with tab3:
            if not res.unmatched_bank.empty:
                st.dataframe(_safe_for_display(res.unmatched_bank), width="stretch")
            else:
                st.info("All bank rows were matched.")
        with tab4:
            if not res.multi_matches.empty:
                st.warning(
                    "These MTP rows matched more than one bank row; resolve manually."
                )
                st.dataframe(_safe_for_display(res.multi_matches), width="stretch")
            else:
                st.info("No multi-matches.")

        # Export
        st.subheader("Export")
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            res.matched.to_excel(writer, sheet_name="Matched", index=False)
            res.unmatched_mtp.to_excel(writer, sheet_name="Unmatched_MTP", index=False)
            res.unmatched_bank.to_excel(
                writer, sheet_name="Unmatched_Bank", index=False
            )
            res.multi_matches.to_excel(
                writer, sheet_name="Multi_matches", index=False
            )
        buf.seek(0)
        st.download_button(
            "Download results (Excel)",
            data=buf,
            file_name="mtp_reconciliation_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Highlighted originals: MTP = control number only (green/white); statement = full row (green/white)
        st.subheader("Download originals with highlighting")
        matched_mtp_indices = set(res.matched["_mtp_index"].values) if not res.matched.empty else set()
        matched_bank_indices = set(res.matched["_bank_index"].values) if not res.matched.empty else set()
        multi_mtp_indices = set(res.multi_matches.index) if not res.multi_matches.empty else set()

        # MTP: one downloadable file per MTP file (no merge)
        mtp_offsets = [0]
        for _name, mdf in mtp_files_data:
            mtp_offsets.append(mtp_offsets[-1] + len(mdf))
        for i, (name, mdf) in enumerate(mtp_files_data):
            start = mtp_offsets[i]
            def _mtp_status(j):
                idx = start + j
                if idx in matched_mtp_indices:
                    return "Matched"
                if idx in multi_mtp_indices:
                    return "Multi-match"
                return "Unmatched"
            mtp_export = mdf.copy()
            mtp_export["Recon status"] = [ _mtp_status(j) for j in range(len(mdf)) ]
            control_col = mtp_control_choices[i]
            mtp_highlighted_bytes = _excel_with_control_column_highlight(
                mtp_export, "Recon status", control_col
            )
            safe_name = _sanitize_sheet_name(name).rstrip("_") or "mtp"
            st.download_button(
                f"Download **{name}** (control number: green = matched, white = unmatched)",
                data=mtp_highlighted_bytes,
                file_name=f"{safe_name}_highlighted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_mtp_hl_{i}",
            )

        # Bank: one downloadable file per statement (no merge)
        bank_offsets = [0]
        for _name, bdf in bank_files_data:
            bank_offsets.append(bank_offsets[-1] + len(bdf))
        for i, (name, bdf) in enumerate(bank_files_data):
            start = bank_offsets[i]
            statuses = [
                "Matched" if (start + j) in matched_bank_indices else "Unmatched"
                for j in range(len(bdf))
            ]
            bank_export = bdf.copy()
            bank_export["Recon status"] = statuses
            bank_highlighted_bytes = _excel_with_highlighted_rows(bank_export)
            safe_name = _sanitize_sheet_name(name).rstrip("_") or "statement"
            st.download_button(
                f"Download **{name}** (row: green = matched, white = unmatched)",
                data=bank_highlighted_bytes,
                file_name=f"{safe_name}_highlighted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_bank_hl_{i}",
            )

else:
    st.info("Upload at least one MTP file and one bank statement Excel file to continue.")
    if mtp_files_data:
        st.subheader("MTP file(s) preview")
        for name, df in mtp_files_data[:3]:
            st.caption(name)
            st.dataframe(_safe_for_display(df.head(5)), width="stretch")
    if bank_files_data:
        st.subheader("Bank statement(s) preview")
        for name, df in bank_files_data[:3]:
            st.caption(name)
            st.dataframe(_safe_for_display(df.head(5)), width="stretch")

# Developer acknowledgment (sidebar)
st.sidebar.markdown("---")
st.sidebar.caption(
    "Built by [**Erick D. Makilagi**](https://github.com/makilagied) · "
    "[@makilagied](https://github.com/makilagied)"
)
