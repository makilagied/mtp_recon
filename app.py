"""
DSE MTP vs Bank Statement Reconciliation — Streamlit UI.
Upload MTP trades and bank statement Excel files, map columns, run reconciliation.
"""

import io
import streamlit as st
import pandas as pd
from reconcile import reconcile, ReconciliationResult


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

st.set_page_config(page_title="DSE MTP Reconciliation", layout="wide")
st.title("DSE MTP Trades vs Bank Statement Reconciliation")
st.caption(
    "Match MTP trades to bank statement rows where the DSE control number appears "
    "in the narration/details and the amount paid equals the credit amount."
)

# Session state for uploaded data and result
if "mtp_df" not in st.session_state:
    st.session_state.mtp_df = None
if "bank_df" not in st.session_state:
    st.session_state.bank_df = None
if "result" not in st.session_state:
    st.session_state.result = None

# --- File upload ---
st.header("1. Upload files")
col1, col2 = st.columns(2)

with col1:
    mtp_file = st.file_uploader(
        "MTP trades (Excel)",
        type=["xlsx", "xls"],
        key="mtp_upload",
        help="Sheet should have a control number column and an amount paid column.",
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


# Load data when files are uploaded (Streamlit re-runs on change)
if mtp_file is not None:
    st.session_state.mtp_df = load_excel(mtp_file)
else:
    st.session_state.mtp_df = None
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

mtp_df = st.session_state.mtp_df
bank_files_data = st.session_state.bank_files_data

# --- Column mapping ---
st.header("2. Map columns")

if mtp_df is not None and bank_files_data:
    mtp_cols = list(mtp_df.columns)

    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("MTP trades")
        mtp_control = st.selectbox(
            "Column for DSE control number",
            options=mtp_cols,
            index=_idx(mtp_cols, "dse_control_number", "control_number", "reference"),
            key="mtp_control",
        )
        mtp_amount = st.selectbox(
            "Column for amount paid",
            options=mtp_cols,
            index=_idx(mtp_cols, "amount_paid", "amount", "credit"),
            key="mtp_amount",
        )
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
                mtp_control_col=mtp_control,
                mtp_amount_col=mtp_amount,
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

else:
    st.info("Upload MTP trades and at least one bank statement Excel file to continue.")
    if mtp_df is not None:
        st.subheader("MTP preview")
        st.dataframe(_safe_for_display(mtp_df.head(10)), width="stretch")
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
