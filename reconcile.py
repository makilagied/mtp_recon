"""
DSE MTP trades vs bank statement reconciliation.
Matches when: dse_control_number appears exactly in narration/details AND amount_paid == credit.
"""

from __future__ import annotations

import pandas as pd
from dataclasses import dataclass


@dataclass
class ReconciliationResult:
    matched: pd.DataFrame      # MTP row + bank row + keys
    unmatched_mtp: pd.DataFrame
    unmatched_bank: pd.DataFrame
    multi_matches: pd.DataFrame  # MTP rows that matched more than one bank row


def _normalize_value(val) -> str:
    """String for matching: strip and treat NaN as empty."""
    if pd.isna(val):
        return ""
    return str(val).strip()


def _amount_match(paid: float, credit: float, tol: float = 1e-6) -> bool:
    """Compare amounts with small tolerance for float."""
    if pd.isna(paid) or pd.isna(credit):
        return False
    try:
        p, c = float(paid), float(credit)
        return abs(p - c) <= tol
    except (TypeError, ValueError):
        return False


def reconcile(
    mtp_df: pd.DataFrame,
    mtp_control_col: str,
    mtp_amount_col: str,
    bank_df: pd.DataFrame,
    bank_narration_col: str,
    bank_credit_col: str,
    amount_tolerance: float = 1e-6,
) -> ReconciliationResult:
    """
    Reconcile MTP trades to bank statement rows.

    - Control number must appear exactly in the bank narration (as a substring).
    - amount_paid must equal bank credit (within tolerance).

    Returns matched pairs, unmatched MTP rows, unmatched bank rows, and multi-matches.
    """
    mtp_control_col = mtp_control_col.strip()
    mtp_amount_col = mtp_amount_col.strip()
    bank_narration_col = bank_narration_col.strip()
    bank_credit_col = bank_credit_col.strip()

    if mtp_control_col not in mtp_df.columns:
        raise KeyError(f"MTP sheet missing column: '{mtp_control_col}'")
    if mtp_amount_col not in mtp_df.columns:
        raise KeyError(f"MTP sheet missing column: '{mtp_amount_col}'")
    if bank_narration_col not in bank_df.columns:
        raise KeyError(f"Bank sheet missing column: '{bank_narration_col}'")
    if bank_credit_col not in bank_df.columns:
        raise KeyError(f"Bank sheet missing column: '{bank_credit_col}'")

    bank_used = set()  # bank row indices already matched
    matched_rows = []
    multi_match_mtp_indices = set()

    for mtp_idx, mtp_row in mtp_df.iterrows():
        control = _normalize_value(mtp_row[mtp_control_col])
        amount_paid = mtp_row[mtp_amount_col]

        if not control:
            continue

        candidates = []
        for bank_idx, bank_row in bank_df.iterrows():
            if bank_idx in bank_used:
                continue
            narration = _normalize_value(bank_row[bank_narration_col])
            credit = bank_row[bank_credit_col]
            if control not in narration:
                continue
            if not _amount_match(amount_paid, credit, amount_tolerance):
                continue
            candidates.append(bank_idx)

        if len(candidates) == 1:
            bank_used.add(candidates[0])
            matched_rows.append(
                {
                    "mtp_index": mtp_idx,
                    "bank_index": candidates[0],
                    "dse_control_number": control,
                    "amount_paid": amount_paid,
                    "bank_credit": bank_df.loc[candidates[0], bank_credit_col],
                }
            )
        elif len(candidates) > 1:
            multi_match_mtp_indices.add(mtp_idx)

    # Build matched dataframe (one row per match: MTP fields + bank fields)
    mtp_cols = [c for c in mtp_df.columns]
    bank_cols = [c for c in bank_df.columns]
    # Prefix bank columns to avoid clash
    bank_cols_prefixed = [f"bank_{c}" for c in bank_cols]

    matched_list = []
    for m in matched_rows:
        mtp_idx = m["mtp_index"]
        bank_idx = m["bank_index"]
        row = mtp_df.loc[mtp_idx].to_dict()
        bank_dict = {f"bank_{k}": v for k, v in bank_df.loc[bank_idx].items()}
        row.update(bank_dict)
        row["_mtp_index"] = mtp_idx
        row["_bank_index"] = bank_idx
        matched_list.append(row)

    matched_df = pd.DataFrame(matched_list) if matched_list else pd.DataFrame()

    # Unmatched MTP: not in any matched pair
    matched_mtp_indices = {m["mtp_index"] for m in matched_rows}
    unmatched_mtp_df = mtp_df.loc[
        ~mtp_df.index.isin(matched_mtp_indices | multi_match_mtp_indices)
    ].copy()

    # Multi-match: MTP rows that had multiple bank matches (user to resolve)
    multi_matches_df = mtp_df.loc[list(multi_match_mtp_indices)].copy() if multi_match_mtp_indices else pd.DataFrame()

    # Unmatched bank: rows not used in any match
    unmatched_bank_df = bank_df.loc[~bank_df.index.isin(bank_used)].copy()

    return ReconciliationResult(
        matched=matched_df,
        unmatched_mtp=unmatched_mtp_df,
        unmatched_bank=unmatched_bank_df,
        multi_matches=multi_matches_df,
    )
