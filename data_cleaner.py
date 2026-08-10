"""
data_cleaner.py — Financial data cleaning engine
Detects issues in TB / GL / any financial Excel upload
"""

import pandas as pd
import numpy as np
from typing import Optional


# ── Column aliases ─────────────────────────────────────────────────────────────
COL_ALIASES = {
    "account":  ["Account", "account", "Acc", "Code", "GL_Code", "AccountCode"],
    "name":     ["Account_Name", "AccountName", "Description", "Name", "account_name"],
    "debit":    ["Debit", "debit", "Dr", "DR"],
    "credit":   ["Credit", "credit", "Cr", "CR"],
    "opening":  ["Opening_Balance", "Opening", "Open", "opening"],
    "closing":  ["Closing_Balance", "Closing", "Close", "closing"],
    "date":     ["Date", "date", "Trans_Date", "TransDate", "Posting_Date"],
    "amount":   ["Amount", "amount", "Value", "Net_Amount"],
    "vendor":   ["Vendor", "vendor", "Supplier", "Party", "Counterparty"],
}


def _find_col(df: pd.DataFrame, key: str) -> Optional[str]:
    for alias in COL_ALIASES[key]:
        if alias in df.columns:
            return alias
    return None


def _severity(critical=False, warning=False) -> str:
    if critical:
        return "Critical"
    if warning:
        return "Warning"
    return "Info"


# ── Main cleaning function ─────────────────────────────────────────────────────
def clean(df: pd.DataFrame) -> dict:
    issues = []
    fixes  = {}   # col → fixed series

    total_rows = len(df)

    # 1. Missing required columns
    required = ["account", "debit", "credit"]
    missing_cols = [r for r in required if _find_col(df, r) is None]
    if missing_cols:
        issues.append({
            "type":     "Missing Columns",
            "severity": "Critical",
            "count":    len(missing_cols),
            "detail":   f"Required columns not found: {', '.join(missing_cols)}. Check column headers.",
            "rows":     [],
        })

    col_acc  = _find_col(df, "account")
    col_dr   = _find_col(df, "debit")
    col_cr   = _find_col(df, "credit")
    col_name = _find_col(df, "name")
    col_open = _find_col(df, "opening")
    col_clos = _find_col(df, "closing")
    col_date = _find_col(df, "date")

    df_clean = df.copy()

    # 2. Blank / null account codes
    if col_acc:
        blank_acc = df_clean[col_acc].isna() | (df_clean[col_acc].astype(str).str.strip() == "")
        if blank_acc.any():
            rows = (df_clean.index[blank_acc] + 2).tolist()  # +2 for header + 1-based
            issues.append({
                "type":     "Blank Account Code",
                "severity": "Critical",
                "count":    int(blank_acc.sum()),
                "detail":   f"{int(blank_acc.sum())} rows have no account code — cannot be mapped to financials.",
                "rows":     rows[:10],
            })

    # 3. Null / non-numeric debit / credit
    for col_key, col in [("debit", col_dr), ("credit", col_cr)]:
        if col:
            non_num = pd.to_numeric(df_clean[col], errors="coerce").isna() & df_clean[col].notna()
            if non_num.any():
                rows = (df_clean.index[non_num] + 2).tolist()
                issues.append({
                    "type":     f"Non-Numeric {col_key.title()}",
                    "severity": "Critical",
                    "count":    int(non_num.sum()),
                    "detail":   f"{int(non_num.sum())} rows have non-numeric values in {col} column.",
                    "rows":     rows[:10],
                })
            # fix: coerce to numeric
            df_clean[col] = pd.to_numeric(df_clean[col], errors="coerce").fillna(0)

    # 4. Negative debit or credit (should always be ≥ 0)
    if col_dr and col_cr:
        neg_dr = df_clean[col_dr] < 0
        neg_cr = df_clean[col_cr] < 0
        neg = neg_dr | neg_cr
        if neg.any():
            rows = (df_clean.index[neg] + 2).tolist()
            issues.append({
                "type":     "Negative Dr/Cr Values",
                "severity": "Warning",
                "count":    int(neg.sum()),
                "detail":   f"{int(neg.sum())} rows have negative debit or credit — typically indicates reversal entries or data error.",
                "rows":     rows[:10],
            })

    # 5. Dr/Cr imbalance (total debit ≠ total credit)
    if col_dr and col_cr:
        total_dr = df_clean[col_dr].sum()
        total_cr = df_clean[col_cr].sum()
        diff = abs(total_dr - total_cr)
        if diff > 0.01:
            issues.append({
                "type":     "Trial Balance Imbalance",
                "severity": "Critical",
                "count":    1,
                "detail":   f"Total Debit ({total_dr:,.2f}) ≠ Total Credit ({total_cr:,.2f}). Difference: {diff:,.2f}. TB does not balance.",
                "rows":     [],
            })

    # 6. Duplicate account codes (same account appearing more than once)
    if col_acc:
        dup = df_clean[col_acc].astype(str).str.strip().duplicated(keep=False)
        dup = dup & ~(df_clean[col_acc].astype(str).str.strip() == "")
        if dup.any():
            rows = (df_clean.index[dup] + 2).tolist()
            dup_codes = df_clean.loc[dup, col_acc].unique().tolist()[:5]
            issues.append({
                "type":     "Duplicate Account Codes",
                "severity": "Warning",
                "count":    int(dup.sum()),
                "detail":   f"{int(dup.sum())} rows share duplicate account codes: {dup_codes}. May cause double-counting in financials.",
                "rows":     rows[:10],
            })

    # 7. Closing balance mismatch: Opening + Dr - Cr ≠ Closing
    if col_open and col_clos and col_dr and col_cr:
        df_clean[col_open] = pd.to_numeric(df_clean[col_open], errors="coerce").fillna(0)
        df_clean[col_clos] = pd.to_numeric(df_clean[col_clos], errors="coerce").fillna(0)
        expected_close = df_clean[col_open] + df_clean[col_dr] - df_clean[col_cr]
        mismatch = (expected_close - df_clean[col_clos]).abs() > 0.5
        if mismatch.any():
            rows = (df_clean.index[mismatch] + 2).tolist()
            issues.append({
                "type":     "Closing Balance Mismatch",
                "severity": "Warning",
                "count":    int(mismatch.sum()),
                "detail":   f"{int(mismatch.sum())} rows: Opening + Debit - Credit ≠ Closing Balance. Likely manual override or rounding error.",
                "rows":     rows[:10],
            })

    # 8. Missing account names
    if col_name:
        blank_name = df_clean[col_name].isna() | (df_clean[col_name].astype(str).str.strip() == "")
        if blank_name.any():
            rows = (df_clean.index[blank_name] + 2).tolist()
            issues.append({
                "type":     "Missing Account Names",
                "severity": "Info",
                "count":    int(blank_name.sum()),
                "detail":   f"{int(blank_name.sum())} rows have no account description — reports will show blank labels.",
                "rows":     rows[:10],
            })

    # 9. Zero debit AND zero credit (ghost rows)
    if col_dr and col_cr:
        ghost = (df_clean[col_dr] == 0) & (df_clean[col_cr] == 0)
        if ghost.any():
            rows = (df_clean.index[ghost] + 2).tolist()
            issues.append({
                "type":     "Zero-Value Rows",
                "severity": "Info",
                "count":    int(ghost.sum()),
                "detail":   f"{int(ghost.sum())} rows have zero debit and zero credit — can be safely removed.",
                "rows":     rows[:10],
            })
            # fix: remove ghost rows
            df_clean = df_clean[~ghost].reset_index(drop=True)

    # ── Summary stats ──────────────────────────────────────────────────────────
    critical_count = sum(1 for i in issues if i["severity"] == "Critical")
    warning_count  = sum(1 for i in issues if i["severity"] == "Warning")
    info_count     = sum(1 for i in issues if i["severity"] == "Info")

    return {
        "total_rows":      total_rows,
        "clean_rows":      len(df_clean),
        "issues":          issues,
        "critical_count":  critical_count,
        "warning_count":   warning_count,
        "info_count":      info_count,
        "total_issues":    len(issues),
        "df_clean":        df_clean,
        "is_balanced":     critical_count == 0,
    }
