"""
data_cleaner.py — Financial data cleaning engine
Two-pass cleaning:
  Pass 1 (openpyxl): structural Excel fixes — merged cells, colors, hidden rows/cols, formulas
  Pass 2 (pandas):   data quality checks — blanks, duplicates, imbalance, whitespace, dates
"""

import io
import re
import pandas as pd
import numpy as np
from typing import Optional


# ── Column aliases ─────────────────────────────────────────────────────────────
COL_ALIASES = {
    "account":  ["Account", "account", "Acc", "Code", "GL_Code", "AccountCode", "Account Code"],
    "name":     ["Account_Name", "AccountName", "Description", "Name", "account_name", "Account Name"],
    "debit":    ["Debit", "debit", "Dr", "DR", "Debit Amount"],
    "credit":   ["Credit", "credit", "Cr", "CR", "Credit Amount"],
    "opening":  ["Opening_Balance", "Opening", "Open", "opening", "Opening Balance"],
    "closing":  ["Closing_Balance", "Closing", "Close", "closing", "Closing Balance"],
    "date":     ["Date", "date", "Trans_Date", "TransDate", "Posting_Date", "Transaction Date"],
    "amount":   ["Amount", "amount", "Value", "Net_Amount", "Net Amount"],
    "vendor":   ["Vendor", "vendor", "Supplier", "Party", "Counterparty"],
}


def _find_col(df: pd.DataFrame, key: str) -> Optional[str]:
    for alias in COL_ALIASES[key]:
        if alias in df.columns:
            return alias
    return None


# ═══════════════════════════════════════════════════════════════════════════════
# PASS 1 — openpyxl structural fixes (call before reading into pandas)
# ═══════════════════════════════════════════════════════════════════════════════
def clean_excel_structural(raw_bytes: bytes) -> tuple[bytes, list]:
    """
    Fix Excel-level issues: merged cells, colors, hidden rows/cols, formulas, blank rows/cols.
    Returns (cleaned_excel_bytes, list_of_fix_dicts)
    """
    try:
        import openpyxl
        from openpyxl.styles import PatternFill
    except ImportError:
        return raw_bytes, []

    fixes = []
    wb = openpyxl.load_workbook(io.BytesIO(raw_bytes))
    ws = wb.active

    # 1. Unmerge cells + fill down value
    merged_ranges = list(ws.merged_cells.ranges)
    merge_count = len(merged_ranges)
    if merge_count > 0:
        for mr in merged_ranges:
            top_val = ws.cell(mr.min_row, mr.min_col).value
            ws.unmerge_cells(str(mr))
            for row in ws.iter_rows(min_row=mr.min_row, max_row=mr.max_row,
                                    min_col=mr.min_col, max_col=mr.max_col):
                for cell in row:
                    if cell.value is None:
                        cell.value = top_val
        fixes.append({
            "type": "Unmerged Cells",
            "severity": "Info",
            "count": merge_count,
            "detail": f"Unmerged {merge_count} merged cell range(s) and filled values down — merged cells break pivot tables and formulas.",
            "rows": [],
        })

    # 2. Strip all cell color / background formatting
    no_fill = PatternFill(fill_type=None)
    color_count = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.fill and cell.fill.fill_type and cell.fill.fill_type != "none":
                cell.fill = no_fill
                color_count += 1
    if color_count > 0:
        fixes.append({
            "type": "Color Formatting Removed",
            "severity": "Info",
            "count": color_count,
            "detail": f"Removed color formatting from {color_count} cell(s) — manual highlights can mislead automated processing.",
            "rows": [],
        })

    # 3. Unhide hidden rows
    hidden_rows = [rn for rn, rd in ws.row_dimensions.items() if rd.hidden]
    if hidden_rows:
        for rn in hidden_rows:
            ws.row_dimensions[rn].hidden = False
        fixes.append({
            "type": "Hidden Rows Unhidden",
            "severity": "Warning",
            "count": len(hidden_rows),
            "detail": f"{len(hidden_rows)} hidden row(s) found and unhidden — hidden rows can contain data that affects totals.",
            "rows": hidden_rows[:10],
        })

    # 4. Unhide hidden columns
    hidden_cols = [cl for cl, cd in ws.column_dimensions.items() if cd.hidden]
    if hidden_cols:
        for cl in hidden_cols:
            ws.column_dimensions[cl].hidden = False
        fixes.append({
            "type": "Hidden Columns Unhidden",
            "severity": "Warning",
            "count": len(hidden_cols),
            "detail": f"{len(hidden_cols)} hidden column(s) found and unhidden — hidden columns may contain amounts excluded from totals.",
            "rows": [],
        })

    # 5. Strip formulas → keep cached value or None
    formula_count = 0
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.value = None  # Can't evaluate without Excel engine
                formula_count += 1
    if formula_count > 0:
        fixes.append({
            "type": "Formulas Removed",
            "severity": "Info",
            "count": formula_count,
            "detail": f"{formula_count} formula cell(s) cleared — formulas with broken references silently return wrong values.",
            "rows": [],
        })

    # 6. Delete fully blank rows (all cells None or empty string)
    blank_rows_deleted = 0
    rows_to_delete = []
    for row in ws.iter_rows():
        if all(cell.value is None or str(cell.value).strip() == "" for cell in row):
            rows_to_delete.append(row[0].row)
    for rn in sorted(rows_to_delete, reverse=True):
        ws.delete_rows(rn)
        blank_rows_deleted += 1
    if blank_rows_deleted > 0:
        fixes.append({
            "type": "Blank Rows Removed",
            "severity": "Info",
            "count": blank_rows_deleted,
            "detail": f"{blank_rows_deleted} fully blank row(s) deleted — blank rows break imports into accounting systems.",
            "rows": [],
        })

    # 7. Delete fully blank columns
    blank_cols_deleted = 0
    cols_to_delete = []
    for col in ws.iter_cols():
        if all(cell.value is None or str(cell.value).strip() == "" for cell in col):
            cols_to_delete.append(col[0].column)
    for cn in sorted(cols_to_delete, reverse=True):
        ws.delete_cols(cn)
        blank_cols_deleted += 1
    if blank_cols_deleted > 0:
        fixes.append({
            "type": "Blank Columns Removed",
            "severity": "Info",
            "count": blank_cols_deleted,
            "detail": f"{blank_cols_deleted} fully blank column(s) deleted.",
            "rows": [],
        })

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read(), fixes


# ═══════════════════════════════════════════════════════════════════════════════
# PASS 2 — pandas data quality checks
# ═══════════════════════════════════════════════════════════════════════════════
def clean(df: pd.DataFrame, structural_fixes: list = None) -> dict:
    issues = list(structural_fixes or [])
    total_rows = len(df)
    df_clean = df.copy()

    col_acc  = _find_col(df_clean, "account")
    col_dr   = _find_col(df_clean, "debit")
    col_cr   = _find_col(df_clean, "credit")
    col_name = _find_col(df_clean, "name")
    col_open = _find_col(df_clean, "opening")
    col_clos = _find_col(df_clean, "closing")
    col_date = _find_col(df_clean, "date")
    col_amt  = _find_col(df_clean, "amount")

    # ── A. Trim whitespace from all string columns ───────────────────────────
    str_cols = df_clean.select_dtypes(include="object").columns.tolist()
    trimmed_cells = 0
    for col in str_cols:
        before = df_clean[col].copy()
        df_clean[col] = df_clean[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
        trimmed_cells += (before != df_clean[col]).sum()
    if trimmed_cells > 0:
        issues.append({
            "type": "Whitespace Trimmed",
            "severity": "Info",
            "count": int(trimmed_cells),
            "detail": f"Trimmed leading/trailing spaces from {trimmed_cells} cell(s) — 'Salaries ' ≠ 'Salaries' in lookups.",
            "rows": [],
        })

    # ── B. Numbers stored as text ────────────────────────────────────────────
    for col_key, col in [("Debit", col_dr), ("Credit", col_cr), ("Amount", col_amt)]:
        if col is None:
            continue
        # Values that look numeric but are stored as string
        text_nums = df_clean[col].apply(
            lambda x: isinstance(x, str) and re.sub(r"[,\s]", "", x).replace(".", "", 1).lstrip("-").isdigit()
        )
        if text_nums.any():
            rows = (df_clean.index[text_nums] + 2).tolist()
            issues.append({
                "type": f"Numbers as Text ({col_key})",
                "severity": "Critical",
                "count": int(text_nums.sum()),
                "detail": f"{int(text_nums.sum())} {col_key} values stored as text — SUM() returns 0, all totals wrong.",
                "rows": rows[:10],
            })
            # Fix: convert to numeric
            df_clean[col] = df_clean[col].apply(
                lambda x: float(re.sub(r"[,\s]", "", x)) if isinstance(x, str) else x
            )

    # ── C. Missing required columns ──────────────────────────────────────────
    required = ["account", "debit", "credit"]
    missing_cols = [r for r in required if _find_col(df_clean, r) is None]
    if missing_cols:
        issues.append({
            "type": "Missing Columns",
            "severity": "Critical",
            "count": len(missing_cols),
            "detail": f"Required columns not found: {', '.join(missing_cols)}. Check column headers.",
            "rows": [],
        })

    # ── D. Blank account codes ───────────────────────────────────────────────
    if col_acc:
        blank_acc = df_clean[col_acc].isna() | (df_clean[col_acc].astype(str).str.strip() == "")
        if blank_acc.any():
            rows = (df_clean.index[blank_acc] + 2).tolist()
            issues.append({
                "type": "Blank Account Code",
                "severity": "Critical",
                "count": int(blank_acc.sum()),
                "detail": f"{int(blank_acc.sum())} rows have no account code — cannot be mapped to financial statements.",
                "rows": rows[:10],
            })

    # ── E. Non-numeric debit / credit ────────────────────────────────────────
    for col_key, col in [("Debit", col_dr), ("Credit", col_cr)]:
        if col:
            non_num = pd.to_numeric(df_clean[col], errors="coerce").isna() & df_clean[col].notna()
            if non_num.any():
                rows = (df_clean.index[non_num] + 2).tolist()
                issues.append({
                    "type": f"Non-Numeric {col_key}",
                    "severity": "Critical",
                    "count": int(non_num.sum()),
                    "detail": f"{int(non_num.sum())} rows have non-numeric values in {col} — totals will be incorrect.",
                    "rows": rows[:10],
                })
            df_clean[col] = pd.to_numeric(df_clean[col], errors="coerce").fillna(0)

    # ── F. Negative debit or credit ──────────────────────────────────────────
    if col_dr and col_cr:
        neg = (df_clean[col_dr] < 0) | (df_clean[col_cr] < 0)
        if neg.any():
            rows = (df_clean.index[neg] + 2).tolist()
            issues.append({
                "type": "Negative Dr/Cr Values",
                "severity": "Warning",
                "count": int(neg.sum()),
                "detail": f"{int(neg.sum())} rows have negative Dr/Cr — typically reversal entries; verify they are intentional.",
                "rows": rows[:10],
            })

    # ── G. Trial Balance imbalance ───────────────────────────────────────────
    if col_dr and col_cr:
        total_dr = df_clean[col_dr].sum()
        total_cr = df_clean[col_cr].sum()
        diff = abs(total_dr - total_cr)
        if diff > 0.01:
            issues.append({
                "type": "Trial Balance Imbalance",
                "severity": "Critical",
                "count": 1,
                "detail": f"Total Debit ({total_dr:,.2f}) ≠ Total Credit ({total_cr:,.2f}). Difference: {diff:,.2f} — TB does not balance.",
                "rows": [],
            })

    # ── H. Duplicate account codes ───────────────────────────────────────────
    if col_acc:
        acc_series = df_clean[col_acc].astype(str).str.strip()
        dup = acc_series.duplicated(keep=False) & (acc_series != "")
        if dup.any():
            rows = (df_clean.index[dup] + 2).tolist()
            dup_codes = df_clean.loc[dup, col_acc].unique().tolist()[:5]
            issues.append({
                "type": "Duplicate Account Codes",
                "severity": "Warning",
                "count": int(dup.sum()),
                "detail": f"{int(dup.sum())} rows share duplicate account codes: {dup_codes} — may cause double-counting.",
                "rows": rows[:10],
            })

    # ── I. Exact duplicate rows ──────────────────────────────────────────────
    dup_rows = df_clean.duplicated(keep=False)
    if dup_rows.any():
        rows = (df_clean.index[dup_rows] + 2).tolist()
        issues.append({
            "type": "Duplicate Rows",
            "severity": "Critical",
            "count": int(dup_rows.sum()),
            "detail": f"{int(dup_rows.sum())} rows are exact duplicates — could represent double-posted journal entries.",
            "rows": rows[:10],
        })
        # Fix: remove duplicates, keep first
        df_clean = df_clean.drop_duplicates(keep="first").reset_index(drop=True)

    # ── J. Round-number anomaly detection ───────────────────────────────────
    amt_col = col_dr or col_amt
    if amt_col:
        round_flag = df_clean[amt_col].apply(
            lambda x: isinstance(x, (int, float)) and x > 0 and x % 10000 == 0
        )
        if round_flag.any():
            rows = (df_clean.index[round_flag] + 2).tolist()
            issues.append({
                "type": "Round Number Anomaly",
                "severity": "Info",
                "count": int(round_flag.sum()),
                "detail": f"{int(round_flag.sum())} entries are exact multiples of 10,000 — may indicate estimates or manual entries worth reviewing.",
                "rows": rows[:10],
            })

    # ── K. Closing balance mismatch ──────────────────────────────────────────
    if col_open and col_clos and col_dr and col_cr:
        df_clean[col_open] = pd.to_numeric(df_clean[col_open], errors="coerce").fillna(0)
        df_clean[col_clos] = pd.to_numeric(df_clean[col_clos], errors="coerce").fillna(0)
        expected = df_clean[col_open] + df_clean[col_dr] - df_clean[col_cr]
        mismatch = (expected - df_clean[col_clos]).abs() > 0.5
        if mismatch.any():
            rows = (df_clean.index[mismatch] + 2).tolist()
            issues.append({
                "type": "Closing Balance Mismatch",
                "severity": "Warning",
                "count": int(mismatch.sum()),
                "detail": f"{int(mismatch.sum())} rows: Opening + Dr - Cr ≠ Closing — likely manual override or rounding error.",
                "rows": rows[:10],
            })

    # ── L. Date format inconsistency ─────────────────────────────────────────
    if col_date:
        parsed_dates = pd.to_datetime(df_clean[col_date], errors="coerce", infer_datetime_format=True)
        bad_dates = parsed_dates.isna() & df_clean[col_date].notna()
        if bad_dates.any():
            rows = (df_clean.index[bad_dates] + 2).tolist()
            issues.append({
                "type": "Invalid Date Format",
                "severity": "Warning",
                "count": int(bad_dates.sum()),
                "detail": f"{int(bad_dates.sum())} dates could not be parsed — inconsistent formats (e.g. '01-Jan-24' vs '1/1/2024') break date filters.",
                "rows": rows[:10],
            })
        else:
            # Standardise all dates to YYYY-MM-DD
            df_clean[col_date] = parsed_dates.dt.strftime("%Y-%m-%d")

    # ── M. Missing account names ─────────────────────────────────────────────
    if col_name:
        blank_name = df_clean[col_name].isna() | (df_clean[col_name].astype(str).str.strip() == "")
        if blank_name.any():
            rows = (df_clean.index[blank_name] + 2).tolist()
            issues.append({
                "type": "Missing Account Names",
                "severity": "Info",
                "count": int(blank_name.sum()),
                "detail": f"{int(blank_name.sum())} rows have no account description — reports will show blank labels.",
                "rows": rows[:10],
            })

    # ── N. Zero-value ghost rows ─────────────────────────────────────────────
    if col_dr and col_cr:
        ghost = (df_clean[col_dr] == 0) & (df_clean[col_cr] == 0)
        if ghost.any():
            rows = (df_clean.index[ghost] + 2).tolist()
            issues.append({
                "type": "Zero-Value Rows",
                "severity": "Info",
                "count": int(ghost.sum()),
                "detail": f"{int(ghost.sum())} rows have zero Debit and zero Credit — safely removed.",
                "rows": rows[:10],
            })
            df_clean = df_clean[~ghost].reset_index(drop=True)

    # ── Summary ───────────────────────────────────────────────────────────────
    critical_count = sum(1 for i in issues if i["severity"] == "Critical")
    warning_count  = sum(1 for i in issues if i["severity"] == "Warning")
    info_count     = sum(1 for i in issues if i["severity"] == "Info")

    return {
        "total_rows":     total_rows,
        "clean_rows":     len(df_clean),
        "issues":         issues,
        "critical_count": critical_count,
        "warning_count":  warning_count,
        "info_count":     info_count,
        "total_issues":   len(issues),
        "df_clean":       df_clean,
        "is_balanced":    critical_count == 0,
    }
