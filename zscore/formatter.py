import pandas as pd

COL_ORDER = [
    "บริษัท", "ปี",
    "Current Assets", "Current Liab", "Total Assets", "Retained Earn",
    "Net Income", "Interest Exp", "Tax", "Total Liab",
    "Sales", "Working Cap", "EBIT", "Equity",
    "X1", "X2", "X3", "X4", "X5", "Z'-Score", "สถานะ"
]

MONEY_COLS = [
    "Current Assets", "Current Liab", "Total Assets", "Retained Earn",
    "Net Income", "Interest Exp", "Tax", "Total Liab",
    "Sales", "Working Cap", "EBIT", "Equity"
]
RATIO_COLS = ["X1", "X2", "X3", "X4", "X5"]
Z_COL = "Z'-Score"

def _finalize_zdf(df):
    df = df.copy()

    if "บริษัท" not in df.columns:
        df["บริษัท"] = pd.NA
    if "ปี" not in df.columns:
        df["ปี"] = pd.NA

    df["ปี"] = pd.to_numeric(df["ปี"], errors="coerce").astype("Int64")

    for c in MONEY_COLS:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").round(2)
        else:
            df[c] = pd.NA

    for c in RATIO_COLS:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").round(4)
        else:
            df[c] = pd.NA

    if Z_COL in df.columns:
        df[Z_COL] = pd.to_numeric(df[Z_COL], errors="coerce").round(2)
    else:
        df[Z_COL] = pd.NA

    if "สถานะ" not in df.columns:
        df["สถานะ"] = pd.NA

    for c in COL_ORDER:
        if c not in df.columns:
            df[c] = pd.NA

    df = df[COL_ORDER].sort_values(["บริษัท", "ปี"], kind="stable")
    return df