import pandas as pd
from zscore.picker import _pick
from zscore.formatter import COL_ORDER, Z_COL, _finalize_zdf
from utils.logging_utils import log


def compute_zscore(bs, is_df, company_id, company_name):
    # -------------------------------------------------
    # Basic validation
    # -------------------------------------------------
    if bs is None or bs.empty or is_df is None or is_df.empty:
        return pd.DataFrame(columns=COL_ORDER)

    # =================================================
    # Balance Sheet
    # =================================================
    CA = _pick(bs, ["สินทรัพย์หมุนเวียน"], "Current Assets")
    CL = _pick(bs, ["หนี้สินหมุนเวียน"], "Current Liab")
    TA = _pick(bs, ["สินทรัพย์รวม"], "Total Assets")
    TL = _pick(bs, ["หนี้สินรวม"], "Total Liab")
    EQ = _pick(bs, ["ส่วนของผู้ถือหุ้น", "ส่วนของเจ้าของ"], "Equity")

    RE = _pick(
        bs,
        [
            "กำไรสะสม",
            r"re:กำไร\s*\(\s*ขาดทุน\s*\)\s*สะสม",
            r"re:กำไรสะสม\s*\(\s*ขาดทุนสะสม\s*\)",
        ],
        "Retained Earn",
    )

    # =================================================
    # Income Statement
    # =================================================
    SALES = _pick(
        is_df,
        [
            "รายได้รวม",
            "รายได้จากการขาย",
            "รายได้จากการให้บริการ",
            r"re:ยอดขาย|ขาย",
        ],
        "Sales",
    )

    EBIT = _pick(
        is_df,
        [
            "EBIT",
            "กำไรจากการดำเนินงาน",
            "กำไรจากกิจกรรมดำเนินงาน",
            r"re:กำไร.*จากการดำเนินงาน",
        ],
        "EBIT",
    )

    # ---- EBIT fallback: PBT + Interest ----
    if EBIT.isna().all():
        PBT = _pick(
            is_df,
            ["กำไร(ขาดทุน) ก่อนภาษี", r"re:กำไร.*ก่อน.*ภาษี"],
            "PBT",
        )

        INT_EBIT = _pick(
            is_df,
            ["ดอกเบี้ยจ่าย", "ต้นทุนทางการเงิน", "ค่าใช้จ่ายดอกเบี้ย"],
            "Interest Exp",
        )

        if not PBT.isna().all():
            EBIT = PBT.add(INT_EBIT, fill_value=0)
            log("ℹ EBIT calculated = PBT + Interest")

    # =================================================
    # Net Income / Interest / Tax (FIXED)
    # =================================================
    NI = _pick(
        is_df,
        [
            r"re:กำไร\s*\(\s*ขาดทุน\s*\)\s*สุทธิ",
            r"re:กำไรสุทธิ.*",
        ],
        "Net Income",
    )

    INT = _pick(
        is_df,
        [
            "ดอกเบี้ยจ่าย",
            "ต้นทุนทางการเงิน",
            "ค่าใช้จ่ายดอกเบี้ย",
            r"re:ดอกเบี้ย.*",
        ],
        "Interest Exp",
    )

    TAX = _pick(
        is_df,
        [
            "ภาษี",
            "ภาษีเงินได้",
            r"re:ค่าใช้จ่าย.*ภาษี.*",
        ],
        "Tax",
    )

    # ---- Force 0 if missing (สำคัญมาก) ----
    if NI.isna().all():
        NI = NI.fillna(0)
        log("ℹ Net Income not found → fill 0")

    if INT.isna().all():
        INT = INT.fillna(0)
        log("ℹ Interest Exp not found → fill 0")

    if TAX.isna().all():
        TAX = TAX.fillna(0)
        log("ℹ Tax not found → fill 0")

    # ---- Retained Earnings fallback ----
    if RE.isna().all() and not NI.isna().all():
        RE = NI.cumsum()
        log("ℹ Retained Earn calculated = cumulative Net Income")

    # =================================================
    # Merge (year index)
    # =================================================
    df = pd.concat(
        [CA, CL, TA, RE, NI, INT, TAX, TL, SALES, EBIT, EQ],
        axis=1,
    )

    df.columns = [
        "Current Assets",
        "Current Liab",
        "Total Assets",
        "Retained Earn",
        "Net Income",
        "Interest Exp",
        "Tax",
        "Total Liab",
        "Sales",
        "EBIT",
        "Equity",
    ]

    df.index.name = "year"
    df = df.dropna(how="all")

    if df.empty:
        return pd.DataFrame(columns=COL_ORDER)

    # =================================================
    # Ratios
    # =================================================
    df["บริษัท"] = company_name
    df["Working Cap"] = df["Current Assets"] - df["Current Liab"]

    TA_ = df["Total Assets"].replace(0, pd.NA)
    TL_ = df["Total Liab"].replace(0, pd.NA)

    df["X1"] = df["Working Cap"] / TA_
    df["X2"] = df["Retained Earn"] / TA_
    df["X3"] = df["EBIT"] / TA_
    df["X4"] = df["Equity"] / TL_
    df["X5"] = df["Sales"] / TA_

    # =================================================
    # ✅ Z-Score (FULL + NON-SALES MODEL)
    # =================================================
    z_full = df[["X1", "X2", "X3", "X4", "X5"]].notna().all(axis=1)

    df.loc[z_full, Z_COL] = (
        1.2 * df.loc[z_full, "X1"]
        + 1.4 * df.loc[z_full, "X2"]
        + 3.3 * df.loc[z_full, "X3"]
        + 0.6 * df.loc[z_full, "X4"]
        + 1.0 * df.loc[z_full, "X5"]
    )

    z_non_sales = (
        df[["X1", "X2", "X3", "X4"]].notna().all(axis=1)
        & df[Z_COL].isna()
    )

    df.loc[z_non_sales, Z_COL] = (
        0.717 * df.loc[z_non_sales, "X1"]
        + 0.847 * df.loc[z_non_sales, "X2"]
        + 3.107 * df.loc[z_non_sales, "X3"]
        + 0.420 * df.loc[z_non_sales, "X4"]
    )

    # =================================================
    # Risk label
    # =================================================
    def risk_th(z):
        if pd.isna(z):
            return "ไม่ทราบ"
        if z > 2.90:
            return "ปลอดภัย"
        if z >= 1.23:
            return "เฝ้าระวัง"
        return "อันตราย"

    df["สถานะ"] = df[Z_COL].apply(risk_th)

    df = df.reset_index().rename(columns={"year": "ปี"})
    return _finalize_zdf(df)