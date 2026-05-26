from config import OUTPUT_DIR
from utils.excel_utils import _autosize, _format_numbers, _colorize
from utils.logging_utils import log
from zscore.formatter import _finalize_zdf
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os

def save_zscore_per_company(zdf, cid):
    """
    Export Z-Score รายบริษัท เป็นไฟล์ Excel แยก
    """
    if zdf is None or zdf.empty:
        log(f"⚠️ Z-Score empty for {cid} → skip company file")
        return

    out_path = os.path.join(OUTPUT_DIR, cid, f"zscore_{cid}.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.title = "Z_SCORE"

    zdf = _finalize_zdf(zdf)

    for r in dataframe_to_rows(zdf, index=False, header=True):
        ws.append(r)

    ws.freeze_panes = "A2"
    _autosize(ws)
    _format_numbers(ws)
    _colorize(ws)

    wb.save(out_path)
    log(f"💾 Saved company Z-Score → {out_path}")