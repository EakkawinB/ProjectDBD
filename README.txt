----------------------------------------------------------------
  Mode 1 — Z-Score Batch  (main.py)
----------------------------------------------------------------

ใช้ main.py เป็น entrypoint หลักสำหรับ:

  - ดึงงบการเงินจาก DBD
  - คำนวณ Altman Z-Score
  - บันทึกผลลัพธ์และ generate report

Workflow:
  1. โหลด company_ids.xlsx
  2. ค้นหาบริษัทบน DBD
  3. ดึง Balance Sheet + Income Statement
  4. คำนวณ Z-Score
  5. save output + notify

Output:
  output/zscore.xlsx
  output/<company_id>/...


----------------------------------------------------------------
  Mode 2 — Registered Capital (addon.py)
----------------------------------------------------------------

ใช้ addon.py สำหรับดึงข้อมูล "ทุนจดทะเบียน" โดยเฉพาะ

Workflow:
  1. โหลด company_ids.xlsx
  2. ค้นหาบริษัทบน DBD
  3. ดึง:
       - Registered Capital
       - Paid-up Capital
       - Corporate Status
  4. บันทึกแบบ incremental ลง Excel

Output:
  N-AUTHORIZED-CAPITAL/Registered capital.xlsx