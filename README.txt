================================================================
  ProjectDBD — DBD Financial Scraper & Altman Z-Score Calculator
================================================================

ระบบดึงงบการเงินจากเว็บ DBD (กรมพัฒนาธุรกิจการค้า) แบบ batch
และคำนวณ Altman Z-Score พร้อมแจ้งผลผ่าน Outlook (Power Automate)


----------------------------------------------------------------
  โครงสร้างโปรเจกต์
----------------------------------------------------------------

  main.py              → entrypoint หลัก (รวม 2 mode ไว้ที่นี่)
  config.py            → ค่า config ทั้งหมด (ปรับได้ที่นี่)
  secrets_manager.py   → จัดการ encrypt/decrypt secrets
  setup_secrets.py     → script ตั้งค่า secrets ครั้งแรก

  browser/             → จัดการ Playwright browser lifecycle
  checkpoint/          → บันทึกความคืบหน้า (resume ได้)
  data/                → ไฟล์ input (company_ids.xlsx)
  notify/              → ส่ง email ผ่าน Power Automate
  output/              → ผลลัพธ์ Z-Score แยกรายบริษัท
  reports/             → สร้างไฟล์ Excel สรุปผล
  scraper/             → ดึงตารางงบการเงินจาก DBD
  secrets/             → ไฟล์ .env และ .env.key (ไม่ถูก commit)
  utils/               → utility functions ทั่วไป
  zscore/              → คำนวณ Altman Z-Score


----------------------------------------------------------------
  วิธีติดตั้ง (ครั้งแรก)
----------------------------------------------------------------

1. ติดตั้ง dependencies:
   pip install playwright pandas openpyxl python-dotenv cryptography requests
   playwright install chromium

2. เตรียม input file:
   วางไฟล์ company_ids.xlsx ไว้ที่ data/company_ids.xlsx
   → sheet ชื่อ "ids"
   → ต้องมี column: company_id (13 หลัก), CompanyName

3. ตั้งค่า secrets:
   - copy secrets/.env.example → secrets/.env
   - ใส่ค่า POWER_AUTOMATE_WEBHOOK จริงลงใน secrets/.env
   - รัน: python setup_secrets.py
   (จะสร้าง secrets/.env.key และ encrypt ค่าใน .env อัตโนมัติ)

4. ปรับ config ตามต้องการใน config.py:
   - TEST_MODE = True/False
   - TEST_LIMIT = จำนวนบริษัทที่จะเทส
   - RUN_START_TIME / RUN_END_TIME = ช่วงเวลาที่อนุญาตให้รัน
   - HEADLESS = True (ไม่เปิดหน้าต่าง browser) / False (เปิด)


----------------------------------------------------------------
  วิธีรัน
----------------------------------------------------------------

  รันทั้งคู่ทีละบริษัท Capital → Z-Score (default):
    python main.py
    python main.py --mode full

  Z-Score batch อย่างเดียว:
    python main.py --mode zscore

  Registered Capital scraper อย่างเดียว:
    python main.py --mode capital


----------------------------------------------------------------
  Mode 1 — Z-Score Batch  (python main.py)
----------------------------------------------------------------

Input:
  data/company_ids.xlsx

กระบวนการ:
  1. โหลดรายชื่อบริษัทจาก company_ids.xlsx
  2. เปิด browser → ค้นหาแต่ละบริษัทบน DBD
  3. ดึงงบแสดงฐานะการเงิน + งบกำไรขาดทุน
  4. คำนวณ Altman Z-Score
  5. บันทึกผลลัพธ์ + แจ้ง Outlook เมื่อเสร็จ

Output:
  output/zscore.xlsx                          → รวมทุกบริษัท
  output/<company_id>/zscore_<id>.xlsx        → รายบริษัท
  output/<company_id>/balance_sheet_<id>.xlsx → งบดุล
  output/<company_id>/income_statement_<id>.xlsx → งบกำไรขาดทุน
  output/run_summary.json                     → สรุปผลการรัน
  output/error_companies.xlsx                 → รายการที่ error


----------------------------------------------------------------
  Mode 2 — Registered Capital  (python main.py --mode capital)
----------------------------------------------------------------

Input:
  data/company_ids.xlsx (ชุดเดียวกัน)

กระบวนการ:
  ค้นหาแต่ละบริษัทบน DBD และดึงข้อมูลทุนจดทะเบียน

Output:
  N-AUTHORIZED-CAPITAL/Registered capital.xlsx
  → columns: company_id, Company Name, Registered Capital,
             Paid-up Capital, Corporate Status


----------------------------------------------------------------
  Resume / Checkpoint
----------------------------------------------------------------

ทั้ง 2 mode รองรับการ resume:
  - Z-Score: บริษัทที่ทำเสร็จแล้วถูกบันทึกใน checkpoint/
    รันซ้ำจะ skip บริษัทที่ทำไปแล้วอัตโนมัติ
  - Capital: ตรวจสอบจาก Registered capital.xlsx ว่า id ไหนมีแล้ว

ถ้าต้องการรันใหม่ทั้งหมด:
  - ลบไฟล์ใน checkpoint/
  - ลบไฟล์ output ที่ต้องการรันใหม่


----------------------------------------------------------------
  Notification (Outlook via Power Automate)
----------------------------------------------------------------

ระบบส่ง email แจ้งเมื่อ:
  - รันเสร็จสมบูรณ์ (COMPLETED / TEST)
  - รันเสร็จแต่มี error บางส่วน (COMPLETED_WITH_ERRORS)
  - หยุดเพราะนอกช่วงเวลา (PAUSED_TIME_WINDOW)
  - script crash (CRASHED)

แก้รายชื่อผู้รับ email ได้ที่ฝั่ง Power Automate โดยตรง
(ไม่ต้องแก้ code)


----------------------------------------------------------------
  Secrets & Security
----------------------------------------------------------------

ไฟล์ใน secrets/ ไม่ถูก commit ขึ้น Git:
  secrets/.env      → ค่า config ที่ sensitive (encrypt แล้ว)
  secrets/.env.key  → Fernet key สำหรับ decrypt

⚠️  ถ้าย้ายไปเครื่องใหม่:
  ต้อง copy ทั้ง secrets/.env และ secrets/.env.key ไปด้วย

⚠️  ถ้า key หาย:
  ต้องตั้งค่าใหม่ใน secrets/.env แล้วรัน python setup_secrets.py


----------------------------------------------------------------
  ปรับแต่ง config.py
----------------------------------------------------------------

  BASE_URL              → URL ของ DBD (ไม่ควรเปลี่ยน)
  HEADLESS              → True = ไม่เปิดหน้าต่าง browser
  TEST_MODE             → True = รันแค่ TEST_LIMIT บริษัท
  TEST_LIMIT            → จำนวนบริษัทในโหมดเทส
  RUN_START_TIME        → เวลาเริ่มอนุญาตให้รัน (default 08:00)
  RUN_END_TIME          → เวลาสิ้นสุด (default 21:30)
  DEFAULT_TIMEOUT       → timeout ของ Playwright (ms)
  MAX_RETRY             → จำนวนครั้ง retry ต่อบริษัท
  WATCHDOG_RESTART_EVERY → restart browser ทุก N บริษัท
  COMPANY_HARD_TIMEOUT  → timeout สูงสุดต่อบริษัท (วินาที)
