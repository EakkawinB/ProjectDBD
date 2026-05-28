# ProjectDBD

ระบบดึงข้อมูลงบการเงินจาก DBD, คำนวณ Altman Z-Score และสรุปรายงาน Excel พร้อมแจ้งเตือนผ่าน Power Automate

## ความสามารถหลัก

- อ่านรายการบริษัทจากไฟล์ Excel
- เข้าเว็บ DBD อัตโนมัติด้วย Playwright
- ดึง Balance Sheet และ Income Statement
- คำนวณ `Z'-Score` และสถานะความเสี่ยง
- สร้างรายงานรวมและรายบริษัทในโฟลเดอร์ `output`
- แจ้งผลการรันผ่าน webhook (Power Automate / Outlook)

## โครงสร้างโปรเจกต์

- `main.py` จุดเริ่มรันงาน batch หลัก
- `addon.py` งานเสริมสำหรับดึงทุนจดทะเบียน
- `config.py` ค่าคอนฟิกหลักและ environment
- `scraper/` โมดูลดึง/รวมข้อมูลตาราง
- `zscore/` โมดูลคำนวณและจัดรูปแบบ Z-Score
- `reports/` โมดูลสร้างไฟล์รายงาน Excel และ dashboard
- `notify/` โมดูลแจ้งเตือนผ่าน webhook
- `utils/` utility ด้าน text/time/excel/log/zip
- `data/company_ids.xlsx` ไฟล์ input รายชื่อบริษัท
- `output/` ไฟล์ผลลัพธ์ทั้งหมด

## รูปแบบไฟล์ Input

ไฟล์ `data/company_ids.xlsx` ต้องมี sheet ชื่อ `ids` และคอลัมน์อย่างน้อย:

- `company_id` (13 หลัก)
- `CompanyName`

## การติดตั้ง

1. ใช้ Python 3.10+ (แนะนำ 3.11)
2. ติดตั้ง dependencies ที่ใช้งานในโค้ด:
   - `pandas`
   - `numpy`
   - `openpyxl`
   - `playwright`
   - `requests`
   - `python-dotenv`
3. ติดตั้ง browser ของ Playwright

ตัวอย่างคำสั่ง:

```bash
pip install pandas numpy openpyxl playwright requests python-dotenv
playwright install
```

## Environment

กำหนดค่าใน `.env` (หรือ environment variable):

- `POWER_AUTOMATE_WEBHOOK` URL webhook สำหรับส่งแจ้งเตือน

## วิธีรัน

รัน batch หลัก:

```bash
python main.py
```

รันงานดึงทุนจดทะเบียน:

```bash
python addon.py
```

## Output

- `output/zscore.xlsx` รายงานรวม
- `output/<company_id>/zscore_<company_id>.xlsx` รายงานรายบริษัท
- `output/zscore_reports.zip` ไฟล์แนบสำหรับการแจ้งเตือน (สร้างเมื่อมีการส่งแจ้งเตือน)

## หมายเหตุ

- ระบบมี time window การรันใน `main.py` (`RUN_START_TIME`, `RUN_END_TIME`)
- ถ้าไฟล์ input ไม่พบ ระบบจะหยุดพร้อมแจ้ง path ที่คาดหวัง
- โค้ดถูกแยกโมดูลแล้ว แต่ `main.py` ยังคงเป็นตัว orchestrate หลัก
