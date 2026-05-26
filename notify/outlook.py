from config import OUTPUT_DIR, Z_SCORE_FILE, POWER_AUTOMATE_WEBHOOK
from utils.zip_utils import zip_and_base64
from pathlib import Path
from utils.logging_utils import ts, log
from checkpoint.state import ERROR_FILE
import requests

def build_attachments():
    zip_path = OUTPUT_DIR / "zscore_reports.zip"

    files = []
    if Z_SCORE_FILE.exists():
        files.append(Z_SCORE_FILE)

    error_file = Path(ERROR_FILE)
    if error_file.exists():
        files.append(error_file)

    if not files:
        return []

    zip_b64 = zip_and_base64(zip_path, files)

    return [{
        "name": "zscore_reports.zip",
        "contentBytes": zip_b64
    }]

def notify_outlook_power_automate(
    *,
    subject: str,
    status: str,
    payload_extra: dict | None = None
):
    """
    ✅ ส่ง Outlook ผ่าน Power Automate (HTTP Webhook)
    """

    # ✅ แนบไฟล์เฉพาะ COMPLETED และ TEST
    if status in ["COMPLETED", "COMPLETED_WITH_ERRORS", "TEST", "PAUSED_TIME_WINDOW"]:
        attachments = build_attachments()
    else:
        attachments = []

    payload = {
        "subject": subject,
        "status": status,
        "timestamp": ts(),
        "attachments": attachments,
        "has_attachments": bool(attachments),
    }

    if payload_extra:
        payload.update(payload_extra)

    try:
        resp = requests.post(
            POWER_AUTOMATE_WEBHOOK,
            json=payload,
            timeout=30
        )
        log(f"📨 Outlook response: {resp.status_code}")
        if resp.text:
            log(f"📨 Outlook response body: {resp.text}")

    except Exception as e:
        log(f"❌ Power Automate notify failed: {e}")