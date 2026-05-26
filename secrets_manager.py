"""
secrets_manager.py
------------------
จัดการ encrypt / decrypt ค่า secrets ด้วย Fernet symmetric encryption.

โครงสร้างไฟล์:
  secrets/.env.key   → Fernet key (ห้าม commit ขึ้น Git เด็ดขาด)
  secrets/.env       → ค่าที่ encrypt แล้ว  (ห้าม commit เช่นกัน)

วิธีใช้ครั้งแรก:
  python setup_secrets.py

วิธีใช้ใน code:
  from secrets_manager import get_secret
  webhook = get_secret("POWER_AUTOMATE_WEBHOOK")
"""

from pathlib import Path
from cryptography.fernet import Fernet, InvalidToken

BASE_DIR = Path(__file__).resolve().parent
SECRETS_DIR = BASE_DIR / "secrets"
KEY_FILE = SECRETS_DIR / ".env.key"
ENV_FILE = SECRETS_DIR / ".env"

# prefix ที่ใช้แยกว่าค่าไหน encrypt แล้ว
ENCRYPTED_PREFIX = "enc::"


def generate_key() -> bytes:
    """สร้าง Fernet key ใหม่"""
    return Fernet.generate_key()


def save_key(key: bytes) -> None:
    """บันทึก key ลงไฟล์ secrets/.env.key"""
    SECRETS_DIR.mkdir(exist_ok=True)
    KEY_FILE.write_bytes(key)
    print(f"[secrets] Key saved → {KEY_FILE}")


def load_key() -> bytes:
    """โหลด key จาก secrets/.env.key"""
    if not KEY_FILE.exists():
        raise FileNotFoundError(
            f"ไม่พบ key file: {KEY_FILE}\n"
            "รัน `python setup_secrets.py` เพื่อสร้าง key ก่อน"
        )
    return KEY_FILE.read_bytes().strip()


def encrypt_value(plaintext: str) -> str:
    """
    Encrypt ค่า string แล้วคืนเป็น  enc::<base64_ciphertext>
    """
    key = load_key()
    f = Fernet(key)
    token = f.encrypt(plaintext.encode()).decode()
    return f"{ENCRYPTED_PREFIX}{token}"


def decrypt_value(encrypted: str) -> str:
    """
    Decrypt ค่าที่มี prefix  enc::
    ถ้าค่าไม่ได้ encrypt (ไม่มี prefix) จะคืนค่าเดิม
    """
    if not encrypted.startswith(ENCRYPTED_PREFIX):
        return encrypted  # ค่าธรรมดา ไม่ต้อง decrypt

    token = encrypted[len(ENCRYPTED_PREFIX):]
    key = load_key()
    f = Fernet(key)
    try:
        return f.decrypt(token.encode()).decode()
    except InvalidToken:
        raise ValueError(
            "Decrypt ล้มเหลว: key ไม่ตรงกับค่าที่ encrypt ไว้\n"
            "ตรวจสอบว่าใช้ secrets/.env.key ไฟล์เดิม"
        )


def get_secret(key_name: str, default: str = "") -> str:
    """
    อ่านค่าจาก environment แล้ว decrypt อัตโนมัติถ้าจำเป็น
    ใช้แทน os.getenv() สำหรับค่าที่ sensitive
    """
    import os
    raw = os.getenv(key_name, default)
    if not raw:
        return default
    return decrypt_value(raw)


def encrypt_env_file() -> None:
    """
    อ่าน secrets/.env ทุก key=value แล้ว encrypt ค่าที่ยังไม่ได้ encrypt
    เขียนกลับลง secrets/.env เดิม
    """
    if not ENV_FILE.exists():
        print(f"[secrets] ไม่พบ {ENV_FILE}")
        return

    lines = ENV_FILE.read_text(encoding="utf-8").splitlines()
    new_lines = []
    changed = 0

    for line in lines:
        stripped = line.strip()
        # ข้ามบรรทัดว่างและ comment
        if not stripped or stripped.startswith("#"):
            new_lines.append(line)
            continue

        if "=" in stripped:
            key, _, value = stripped.partition("=")
            key = key.strip()
            value = value.strip()

            if value and not value.startswith(ENCRYPTED_PREFIX):
                encrypted = encrypt_value(value)
                new_lines.append(f"{key}={encrypted}")
                print(f"[secrets] ✅ Encrypted: {key}")
                changed += 1
            else:
                new_lines.append(line)
        else:
            new_lines.append(line)

    ENV_FILE.write_text("\n".join(new_lines) + "\n", encoding="utf-8")
    print(f"[secrets] เสร็จแล้ว — encrypt {changed} ค่า")
