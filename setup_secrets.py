"""
setup_secrets.py
----------------
รัน script นี้ครั้งเดียวเพื่อ:
  1. สร้าง secrets/.env.key (Fernet key)
  2. Encrypt ค่าทุกตัวใน secrets/.env

คำสั่ง:
  python setup_secrets.py

⚠️  สำคัญ:
  - เก็บ secrets/.env.key ไว้ในที่ปลอดภัย (ห้าม commit ขึ้น Git)
  - ถ้าทำงานหลายเครื่อง ต้อง copy secrets/.env.key ไปด้วยทุกครั้ง
  - ถ้า key หาย จะ decrypt ค่าเดิมไม่ได้ → ต้องตั้งค่าใหม่ใน secrets/.env
"""

from secrets_manager import (
    KEY_FILE,
    ENV_FILE,
    generate_key,
    save_key,
    encrypt_env_file,
)


def main():
    print("=" * 50)
    print("  ProjectDBD — Secrets Setup")
    print("=" * 50)

    # 1. สร้าง key ถ้ายังไม่มี
    if KEY_FILE.exists():
        overwrite = input(
            f"\n⚠️  พบ {KEY_FILE} อยู่แล้ว\n"
            "สร้าง key ใหม่จะทำให้ decrypt ค่าเดิมไม่ได้\n"
            "สร้างใหม่? (y/N): "
        ).strip().lower()
        if overwrite != "y":
            print("[setup] ใช้ key เดิม")
        else:
            key = generate_key()
            save_key(key)
            print("[setup] สร้าง key ใหม่แล้ว")
    else:
        key = generate_key()
        save_key(key)
        print("[setup] สร้าง key ใหม่แล้ว")

    # 2. Encrypt secrets/.env
    if not ENV_FILE.exists():
        print(f"\n[setup] ไม่พบ {ENV_FILE} — ข้ามขั้นตอน encrypt")
    else:
        print(f"\n[setup] กำลัง encrypt ค่าใน {ENV_FILE} ...")
        encrypt_env_file()

    print("\n✅ เสร็จสิ้น")
    print(f"   Key file : {KEY_FILE}")
    print(f"   Env file : {ENV_FILE}")
    print("\n⚠️  ตรวจสอบ .gitignore ว่ามี secrets/ อยู่แล้ว")


if __name__ == "__main__":
    main()
