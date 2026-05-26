from datetime import datetime


def ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def log(msg: str) -> None:
    """INFO level — ใช้ทั่วไป"""
    print(f"[{ts()}] [INFO]  {msg}", flush=True)


def log_warn(msg: str) -> None:
    """WARN level — สิ่งที่ผิดปกติแต่ยังไม่ crash"""
    print(f"[{ts()}] [WARN]  {msg}", flush=True)


def log_error(msg: str) -> None:
    """ERROR level — error ที่ต้องสังเกต"""
    print(f"[{ts()}] [ERROR] {msg}", flush=True)


def log_debug(msg: str) -> None:
    """DEBUG level — รายละเอียดเพิ่มเติม"""
    print(f"[{ts()}] [DEBUG] {msg}", flush=True)


def log_section(title: str) -> None:
    """พิมพ์หัวข้อ section ให้อ่านง่ายขึ้น"""
    bar = "─" * 52
    print(f"\n[{ts()}] ┌{bar}┐", flush=True)
    print(f"[{ts()}] │  {title:<50}│", flush=True)
    print(f"[{ts()}] └{bar}┘", flush=True)


def log_company_start(idx: int, total: int, cid: str, cname: str) -> None:
    """แสดงหัวบริษัทที่กำลังประมวลผล"""
    print(f"\n[{ts()}] {'='*56}", flush=True)
    print(f"[{ts()}] ❤️  [{idx}/{total}] {cname}", flush=True)
    print(f"[{ts()}]     ID : {cid}", flush=True)
    print(f"[{ts()}] {'='*56}", flush=True)


def log_step(step: int, label: str) -> None:
    """แสดงขั้นตอนย่อยภายในบริษัท"""
    print(f"[{ts()}]   ▶ Step {step}: {label}", flush=True)


def log_ok(msg: str) -> None:
    """สำเร็จ"""
    print(f"[{ts()}]   ✅ {msg}", flush=True)


def log_skip(msg: str) -> None:
    """ข้ามเพราะมีข้อมูลแล้ว"""
    print(f"[{ts()}]   ⏭  {msg}", flush=True)


def log_retry(attempt: int, max_retry: int, reason: str) -> None:
    """กำลัง retry"""
    print(f"[{ts()}]   🔄 Retry {attempt}/{max_retry} — {reason}", flush=True)


def log_summary(summary: dict) -> None:
    """พิมพ์สรุปผลการรันทั้งหมด"""
    bar = "─" * 52
    print(f"\n[{ts()}] ┌{bar}┐", flush=True)
    print(f"[{ts()}] │  {'RUN SUMMARY':<50}│", flush=True)
    print(f"[{ts()}] ├{bar}┤", flush=True)
    for k, v in summary.items():
        if isinstance(v, dict):
            print(f"[{ts()}] │  {k:<20} {str(v):<30}│", flush=True)
        else:
            print(f"[{ts()}] │  {k:<20} {str(v):<30}│", flush=True)
    print(f"[{ts()}] └{bar}┘\n", flush=True)
