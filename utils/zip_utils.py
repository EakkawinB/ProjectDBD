import zipfile, base64
import time
from pathlib import Path

def zip_and_base64(zip_path: Path, files: list[Path]):
    if zip_path.exists():
        zip_path.unlink()

    time.sleep(1)

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for f in files:
            if not f.exists():
                raise FileNotFoundError(f"Missing file: {f}")
            if f.stat().st_size <= 0:
                raise RuntimeError(f"File not ready / empty: {f}")
            z.write(f, arcname=f.name)

    # ตรวจว่า ZIP valid จริง
    if not zipfile.is_zipfile(zip_path):
        raise RuntimeError("ZIP corrupted before sending")

    with open(zip_path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")