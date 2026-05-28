from pathlib import Path
from datetime import time
from dotenv import load_dotenv
import os

load_dotenv()

POWER_AUTOMATE_WEBHOOK = os.getenv("POWER_AUTOMATE_WEBHOOK")

BASE_DIR = Path(__file__).resolve().parent

# ===== INPUT =====
INPUT_IDS_FILE = BASE_DIR / "data" / "company_ids.xlsx"
IDS_SHEET = "ids"
ID_COL = "company_id"

# ===== WEB =====
BASE_URL = "https://datawarehouse.dbd.go.th/"
DEFAULT_TIMEOUT = 30000
WATCHDOG_RESTART_EVERY = 40
COMPANY_HARD_TIMEOUT = 3 * 60

# ===== OUTPUT =====
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

Z_SCORE_FILE = OUTPUT_DIR / "zscore.xlsx"
Z_SCORE_SHEET = "Z_SCORE"

# ===== RUNTIME =====
HEADLESS = False
MAX_RETRY = 2

RUN_START_TIME = time(14, 0)
RUN_END_TIME   = time(21, 30)
