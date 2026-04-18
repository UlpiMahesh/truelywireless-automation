from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_FILE = BASE_DIR / f"data/allocation_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.xlsx"
print(OUTPUT_FILE)