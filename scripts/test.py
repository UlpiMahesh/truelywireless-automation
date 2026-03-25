from pathlib import Path
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent.parent   # 👈 go up one level
OUTPUT_FILE = BASE_DIR / "data" / "outputs" / "allocation_all_markets.xlsx"

print(OUTPUT_FILE)
file = pd.read_excel(OUTPUT_FILE)
print(file)