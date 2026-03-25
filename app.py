from scripts import amounts, allocation
from flask import Flask, render_template, request, send_file
import pandas as pd
from pathlib import Path



app = Flask(__name__)
BASE_DIR = Path(__file__).resolve().parent
@app.route("/markets")
def get_markets():
    df = pd.read_excel(BASE_DIR / "data/marketlogins.xlsx")

    # 🔥 ADD THIS LINE
    df.columns = df.columns.str.strip()

    markets = df["Market"].dropna().unique().tolist()
    return {"markets": markets}
@app.route("/")
def home():
    return render_template("index.html")

@app.route("/run/amounts", methods=["POST"])
def run_amounts():
    selected_markets = request.form.getlist("markets")
    file_path = amounts.run(selected_markets)
    return send_file(file_path, as_attachment=True)

@app.route("/run/allocation")
def run_allocation():
    file_path = allocation.run()
    return send_file(file_path, as_attachment=True)

if __name__ == "__main__":
    app.run()