from flask import Flask, render_template, send_file
from scripts import amounts, allocation

app = Flask(__name__)

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/run/amounts")
def run_amounts():
    file_path = amounts.run()
    return send_file(file_path, as_attachment=True)

@app.route("/run/allocation")
def run_allocation():
    file_path = allocation.run()
    return send_file(file_path, as_attachment=True)

if __name__ == "__main__":
    app.run()