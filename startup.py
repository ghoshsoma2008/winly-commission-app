from flask import Flask
import os
import subprocess

app = Flask(__name__)

@app.route('/')
def home():
    subprocess.Popen([
        "streamlit", "run", "commission_calculator_app_FINAL.py",
        "--server.port", "8501",
        "--server.address", "0.0.0.0"
    ])
    return "Streamlit app is starting..."

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port)
