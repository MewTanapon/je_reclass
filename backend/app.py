"""
Flask API server — serves the frontend and exposes /api/run.
"""

import os
import sys

from flask import Flask, jsonify, request, send_from_directory

# Allow running directly from backend/ or from the project root
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from processor import process_reclass

FRONTEND_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "frontend")

app = Flask(__name__, static_folder=FRONTEND_DIR, static_url_path="")


# ── Static frontend ───────────────────────────────────────────────────────────

@app.route("/")
def index():
    return send_from_directory(FRONTEND_DIR, "index.html")


# ── API ───────────────────────────────────────────────────────────────────────

@app.route("/api/health")
def health():
    return jsonify({"status": "ok"})


@app.route("/api/run", methods=["POST"])
def run_reclass():
    body = request.get_json(silent=True) or {}
    input_folder  = (body.get("input_folder")  or "").strip()
    output_folder = (body.get("output_folder") or "").strip()

    if not input_folder or not output_folder:
        return jsonify({"error": "Both input_folder and output_folder are required."}), 400

    if not os.path.isdir(input_folder):
        return jsonify({"error": f"Input folder not found: {input_folder}"}), 400

    try:
        output_path, output_filename = process_reclass(input_folder, output_folder)
        return jsonify({
            "success":         True,
            "output_filename": output_filename,
            "output_path":     output_path,
        })
    except FileNotFoundError as exc:
        return jsonify({"error": str(exc)}), 404
    except KeyError as exc:
        return jsonify({"error": f"Missing expected column in Excel: {exc}"}), 400
    except Exception as exc:  # noqa: BLE001
        return jsonify({"error": str(exc)}), 500


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("Starting Reclass JE Generator...")
    print("Open http://localhost:5000 in your browser.")
    app.run(debug=True, port=5000)
