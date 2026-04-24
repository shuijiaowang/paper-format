from __future__ import annotations

import base64
from pathlib import Path
from tempfile import TemporaryDirectory

from flask import Flask, jsonify, request, send_file
from werkzeug.utils import secure_filename

from format_docs import OUTPUT_SUFFIX, apply_mvp_format, load_config

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"
INDEX_PATH = BASE_DIR / "index.html"

app = Flask(__name__)
config = load_config(CONFIG_PATH)


@app.get("/")
def index():
    return send_file(INDEX_PATH)


@app.post("/api/format")
def format_docx():
    upload = request.files.get("file")
    if upload is None or not upload.filename:
        return jsonify({"error": "请先上传 .docx 文件"}), 400

    raw_name = upload.filename
    client_name = raw_name.replace("\\", "/").split("/")[-1] or "upload.docx"
    if not client_name.lower().endswith(".docx"):
        return jsonify({"error": "只支持 .docx 文件"}), 400
    safe_name = secure_filename(client_name) or "upload.docx"

    with TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        src = tmp_path / safe_name
        out_name = f"{Path(client_name).stem}{OUTPUT_SUFFIX}.docx"
        out = tmp_path / out_name

        upload.save(src)
        report = apply_mvp_format(src, out, config)
        output_bytes = out.read_bytes()
        encoded_file = base64.b64encode(output_bytes).decode("ascii")
        return jsonify(
            {
                "download_name": out_name,
                "file_base64": encoded_file,
                "report": report.to_dict(),
            }
        )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=9001)
