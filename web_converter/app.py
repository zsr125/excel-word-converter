#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel/Numbers → Word 产品技术手册 - 在线转换服务
支持 .xlsx 与 .numbers，单文件上限 100MB
"""

import os
import sys
import tempfile
import uuid
from pathlib import Path

# 导入本地转换器（需将上一级目录加入路径）
_sys_path = str(Path(__file__).resolve().parent.parent)
if _sys_path not in sys.path:
    sys.path.insert(0, _sys_path)

from flask import Flask, request, send_file, jsonify
from excel_to_word_converter import convert_to_word

app = Flask(__name__, static_folder="static", static_url_path="")
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100MB
ALLOWED_EXTENSIONS = {".xlsx", ".numbers"}


def allowed_file(filename):
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


@app.route("/")
def index():
    return send_file("static/index.html")


@app.route("/api/convert", methods=["POST"])
def convert():
    if "file" not in request.files:
        return jsonify({"ok": False, "msg": "请选择要上传的文件"}), 400
    f = request.files["file"]
    if not f or not f.filename:
        return jsonify({"ok": False, "msg": "未检测到文件"}), 400
    if not allowed_file(f.filename):
        return jsonify({"ok": False, "msg": "仅支持 .xlsx 和 .numbers 格式"}), 400

    tmp_dir = tempfile.mkdtemp()
    try:
        uid = str(uuid.uuid4())[:8]
        ext = Path(f.filename).suffix.lower()
        inp = os.path.join(tmp_dir, f"in_{uid}{ext}")
        out = os.path.join(tmp_dir, f"out_{uid}.docx")
        f.save(inp)
        convert_to_word(inp, out)
        return send_file(
            out,
            as_attachment=True,
            download_name=Path(f.filename).stem + ".docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500
    finally:
        for p in Path(tmp_dir).iterdir():
            try:
                p.unlink()
            except Exception:
                pass
        try:
            os.rmdir(tmp_dir)
        except Exception:
            pass


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
