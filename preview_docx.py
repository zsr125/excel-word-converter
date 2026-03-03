#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 文档浏览器预览：将 .docx 转为 HTML 并在默认浏览器中打开。
"""

import sys
import webbrowser
from pathlib import Path

try:
    import mammoth
except ImportError:
    print("正在安装 mammoth...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "mammoth", "-q"])
    import mammoth


def preview_docx(docx_path: str) -> None:
    """将 docx 转为 HTML 并在浏览器中打开预览。"""
    path = Path(docx_path).resolve()
    if not path.exists():
        raise FileNotFoundError(f"文件不存在: {path}")

    with open(path, "rb") as f:
        result = mammoth.convert_to_html(f)
    body_html = result.value

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>预览 - {path.name}</title>
<style>
  body {{ font-family: "Songti SC", "SimSun", "宋体", serif; margin: 2rem; max-width: 900px; line-height: 1.6; }}
  h1 {{ font-size: 1.4rem; margin-top: 1.5rem; border-bottom: 1px solid #ccc; padding-bottom: 0.3rem; }}
  table {{ border-collapse: collapse; width: 100%; margin: 1rem 0; font-size: 14px; }}
  td, th {{ border: 1px solid #333; padding: 6px 10px; text-align: left; vertical-align: top; }}
  th {{ background: #f0f0f0; font-weight: 600; }}
  p {{ margin: 0.5rem 0; }}
</style>
</head>
<body>
{body_html}
</body>
</html>"""

    out_path = path.with_suffix(".preview.html")
    out_path.write_text(html, encoding="utf-8")
    url = f"file://{out_path}"
    webbrowser.open(url)
    print(f"已在浏览器中打开预览: {out_path}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python preview_docx.py <文档.docx>")
        sys.exit(1)
    preview_docx(sys.argv[1])
