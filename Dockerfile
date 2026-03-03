# 放在仓库根目录，确保能访问 excel_to_word_converter.py
FROM python:3.11-slim

WORKDIR /app

COPY excel_to_word_converter.py /app/
COPY web_converter/requirements.txt /app/
COPY web_converter/app.py /app/
COPY web_converter/static /app/static/

RUN pip install --no-cache-dir -r requirements.txt

EXPOSE 5000

ENV PORT=5000
CMD sh -c "python3 -m gunicorn -w 2 -b 0.0.0.0:\${PORT:-5000} --timeout 300 app:app"
