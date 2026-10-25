# AI お絵描き — サーバー公開用イメージ
FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    APP_MODE=server \
    DATA_DIR=/data \
    PORT=8080

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY desktopaap.py kiosk.py index.html admin.html login.html ./
COPY assets ./assets
COPY BGM ./BGM

RUN mkdir -p /data
VOLUME ["/data"]
EXPOSE 8080

# 設定・ログを1つのプロセスで持つので worker は1つ（同時アクセスはスレッドで処理）
# 絵の生成は最大90秒かかるのでタイムアウトは長め
CMD exec gunicorn --bind 0.0.0.0:${PORT} --workers 1 --threads 8 --timeout 180 desktopaap:app
