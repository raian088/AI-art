#!/bin/bash
# AI お絵描き — Mac 用起動スクリプト（ダブルクリックで起動）
cd "$(dirname "$0")/.." || exit 1

if ! command -v python3 >/dev/null 2>&1; then
  echo "[エラー] python3 が見つかりません。https://www.python.org/ からインストールしてください。"
  read -r -p "Enterで閉じる"; exit 1
fi
if [ ! -f .env ]; then
  cp .env.example .env
  echo "[注意] .env を作成しました。GEMINI_API_KEY= にAPIキーを書いて保存し、もう一度起動してください。"
  open -e .env
  read -r -p "Enterで閉じる"; exit 1
fi

# 初回だけ仮想環境を作ってライブラリを入れる
if [ ! -d .venv ]; then
  echo "初回セットアップ中（ライブラリをインストールします）..."
  python3 -m venv .venv || exit 1
  ./.venv/bin/pip install -q flask flask-cors requests python-dotenv pillow openpyxl || exit 1
fi

echo "AI お絵描きを起動します（終了: 管理画面の「AIソフトを閉じる」またはこの画面で Ctrl+C）"
exec ./.venv/bin/python desktopaap.py --browser
