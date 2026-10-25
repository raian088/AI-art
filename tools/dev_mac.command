#!/bin/bash
# AI お絵描き — Mac での開発・動作確認用（ダブルクリックで起動）
#   サーバー版と同じ動きを Mac 上で確認できます: http://localhost:8080/
#   修正したら Ctrl+C で止めて、もう一度ダブルクリック
cd "$(dirname "$0")/.." || exit 1

if ! command -v python3 >/dev/null 2>&1; then
  echo "[エラー] python3 が見つかりません。https://www.python.org/ からインストールしてください。"
  read -r -p "Enterで閉じる"; exit 1
fi

# 初回だけ仮想環境を作ってライブラリを入れる（requirements.txt が変わったら入れ直す）
if [ ! -d .venv ] || [ requirements.txt -nt .venv/.installed ]; then
  echo "ライブラリをインストールしています..."
  [ -d .venv ] || python3 -m venv .venv || exit 1
  ./.venv/bin/pip install -q -r requirements.txt && touch .venv/.installed || exit 1
fi

# 開発用データ（.env・ログ）は data/ に置く（Git には入りません）
mkdir -p data
[ -f data/.env ] || cp .env.example data/.env

export APP_MODE=server DATA_DIR="$PWD/data" HOST=127.0.0.1 PORT=8080 COOKIE_SECURE=0
( sleep 2; open "http://localhost:8080/" ) &
echo "AI お絵描き（開発用）を起動します。止めるときは Ctrl+C"
exec ./.venv/bin/python desktopaap.py --server
