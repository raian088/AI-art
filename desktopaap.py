import os
import sys
import io
import json
import base64
import tempfile
import traceback
import textwrap
import time
from datetime import datetime
import threading
import subprocess

import requests
import socket
import ctypes
from dotenv import load_dotenv
from flask import Flask, request, jsonify, send_from_directory, session, redirect
from markupsafe import escape
from flask_cors import CORS
from PIL import Image, ImageDraw, ImageFont

import kiosk

# ==========================================
# Excel debug log (debug.xlsx)
# ==========================================
try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    _HAS_OPENPYXL = True
except Exception as _e_xlsx:
    _HAS_OPENPYXL = False
    print(f"[XLSX] openpyxl unavailable; debug.xlsx logging disabled: {_e_xlsx}")

import re as _re_xlsx
_xlsx_lock = threading.Lock()
_xlsx_pending = []
_XLSX_PENDING_MAX = 500
_events_buffer = []
_EVENTS_BUFFER_MAX = 1000
_xlsx_history_loaded = False

_XLSX_COLUMNS = [
    ("日時", 22), ("種別", 18), ("お題", 24), ("所要秒数", 10),
    ("キャンバス幅", 12), ("キャンバス高", 12),
    ("AIコマンド数", 13), ("採用コマンド数", 14), ("除去背景コマンド", 16),
    ("リトライ回数", 12), ("エラー内容", 40), ("画像ファイル", 32),
    ("モデル名", 22), ("入力トークン", 13), ("出力トークン", 13),
    ("思考トークン", 13), ("概算費用(円)", 14), ("備考", 40),
    ("評価★", 12), ("ID", 22),
]
_XLSX_KEYS = [
    "_ts", "event_type", "prompt", "duration_sec", "canvas_w", "canvas_h",
    "cmd_raw", "cmd_filtered", "removed_bg", "retries", "error", "image_file",
    "model", "tokens_in", "tokens_out", "tokens_thought", "cost_jpy", "detail",
    "rating", "event_id",
]

def _xlsx_dir():
    env = (os.environ.get("DEBUG_XLSX_DIR", "") or "").strip().strip('"')
    if env and os.path.isdir(env):
        return env
    return DATA_DIR

def _xlsx_path():
    return os.path.join(_xlsx_dir(), "debug.xlsx")

def _images_dir():
    p = os.path.join(_xlsx_dir(), "logs", "images")
    os.makedirs(p, exist_ok=True)
    return p

def _ensure_xlsx():
    path = _xlsx_path()
    if os.path.exists(path):
        try:
            return load_workbook(path), path
        except Exception as e:
            print(f"[XLSX] existing debug.xlsx read failed; recreating: {e}")
    wb = Workbook()
    ws = wb.active
    ws.title = "events"
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", start_color="FF6B35")
    center = Alignment(horizontal="center", vertical="center")
    for i, (name, width) in enumerate(_XLSX_COLUMNS, start=1):
        cell = ws.cell(row=1, column=i, value=name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A2"
    return wb, path

def _sanitize_filename(s, maxlen=40):
    if not s:
        return "noprompt"
    s = _re_xlsx.sub(r'[\\/:*?"<>|\r\n\t]', "_", str(s))
    s = s.strip().strip(".")
    return (s[:maxlen] if len(s) > maxlen else s) or "noprompt"

def save_drawing_image(b64_dataurl, prompt):
    if not b64_dataurl:
        return ""
    try:
        b64 = b64_dataurl.split(",", 1)[1] if "," in b64_dataurl else b64_dataurl
        raw = base64.b64decode(b64)
        ts = datetime.now().strftime("%Y%m%d-%H%M%S")
        fname = f"{ts}_{_sanitize_filename(prompt)}.png"
        full = os.path.join(_images_dir(), fname)
        with open(full, "wb") as f:
            f.write(raw)
        return os.path.join("logs", "images", fname).replace("\\", "/")
    except Exception as e:
        print(f"[XLSX] image save failed: {e}")
        return ""

def count_commands(json_text):
    if not json_text:
        return None
    try:
        clean = str(json_text).replace("```json", "").replace("```", "").strip()
        arr = json.loads(clean)
        return len(arr) if isinstance(arr, list) else None
    except Exception:
        return None

def estimate_cost_jpy(model_name, in_tok, out_tok, thought_tok=0):
    try:
        m = (model_name or "").lower()
        if "3.1" in m and "pro" in m:
            input_usd_per_m, output_usd_per_m = 2.0, 12.0
        elif "pro" in m:
            input_usd_per_m, output_usd_per_m = 1.25, 10.0
        else:
            input_usd_per_m, output_usd_per_m = 0.30, 2.50
        billed_out = int(out_tok or 0) + int(thought_tok or 0)
        usd = (int(in_tok or 0) * input_usd_per_m + billed_out * output_usd_per_m) / 1_000_000.0
        return round(usd * 155.0, 4)
    except Exception:
        return None

def _write_one_xlsx_row(ws, ev):
    row = ws.max_row + 1
    values = [
        ev.get("_ts"), ev.get("event_type"), ev.get("prompt"),
        round(float(ev["duration_sec"]), 2) if ev.get("duration_sec") is not None else None,
        int(ev["canvas_w"]) if ev.get("canvas_w") is not None else None,
        int(ev["canvas_h"]) if ev.get("canvas_h") is not None else None,
        int(ev["cmd_raw"]) if ev.get("cmd_raw") is not None else None,
        int(ev["cmd_filtered"]) if ev.get("cmd_filtered") is not None else None,
        int(ev["removed_bg"]) if ev.get("removed_bg") is not None else None,
        int(ev["retries"]) if ev.get("retries") is not None else None,
        str(ev.get("error"))[:1000] if ev.get("error") else None,
        ev.get("image_file") or None, ev.get("model") or None,
        int(ev["tokens_in"]) if ev.get("tokens_in") is not None else None,
        int(ev["tokens_out"]) if ev.get("tokens_out") is not None else None,
        int(ev["tokens_thought"]) if ev.get("tokens_thought") is not None else None,
        float(ev["cost_jpy"]) if ev.get("cost_jpy") is not None else None,
        str(ev.get("detail"))[:1000] if ev.get("detail") else None,
        int(ev["rating"]) if ev.get("rating") is not None else None,
        ev.get("event_id") or None,
    ]
    for col, value in enumerate(values, start=1):
        cell = ws.cell(row=row, column=col, value=value)
        if col == 12 and value:
            try:
                cell.hyperlink = value
                cell.font = Font(color="0563C1", underline="single")
            except Exception:
                pass
        if col == 17 and value is not None:
            cell.number_format = '#,##0.0000"円"'

def log_to_xlsx(event_type, prompt="", duration_sec=None, canvas_w=None,
                canvas_h=None, cmd_raw=None, cmd_filtered=None,
                removed_bg=None, retries=None, error="", image_file="",
                detail="", model="", tokens_in=None, tokens_out=None,
                tokens_thought=None, cost_jpy=None):
    if not _HAS_OPENPYXL:
        return
    now = datetime.now()
    ev = {
        "_ts": now.strftime("%Y-%m-%d %H:%M:%S"),
        "event_id": now.strftime("%Y%m%d-%H%M%S-%f"),
        "event_type": str(event_type or "").upper(),
        "prompt": prompt,
        "duration_sec": duration_sec,
        "canvas_w": canvas_w,
        "canvas_h": canvas_h,
        "cmd_raw": cmd_raw,
        "cmd_filtered": cmd_filtered,
        "removed_bg": removed_bg,
        "retries": retries,
        "error": error,
        "image_file": image_file,
        "detail": detail,
        "model": model,
        "tokens_in": tokens_in,
        "tokens_out": tokens_out,
        "tokens_thought": tokens_thought,
        "cost_jpy": cost_jpy,
        "rating": None,
    }
    with _xlsx_lock:
        _events_buffer.append(ev)
        while len(_events_buffer) > _EVENTS_BUFFER_MAX:
            _events_buffer.pop(0)
        batch = list(_xlsx_pending) + [ev]
        try:
            wb, path = _ensure_xlsx()
            ws = wb["events"] if "events" in wb.sheetnames else wb.active
            for one in batch:
                _write_one_xlsx_row(ws, one)
            wb.save(path)
            _xlsx_pending.clear()
        except PermissionError:
            _xlsx_pending.append(ev)
            while len(_xlsx_pending) > _XLSX_PENDING_MAX:
                _xlsx_pending.pop(0)
            print(f"[XLSX] debug.xlsx is open; pending rows={len(_xlsx_pending)}")
        except Exception as e:
            _xlsx_pending.append(ev)
            while len(_xlsx_pending) > _XLSX_PENDING_MAX:
                _xlsx_pending.pop(0)
            print(f"[XLSX] write failed; pending rows={len(_xlsx_pending)}: {e}")

def _load_xlsx_history_into_buffer():
    global _xlsx_history_loaded
    if _xlsx_history_loaded or not _HAS_OPENPYXL or not os.path.exists(_xlsx_path()):
        return
    try:
        wb = load_workbook(_xlsx_path(), read_only=True, data_only=True)
        ws = wb["events"] if "events" in wb.sheetnames else wb.active
        loaded = []
        first = True
        for row in ws.iter_rows(values_only=True):
            if first:
                first = False
                continue
            ev = {}
            for i, key in enumerate(_XLSX_KEYS):
                if i < len(row) and row[i] is not None:
                    ev[key] = row[i]
            if ev:
                loaded.append(ev)
        wb.close()
        with _xlsx_lock:
            _events_buffer.clear()
            _events_buffer.extend(loaded[-_EVENTS_BUFFER_MAX:])
        _xlsx_history_loaded = True
    except Exception as e:
        print(f"[XLSX] history load failed: {e}")

# ==========================================
# ★多重起動防止（Windowsミューテックス + ポートチェック）
# ==========================================
# 1) 名前付きミューテックスで「すでに起動しているか」を検出
#    名前は他のアプリと衝突しないユニークな文字列にする
APP_MUTEX_NAME = "Local\\AIDrawing-Mutex-7c4f1e2a-8b6d-4a91-9d3e-5f0a7b2c1d4e"
APP_WINDOW_TITLE = "AI お絵描き"  # 既存ウィンドウを前面に出すときに使う
ERROR_ALREADY_EXISTS = 183
_mutex_handle = None  # GC で閉じられないようにグローバル保持

def _bring_existing_window_to_front():
    """すでに起動中の AI お絵描きウィンドウを探して前面に出す。"""
    try:
        user32 = ctypes.windll.user32
        # FindWindowW(クラス名, ウィンドウタイトル)
        hwnd = user32.FindWindowW(None, APP_WINDOW_TITLE)
        if hwnd:
            # 最小化されていれば復元
            SW_RESTORE = 9
            user32.ShowWindow(hwnd, SW_RESTORE)
            user32.SetForegroundWindow(hwnd)
            return True
    except Exception:
        pass
    return False

def _show_already_running_message():
    """すでに起動中の旨をメッセージボックスで知らせる。"""
    try:
        MB_ICONINFORMATION = 0x40
        ctypes.windll.user32.MessageBoxW(
            0,
            "AI お絵描きはすでに起動しています。\n"
            "タスクバーかタスクトレイから既存のウィンドウを開いてください。",
            "AI お絵描き",
            MB_ICONINFORMATION,
        )
    except Exception:
        pass

def ensure_single_instance():
    """二重起動を検出したら既存ウィンドウを前面に出して終了する。"""
    global _mutex_handle
    try:
        kernel32 = ctypes.windll.kernel32
        # CreateMutexW(NULL, FALSE, name)
        _mutex_handle = kernel32.CreateMutexW(None, False, APP_MUTEX_NAME)
        last_err = kernel32.GetLastError()
        if last_err == ERROR_ALREADY_EXISTS:
            # すでに起動中
            _bring_existing_window_to_front()
            _show_already_running_message()
            sys.exit(0)
    except Exception as e:
        # ミューテックスが使えない環境ではポートチェックにフォールバック
        print(f"[WARN] mutex check failed: {e}")

def is_port_in_use(host: str, port: int) -> bool:
    """指定ポートが既に使われているか確認（フォールバック用）。"""
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(0.5)
    try:
        s.connect((host, port))
        return True
    except OSError:
        return False
    finally:
        s.close()

# ==========================================
# ★パス解決（EXE化・通常実行 両対応）
# ==========================================
def get_base_path():
    if getattr(sys, 'frozen', False):
        # EXE化（PyInstaller）されて実行されている場合の一時フォルダ
        return sys._MEIPASS
    # 通常のpythonコマンドで実行されている場合
    return os.path.dirname(os.path.abspath(__file__))

base_dir = get_base_path()
# .envやログファイルは、EXE本体と同じ場所を読み書きするように設定
exe_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else base_dir

# EXE化して --windowed で起動された場合、stdout/stderr が None になって print で落ちることがある。
# ここでログファイルにリダイレクトし、同時にクラッシュを防ぐ。
if getattr(sys, 'frozen', False):
    _log_dir = os.path.join(exe_dir, "logs")
    os.makedirs(_log_dir, exist_ok=True)
    _log_path = os.path.join(_log_dir, "app.log")
    try:
        _fp = open(_log_path, "a", encoding="utf-8", buffering=1)
        sys.stdout = _fp
        sys.stderr = _fp
    except Exception:
        # 書き込めない場合は黒穴へ
        sys.stdout = open(os.devnull, "w")
        sys.stderr = open(os.devnull, "w")

# ==========================================
# 動作モード
#   desktop : PC 上で exe / python を起動し、その PC の Chrome で開く（従来・オフライン向け）
#   server  : サーバー（クラウド）で公開し、Windows PC は Chrome でドメインを開くだけ
#             APP_MODE=server（Docker / gunicorn）か、起動引数 --server
# ==========================================
SERVER_MODE = ("--server" in [a.lower() for a in sys.argv[1:]]) or \
    (os.environ.get("APP_MODE", "").strip().lower() == "server")

# データ（.env・ログ・debug.xlsx・描いた絵）の保存先。サーバーでは永続ディスクを DATA_DIR に指定する
DATA_DIR = (os.environ.get("DATA_DIR", "") or "").strip() or \
    (os.path.join(base_dir, "data") if SERVER_MODE else exe_dir)
os.makedirs(DATA_DIR, exist_ok=True)

# .env ファイルの読み込み（サーバーでは管理画面で保存した値を、ホスティング側の環境変数より優先）
load_dotenv(os.path.join(DATA_DIR, ".env"), override=SERVER_MODE)
API_KEY = os.environ.get("GEMINI_API_KEY", "")

def _gemini_headers():
    """APIキーはURLではなくヘッダーで送る（URLに入れるとエラーログにキーが残るため）。"""
    return {"Content-Type": "application/json", "x-goog-api-key": API_KEY}
GEMINI_MODELS = [
    # --- Gemini 3 系 ---
    "gemini-3.1-pro-preview",
    # --- Gemini 2.5 系 ---
    "gemini-2.5-pro",
    "gemini-2.5-flash",
    "gemini-2.5-flash-lite",
    "gemini-3.5-flash",
    "gemini-3-flash-preview",
    "gemini-3.1-flash-lite",
    "gemini-3.1-flash-lite-preview",
    "gemini-3-pro-preview",
    # --- Gemini 2.0 系 ---
    "gemini-2.0-flash",
    "gemini-2.0-flash-lite",
    # --- エイリアス ---
    "gemini-flash-latest",
]

# --- お題の難しさに応じた思考トークンの割り当て ---
THINK_MIN = 800    # かんたん　：単語1つ（りんご・ねこ）
THINK_MID = 1100   # ふつう　　：修飾つき（かわいい ねこ・ゆうやけのふじさん）
THINK_MAX = 1500   # むずかしい：場面や複数の要素（そらをとぶいぬとねこ）

def _thinking_budget_for(prompt: str) -> int:
    """お題の複雑さを見て思考トークン数を決める。
    単語1つだけなら少なく、場面や複数の要素が入るほど多く考えさせる。"""
    import re
    t = (prompt or "").strip()
    core = re.sub(r"[ 　]", "", t)
    if not core:
        return THINK_MIN

    score = 0
    # ① 文字数（情報量の目安）
    n = len(core)
    if n >= 6:  score += 1
    if n >= 10: score += 1
    if n >= 15: score += 1
    if n >= 20: score += 1
    # ② 単語数（スペース区切り）「あかい りんご」など
    words = [w for w in re.split(r"[ 　]+", t) if w]
    if len(words) >= 2: score += 1
    if len(words) >= 3: score += 1
    # ③ 助詞＝場面や動作の指定（「ぺんぎんがはしる」「うちゅうをおよぐくじら」）
    score += min(2, sum(core.count(p) for p in ("が", "を", "に", "で", "は", "へ")))
    # ④ 「の」「と」＝修飾・複数の登場人物（「ゆうやけのふじさん」「いぬとねこ」）
    if "の" in core: score += 1
    if "と" in core: score += 1
    # ⑤ 数量・複数表現
    if any(k in core for k in ("ふたり", "さんにん", "たくさん", "にひき", "さんびき",
                               "みんな", "いっぱい", "むれ")):
        score += 1

    # 3段階に振り分ける（中間で細かく揺れるより、はっきり差をつけた方が扱いやすい）
    if score <= 1:
        return THINK_MIN          # かんたん：単語1つ（ねこ／りんご）
    if score <= 3:
        return THINK_MID          # ふつう　：修飾つき（かわいい ねこ／ゆうやけのふじさん）
    return THINK_MAX              # むずかしい：場面・複数の要素

def _build_thinking_config(model: str, prompt: str = "") -> dict | None:
    """モデルとお題に応じた thinkingConfig を返す。不要なら None。"""
    m = model.lower()
    # 3系・2.5系のみ thinkingConfig 対応
    is_new = any(m.startswith(p) for p in ("gemini-3", "gemini-2.5"))
    if not is_new:
        return None
    # Pro 系はお題の難易度で可変、Flash 系は明示的に無効化
    if "pro" in m:
        return {"thinkingBudget": _thinking_budget_for(prompt)}
    return {"thinkingBudget": 0}
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", GEMINI_MODELS[0]).strip() or GEMINI_MODELS[0]
if GEMINI_MODEL not in GEMINI_MODELS:
    GEMINI_MODELS.insert(0, GEMINI_MODEL)  # .env に別モデルが保存されていれば先頭に追加
_settings_lock = threading.Lock()

# アプリ設定（管理者画面から変更可能）
app_settings = {
    "bgm_enabled": True,
    "bgm_volume": 80,       # 0〜100
    "draw_speed": 2,        # 0=超高速 1=速い 2=普通 3=ゆっくり 4=とても遅い
    "ai_temperature": 0.85, # 0.5〜1.5
    "draw_status": "idle",  # idle / thinking / drawing
}

# 最近のログ（管理者画面に表示）
_recent_logs = []
_recent_logs_lock = threading.Lock()
MAX_RECENT_LOGS = 50

_env_file_lock = threading.Lock()

def _format_env_value(value) -> str:
    """改行は取り除き、空白や # を含む値は ' ' で囲む（python-dotenv で正しく読めるように）。"""
    v = str(value).replace("\r", "").replace("\n", "").strip()
    if v and (" " in v or "#" in v or v[0] in "'\""):
        v = "'" + v.replace("'", "") + "'"
    return v

def _save_env_value(key, value):
    value = _format_env_value(value)
    os.environ[key] = value.strip("'")
    with _env_file_lock:
        _write_env_line(key, value)

def _write_env_line(key, value):
    path = os.path.join(DATA_DIR, ".env")
    lines = []
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            lines = f.read().splitlines()
    out = []
    found = False
    for line in lines:
        if line.startswith(key + "="):
            out.append(f"{key}={value}")
            found = True
        else:
            out.append(line)
    if not found:
        out.append(f"{key}={value}")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(out) + "\n")

# ===== Flaskの設定 =====
# static_folderをbase_dirにすることで index.html を見つけられるようにする
app = Flask(__name__, static_folder=base_dir, static_url_path='')
# 自分のアプリ（127.0.0.1 / localhost:5000）以外のサイトからは API を呼べないようにする
CORS(app, origins=["http://127.0.0.1:5000", "http://localhost:5000"])

# ==========================================
# サーバー公開時のログイン（ID・パスワード）
# ==========================================
import secrets as _secrets
import hmac as _hmac
from datetime import timedelta

def _load_secret_key() -> str:
    """セッション用の鍵。環境変数 SECRET_KEY が無ければ DATA_DIR に作って保存（再起動してもログイン維持）。"""
    k = (os.environ.get("SECRET_KEY", "") or "").strip()
    if k:
        return k
    path = os.path.join(DATA_DIR, ".secret_key")
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                k = f.read().strip()
        if not k:
            k = _secrets.token_hex(32)
            with open(path, "w", encoding="utf-8") as f:
                f.write(k)
    except Exception:
        k = k or _secrets.token_hex(32)
    return k

app.secret_key = _load_secret_key()
app.config.update(
    PERMANENT_SESSION_LIFETIME=timedelta(days=365),   # 展示 PC は1年間ログインしたまま
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
)

if SERVER_MODE:
    # Render / Cloud Run などのプロキシ越しでも https とアクセス元 IP を正しく扱う
    from werkzeug.middleware.proxy_fix import ProxyFix
    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)
    app.config["SESSION_COOKIE_SECURE"] = (os.environ.get("COOKIE_SECURE", "1").strip() != "0")

def _app_user() -> str:
    return (os.environ.get("APP_USER", "") or "").strip()

def _app_password() -> str:
    return (os.environ.get("APP_PASSWORD", "") or "").strip()

def _is_local_request() -> bool:
    host = (request.host or "").split(":")[0]
    return host in ("127.0.0.1", "localhost", "::1")

_login_fail = {}          # ip -> [失敗回数, ロック解除時刻]
_login_fail_lock = threading.Lock()
_AUTH_OPEN_PATHS = ("/login", "/healthz", "/favicon.ico")

@app.before_request
def _require_login():
    if not SERVER_MODE:
        return None                     # PC 版（desktop）はログイン不要
    if request.path in _AUTH_OPEN_PATHS:
        return None
    if not _app_password():
        # パスワード未設定：Mac での開発（localhost）だけ通す。公開ドメインでは使わせない
        if _is_local_request():
            return None
        return ("<h2>APP_PASSWORD が設定されていません</h2>"
                "<p>サーバーの環境変数に APP_USER と APP_PASSWORD を設定してください。</p>"), 503
    if session.get("auth"):
        return None
    if request.path.startswith("/api/"):
        return jsonify({"error": "login required"}), 401
    return redirect("/login?next=" + request.full_path.rstrip("?"))

def _render_login(error="", nxt="/"):
    with open(os.path.join(base_dir, "login.html"), "r", encoding="utf-8") as f:
        html = f.read()
    return html.replace("{{ERROR}}", str(escape(error))).replace("{{NEXT}}", str(escape(nxt)))

@app.route('/login', methods=['GET', 'POST'])
def login_page():
    if not SERVER_MODE:
        return redirect("/")
    nxt = request.values.get("next", "/") or "/"
    if not nxt.startswith("/") or nxt.startswith("//"):
        nxt = "/"
    if request.method == 'GET':
        return _render_login("", nxt)
    ip = request.remote_addr or "?"
    now = time.time()
    with _login_fail_lock:
        cnt, until = _login_fail.get(ip, [0, 0.0])
    if now < until:
        return _render_login(f"しばらく待ってから試してください（{int(until - now) + 1}秒）", nxt), 429
    user = (request.form.get("username") or "").strip()
    pw = request.form.get("password") or ""
    ok = bool(_app_password()) and _hmac.compare_digest(user, _app_user()) \
        and _hmac.compare_digest(pw, _app_password())
    if ok:
        with _login_fail_lock:
            _login_fail.pop(ip, None)
        session.clear()
        session.permanent = True
        session["auth"] = True
        return redirect(nxt)
    cnt += 1
    with _login_fail_lock:
        _login_fail[ip] = [0, now + 300] if cnt >= 10 else [cnt, 0.0]   # 10回失敗で5分ロック
    time.sleep(0.5)
    return _render_login("ID またはパスワードがちがいます", nxt), 401

@app.route('/logout')
def logout_page():
    session.clear()
    return redirect("/login")

@app.route('/healthz')
def healthz():
    return "ok"

# ==========================================
# 管理者 PIN（タイトル長押し → 暗証番号）
# ==========================================
import secrets
import hmac
from functools import wraps

ADMIN_PIN = (os.environ.get("ADMIN_PIN", "0001") or "0001").strip()
_admin_tokens = set()
_admin_tokens_lock = threading.Lock()
_pin_fail = {"count": 0, "until": 0.0}

def _has_admin_token() -> bool:
    tok = request.headers.get("X-Admin-Token", "") or request.args.get("token", "")
    with _admin_tokens_lock:
        return bool(tok) and tok in _admin_tokens

def require_admin(fn):
    """管理者用 API：PIN 認証で発行したトークンが無ければ 403。"""
    @wraps(fn)
    def wrapper(*a, **kw):
        if not _has_admin_token():
            return jsonify({"error": "admin pin required"}), 403
        return fn(*a, **kw)
    return wrapper

@app.route('/api/admin/verify-pin', methods=['POST'])
def api_admin_verify_pin():
    now = time.time()
    if now < _pin_fail["until"]:
        return jsonify({"ok": False, "error": "locked", "wait": int(_pin_fail["until"] - now) + 1}), 429
    pin = str((request.json or {}).get("pin", "")).strip()
    if hmac.compare_digest(pin, ADMIN_PIN):
        _pin_fail["count"] = 0
        tok = secrets.token_urlsafe(24)
        with _admin_tokens_lock:
            _admin_tokens.add(tok)
        return jsonify({"ok": True, "token": tok, "mode": "server" if SERVER_MODE else LAUNCH_MODE})
    _pin_fail["count"] += 1
    if _pin_fail["count"] >= 5:          # 5回間違えたら30秒ロック
        _pin_fail["count"] = 0
        _pin_fail["until"] = now + 30
    return jsonify({"ok": False}), 401

# ==========================================
# 管理画面から .env を編集する
# ==========================================
def _mask_key(k: str) -> str:
    return (k[:4] + "…" + k[-4:]) if len(k) > 10 else ("設定済み" if k else "")

@app.route('/api/admin/env', methods=['GET'])
@require_admin
def api_admin_env_get():
    return jsonify({
        "api_key_set": bool(API_KEY),
        "api_key_masked": _mask_key(API_KEY),
        "admin_pin": ADMIN_PIN,
        "launch_mode": LAUNCH_MODE_SETTING,
        "launch_mode_running": LAUNCH_MODE,
        "kiosk_block_keys": KIOSK_BLOCK_KEYS,
        "chrome_path": CHROME_PATH,
        "chrome_found": kiosk.find_chrome(CHROME_PATH) or "",
        "fullscreen": (os.environ.get("FULLSCREEN", "1") or "").strip().lower() not in ("0", "false", "no", "off"),
        "env_path": os.path.join(DATA_DIR, ".env"),
        "server_mode": SERVER_MODE,
        "app_user": _app_user(),
        "app_password_set": bool(_app_password()),
    })

@app.route('/api/admin/env', methods=['POST'])
@require_admin
def api_admin_env_set():
    global API_KEY, ADMIN_PIN, LAUNCH_MODE_SETTING, KIOSK_BLOCK_KEYS, CHROME_PATH, _keyblocker
    data = request.json or {}
    changed, restart = [], []

    if "api_key" in data:
        key = str(data["api_key"] or "").strip()
        if key:                                   # 空欄のときは変更しない
            if any(c.isspace() for c in key) or len(key) < 20:
                return jsonify({"error": "APIキーの形式が正しくありません"}), 400
            _save_env_value("GEMINI_API_KEY", key)
            API_KEY = key
            changed.append("APIキー")

    if "admin_pin" in data:
        pin = str(data["admin_pin"] or "").strip()
        if pin != ADMIN_PIN:
            if not (len(pin) == 4 and pin.isdigit()):
                return jsonify({"error": "暗証番号は数字4桁にしてください"}), 400
            _save_env_value("ADMIN_PIN", pin)
            ADMIN_PIN = pin
            changed.append("暗証番号")

    if "launch_mode" in data:
        mode = "window" if str(data["launch_mode"]).lower() == "window" else "browser"
        if mode != LAUNCH_MODE_SETTING:
            _save_env_value("LAUNCH_MODE", mode)
            LAUNCH_MODE_SETTING = mode
            changed.append("起動方法")
            restart.append("起動方法")

    if "kiosk_block_keys" in data:
        on = bool(data["kiosk_block_keys"])
        if on != KIOSK_BLOCK_KEYS:
            _save_env_value("KIOSK_BLOCK_KEYS", "1" if on else "0")
            KIOSK_BLOCK_KEYS = on
            changed.append("キー制限")
            # 今動いているロックにもすぐ反映
            if _keyblocker is not None:
                _keyblocker.enabled = on
            elif on and _kiosk is not None:
                _keyblocker = kiosk.KeyboardBlocker(lambda: _kiosk.pid if _kiosk else None)
                _keyblocker.start()

    if "chrome_path" in data:
        path = str(data["chrome_path"] or "").strip().strip('"')
        if path != CHROME_PATH:
            if path and not os.path.isfile(path):
                return jsonify({"error": "そのパスに Chrome が見つかりません"}), 400
            _save_env_value("CHROME_PATH", path)
            CHROME_PATH = path
            if _kiosk is not None:
                _kiosk.chrome_path = path
            changed.append("Chromeの場所")
            restart.append("Chromeの場所")

    if "fullscreen" in data:
        val = "1" if data["fullscreen"] else "0"
        cur = (os.environ.get("FULLSCREEN", "1") or "").strip().lower() not in ("0", "false", "no", "off")
        if bool(data["fullscreen"]) != cur:
            _save_env_value("FULLSCREEN", val)
            changed.append("全画面（ウィンドウモード）")
            restart.append("全画面（ウィンドウモード）")

    if SERVER_MODE and "app_user" in data:
        u = str(data["app_user"] or "").strip()
        if u and u != _app_user():
            if len(u) < 3 or any(c.isspace() for c in u):
                return jsonify({"error": "ログインIDは空白なしの3文字以上にしてください"}), 400
            _save_env_value("APP_USER", u)
            changed.append("ログインID")
    if SERVER_MODE and data.get("app_password"):
        pw = str(data["app_password"])
        if len(pw) < 8 or any(c.isspace() for c in pw):
            return jsonify({"error": "パスワードは空白なしの8文字以上にしてください"}), 400
        _save_env_value("APP_PASSWORD", pw)
        changed.append("ログインパスワード")

    return jsonify({"status": "ok", "changed": changed, "restart_required": restart})

@app.route('/api/admin/test-key', methods=['POST'])
@require_admin
def api_admin_test_key():
    """APIキーで Gemini に接続できるか確認（入力中のキー、空なら保存済みのキー）。"""
    key = str((request.json or {}).get("api_key") or "").strip() or API_KEY
    if not key:
        return jsonify({"ok": False, "message": "APIキーが入力されていません"})
    try:
        r = requests.get("https://generativelanguage.googleapis.com/v1beta/models",
                         headers={"x-goog-api-key": key}, params={"pageSize": 1}, timeout=10)
        if r.status_code == 200:
            return jsonify({"ok": True, "message": "接続できました"})
        try:
            msg = r.json().get("error", {}).get("message", "")
        except Exception:
            msg = ""
        return jsonify({"ok": False, "message": f"エラー {r.status_code}: {msg[:120]}"})
    except Exception as e:
        return jsonify({"ok": False, "message": f"接続できません（ネットワークを確認）: {type(e).__name__}"})

# 開発時（python desktopaap.py）はプログラムのフォルダがそのまま公開されるので、
# .env や .py などの内部ファイルは返さない
_DENY_EXT = (".py", ".pyc", ".env", ".xlsx", ".log", ".json", ".bat", ".ps1", ".spec", ".md", ".bak")

@app.before_request
def _deny_internal_files():
    path = request.path or ""
    parts = [x for x in path.split("/") if x]
    if any(x.startswith(".") for x in parts) or path.lower().endswith(_DENY_EXT) \
            or (parts and parts[0] in ("logs", "dist", "build", "__pycache__", "data", "tools", "windows", "offline") and path != "/logs/images" and not path.startswith("/logs/images/")):
        return "Not Found", 404

# ==========================================
# 画面表示ルート
# ==========================================
@app.route('/')
def serve_index():
    """ルートURLにアクセスがあったら index.html を返す"""
    target_file = os.path.join(base_dir, 'index.html')
    if not os.path.exists(target_file):
        return f"<h1>404 Not Found</h1><p>index.html が見つかりません。</p><p>探した場所: {target_file}</p>", 404
    return app.send_static_file('index.html')

@app.route('/api/admin/settings', methods=['GET'])
def api_admin_settings():
    with _settings_lock:
        return jsonify({
            "server_mode": SERVER_MODE,
            "print_mode": "browser" if SERVER_MODE else "local",
            "model": GEMINI_MODEL,
            "models": GEMINI_MODELS,
            **app_settings,
        })

@app.route('/api/admin/settings', methods=['POST'])
@require_admin
def api_admin_settings_update():
    global GEMINI_MODEL
    data = request.json or {}
    with _settings_lock:
        if "bgm_enabled" in data:
            app_settings["bgm_enabled"] = bool(data["bgm_enabled"])
        if "bgm_volume" in data:
            app_settings["bgm_volume"] = max(0, min(100, int(data["bgm_volume"])))
        if "draw_speed" in data:
            app_settings["draw_speed"] = max(0, min(4, int(data["draw_speed"])))
        if "ai_temperature" in data:
            app_settings["ai_temperature"] = max(0.1, min(2.0, float(data["ai_temperature"])))
        if "model" in data:
            model = str(data["model"]).strip()
            if model:
                if model not in GEMINI_MODELS:
                    GEMINI_MODELS.append(model)
                GEMINI_MODEL = model
                _save_env_value("GEMINI_MODEL", model)
    return jsonify({"status": "ok", "model": GEMINI_MODEL, "models": GEMINI_MODELS, **app_settings})

@app.route('/api/admin/save', methods=['GET'])
@require_admin
def api_admin_save():
    """BGM・スピード・温度などをクエリパラメーターで保存（WebView GET専用）"""
    with _settings_lock:
        v = request.args.get("bgm_enabled")
        if v is not None:
            app_settings["bgm_enabled"] = v.lower() not in ("false", "0", "")
        v = request.args.get("bgm_volume")
        if v is not None:
            try: app_settings["bgm_volume"] = max(0, min(100, int(v)))
            except ValueError: pass
        v = request.args.get("draw_speed")
        if v is not None:
            try: app_settings["draw_speed"] = max(0, min(4, int(v)))
            except ValueError: pass
        v = request.args.get("ai_temperature")
        if v is not None:
            try: app_settings["ai_temperature"] = max(0.1, min(2.0, float(v)))
            except ValueError: pass
    return jsonify({"status": "ok", **app_settings})

# ==========================================
# BGM 一覧（考え中BGMのランダム再生用）
# ==========================================
BGM_DIR_NAME   = "BGM"
BGM_COMPLETE   = "BGM2.mp3"          # 完成時の効果音。ランダム対象からは除外する
BGM_EXTENSIONS = (".mp3", ".wav", ".ogg", ".m4a")

@app.route('/api/bgm/list', methods=['GET'])
def api_bgm_list():
    """BGMフォルダ内の音源一覧を返す。
    BGM3.mp3 / BGM4.mp3 … とファイルを置くだけでランダム再生の対象が増える。
    完成音（BGM2.mp3）だけは対象から外す。"""
    tracks = []
    bgm_dir = os.path.join(base_dir, BGM_DIR_NAME)
    try:
        if os.path.isdir(bgm_dir):
            for name in sorted(os.listdir(bgm_dir)):
                if not name.lower().endswith(BGM_EXTENSIONS):
                    continue
                if name.lower() == BGM_COMPLETE.lower():
                    continue
                tracks.append(f"{BGM_DIR_NAME}/{name}")
    except Exception as e:
        print(f"[BGM] 一覧の取得に失敗: {e}")
    return jsonify({
        "tracks": tracks,
        "complete": f"{BGM_DIR_NAME}/{BGM_COMPLETE}",
    })

@app.route('/api/admin/status', methods=['GET', 'POST'])
def api_admin_status_update():
    # GET（クエリパラメーター）と POST（JSON）の両対応
    if request.method == 'GET':
        val = request.args.get("draw_status", "")
    else:
        data = request.json or {}
        val = data.get("draw_status", "")
    if val:
        with _settings_lock:
            app_settings["draw_status"] = str(val)
    return jsonify({"status": "ok"})

@app.route('/api/admin/model', methods=['GET', 'POST'])
@require_admin
def api_admin_model():
    global GEMINI_MODEL
    # GET / POST 両対応（WebView 環境では GET の方が確実）
    if request.method == 'GET':
        model = (request.args.get("model") or "").strip()
        temp_str = request.args.get("temperature", "")
    else:
        data = request.json or {}
        model = (data.get("model") or "").strip()
        temp_str = str(data.get("temperature", ""))
    if not model:
        return jsonify({"error": "model is empty"}), 400
    with _settings_lock:
        if model not in GEMINI_MODELS:
            GEMINI_MODELS.append(model)
        GEMINI_MODEL = model
        _save_env_value("GEMINI_MODEL", model)
        if temp_str:
            try:
                app_settings["ai_temperature"] = max(0.1, min(2.0, float(temp_str)))
            except ValueError:
                pass
    return jsonify({"status": "ok", "model": GEMINI_MODEL, "models": GEMINI_MODELS,
                    "ai_temperature": app_settings["ai_temperature"]})

@app.route('/api/admin/logs', methods=['GET'])
def api_admin_logs():
    with _recent_logs_lock:
        return jsonify({"logs": list(_recent_logs)})

@app.route('/debug')
def debug_page():
    """リアルタイムデバッグページ（ブラウザで表示）"""
    return '''<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>AI Drawing - Debug</title>
<style>
  body{font-family:monospace;background:#111;color:#ddd;margin:0;padding:12px;font-size:13px}
  h1{color:#7cf;margin:0 0 8px;font-size:16px}
  #status{display:flex;gap:16px;flex-wrap:wrap;background:#1a1a1a;padding:8px 12px;border-radius:6px;margin-bottom:10px}
  .kv{display:flex;gap:6px;align-items:center}
  .k{color:#888}
  .v{color:#afa;font-weight:bold}
  #log-wrap{background:#1a1a1a;border-radius:6px;padding:8px 12px;height:calc(100vh - 130px);overflow-y:auto}
  .row{padding:3px 0;border-bottom:1px solid #222;display:flex;gap:8px}
  .t{color:#555;min-width:80px}
  .tp{min-width:120px;font-weight:bold}
  .tp.DRAW_START{color:#58f}
  .tp.DRAW_SUCCESS{color:#4d4}
  .tp.DRAW_ERROR{color:#f44}
  .tp.INPUT_MODE{color:#fa0}
  .d{color:#ccc;word-break:break-all}
  #indicator{float:right;font-size:11px;color:#555}
</style>
</head>
<body>
<h1>🔍 AI Drawing Debug <span id="indicator">●</span></h1>
<div id="status"></div>
<div id="log-wrap"><div id="logs"></div></div>
<script>
let lastCount = 0;
function esc(s){return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;")}
async function poll(){
  try{
    const r = await fetch("/api/debug/data");
    const d = await r.json();
    // ステータス
    const s = d.settings||{};
    document.getElementById("status").innerHTML =
      `<div class="kv"><span class="k">model</span><span class="v">${esc(s.model||"?")}</span></div>`+
      `<div class="kv"><span class="k">status</span><span class="v">${esc(s.draw_status||"idle")}</span></div>`+
      `<div class="kv"><span class="k">BGM</span><span class="v">${s.bgm_enabled?"ON":"OFF"} ${s.bgm_volume??80}%</span></div>`+
      `<div class="kv"><span class="k">speed</span><span class="v">${s.draw_speed??2}</span></div>`+
      `<div class="kv"><span class="k">temp</span><span class="v">${(s.ai_temperature??0.85).toFixed(2)}</span></div>`;
    // ログ
    const logs = d.logs||[];
    if(logs.length !== lastCount){
      lastCount = logs.length;
      const wrap = document.getElementById("log-wrap");
      const atBottom = wrap.scrollHeight - wrap.scrollTop <= wrap.clientHeight + 40;
      document.getElementById("logs").innerHTML = [...logs].reverse().map(l=>
        `<div class="row"><span class="t">${esc(l.time)}</span><span class="tp ${esc(l.type)}">${esc(l.type)}</span><span class="d">${esc(l.data)}</span></div>`
      ).join("");
      if(atBottom) wrap.scrollTop = wrap.scrollHeight;
    }
    document.getElementById("indicator").style.color="#4d4";
  }catch(e){document.getElementById("indicator").style.color="#f44"}
}
poll();
setInterval(poll,1500);
</script>
</body>
</html>'''

def debug_page_excel():
    return """<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>AI Drawing Debug</title>
<style>
  body{margin:0;padding:14px;font-family:Segoe UI,system-ui,sans-serif;background:#fafafa;color:#222;font-size:13px}
  h1{font-size:18px;margin:0 0 8px}
  #indicator{font-size:12px;color:#999}
  #status{display:flex;gap:14px;flex-wrap:wrap;background:#fff;border:1px solid #e5e5e5;border-radius:6px;padding:8px 10px;margin-bottom:10px}
  .kv{display:flex;gap:5px}.k{color:#666}.v{font-weight:700;color:#0b7a2a}
  .meta{font-size:12px;color:#555;margin-bottom:10px}
  .wrap{height:calc(100vh - 136px);overflow:auto;background:#fff;border:1px solid #ddd;border-radius:6px}
  table{border-collapse:collapse;width:100%;font-size:12px}
  th,td{border-right:1px solid #eee;border-bottom:1px solid #eee;padding:5px 7px;white-space:nowrap;vertical-align:top}
  th{position:sticky;top:0;background:#ff6b35;color:#fff;z-index:1}
  td.num{text-align:right;font-variant-numeric:tabular-nums}
  td.prompt,td.detail,td.err{max-width:360px;white-space:normal}
  td.err{color:#c62828}
  .ev-DRAW_SUCCESS{color:#2e7d32;font-weight:700}
  .ev-DRAW_ERROR,.ev-PRINT_ERROR{color:#c62828;font-weight:700}
  .ev-PRINT_SUCCESS{color:#1565c0;font-weight:700}
  .ev-DRAW_START,.ev-INPUT_MODE{color:#6a1b9a;font-weight:700}
</style>
</head>
<body>
<h1>AI Drawing Debug <span id="indicator">●</span></h1>
<div id="status"></div>
<div class="meta">debug.xlsx: <b id="xlsxPath"></b> / 保留: <b id="pending">0</b>件 / 概算費用: <b id="cost">0</b>円</div>
<div class="wrap">
<table>
<thead><tr>
<th>日時</th><th>種別</th><th>お題</th><th>所要秒数</th><th>キャンバス幅</th><th>キャンバス高</th>
<th>AIコマンド数</th><th>採用コマンド数</th><th>除去背景コマンド</th><th>リトライ回数</th>
<th>エラー内容</th><th>画像ファイル</th><th>モデル名</th><th>入力トークン</th><th>出力トークン</th>
<th>思考トークン</th><th>概算費用(円)</th><th>備考</th><th>評価★</th><th>ID</th>
</tr></thead>
<tbody id="rows"></tbody>
</table>
</div>
<script>
const KEYS=["_ts","event_type","prompt","duration_sec","canvas_w","canvas_h","cmd_raw","cmd_filtered","removed_bg","retries","error","image_file","model","tokens_in","tokens_out","tokens_thought","cost_jpy","detail","rating","event_id"];
const NUM=new Set(["duration_sec","canvas_w","canvas_h","cmd_raw","cmd_filtered","removed_bg","retries","tokens_in","tokens_out","tokens_thought","cost_jpy","rating"]);
let last=-1;
function esc(v){return String(v??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]))}
async function poll(){
  try{
    const r=await fetch("/api/debug/data",{cache:"no-store"});
    const d=await r.json();
    const s=d.settings||{};
    document.getElementById("status").innerHTML=
      `<div class="kv"><span class="k">model</span><span class="v">${esc(s.model||"?")}</span></div>`+
      `<div class="kv"><span class="k">status</span><span class="v">${esc(s.draw_status||"idle")}</span></div>`+
      `<div class="kv"><span class="k">BGM</span><span class="v">${s.bgm_enabled?"ON":"OFF"} ${s.bgm_volume??80}%</span></div>`+
      `<div class="kv"><span class="k">speed</span><span class="v">${s.draw_speed??2}</span></div>`+
      `<div class="kv"><span class="k">temp</span><span class="v">${Number(s.ai_temperature??0.85).toFixed(2)}</span></div>`;
    document.getElementById("xlsxPath").textContent=d.xlsx_path||"";
    document.getElementById("pending").textContent=d.pending_xlsx??0;
    document.getElementById("cost").textContent=Number(d.total_cost_jpy??0).toLocaleString(undefined,{minimumFractionDigits:4,maximumFractionDigits:4});
    const evs=d.events||[];
    if(evs.length!==last){
      last=evs.length;
      document.getElementById("rows").innerHTML=evs.slice().reverse().map(ev=>"<tr>"+KEYS.map(k=>{
        let v=ev[k], cls="";
        if(k==="event_type") cls="ev-"+esc(v);
        if(NUM.has(k)) cls=(cls+" num").trim();
        if(k==="prompt") cls=(cls+" prompt").trim();
        if(k==="detail") cls=(cls+" detail").trim();
        if(k==="error"&&v) cls=(cls+" err").trim();
        if(k==="image_file"&&v) v='<a href="/'+esc(v)+'" target="_blank">'+esc(v)+'</a>'; else v=esc(v);
        return '<td'+(cls?' class="'+cls+'"':'')+'>'+v+'</td>';
      }).join("")+"</tr>").join("");
    }
    document.getElementById("indicator").style.color="#2e7d32";
  }catch(e){document.getElementById("indicator").style.color="#c62828"}
}
poll(); setInterval(poll,1500);
</script>
</body>
</html>"""

app.view_functions["debug_page"] = debug_page_excel

@app.route('/api/debug/data')
def api_debug_data():
    _load_xlsx_history_into_buffer()
    with _settings_lock:
        s = {"model": GEMINI_MODEL, "models": GEMINI_MODELS, **app_settings}
    with _recent_logs_lock:
        logs = list(_recent_logs)
    with _xlsx_lock:
        events = list(_events_buffer)
        pending = len(_xlsx_pending)
    total_cost = 0.0
    for ev in events:
        c = ev.get("cost_jpy")
        if isinstance(c, (int, float)):
            total_cost += float(c)
    return jsonify({
        "settings": s,
        "logs": logs,
        "events": events,
        "pending_xlsx": pending,
        "total_cost_jpy": round(total_cost, 4),
        "xlsx_path": _xlsx_path(),
    })

@app.route('/logs/images/<path:fname>', methods=['GET'])
def debug_image(fname):
    return send_from_directory(_images_dir(), fname)

@app.route('/api/admin/open-debug', methods=['POST'])
@require_admin
def api_admin_open_debug():
    if SERVER_MODE or LAUNCH_MODE == "browser":
        # キオスク中は別ウィンドウが裏に隠れるので、管理画面の中で開く
        return jsonify({"status": "inline", "url": "/debug"})
    subprocess.Popen(["cmd", "/c", "start", "", "http://127.0.0.1:5000/debug"], shell=False)
    return jsonify({"status": "ok"})

@app.route('/api/admin/shutdown', methods=['POST'])
@require_admin
def api_admin_shutdown():
    if SERVER_MODE:
        return jsonify({"error": "サーバー版ではアプリを終了できません。画面は Alt+F4 で閉じてください"}), 400
    def quit_later():
        time.sleep(0.25)
        _shutdown_everything()
    threading.Thread(target=quit_later, daemon=True).start()
    return jsonify({"status": "closing"})

@app.route('/api/admin/open-window', methods=['POST'])
@require_admin
def api_admin_open_window():
    if SERVER_MODE:
        return jsonify({"error": "server mode"}), 400
    browser = kiosk.find_chrome(CHROME_PATH) or kiosk.find_edge()
    if browser:
        subprocess.Popen([browser, "--app=http://127.0.0.1:5000/admin.html", "--window-size=480,700"])
    return jsonify({"status": "ok"})

# ==========================================
# ログ保存機能
# ==========================================
LOG_DIR = os.path.join(DATA_DIR, "logs")
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)

LOG_FILE = os.path.join(LOG_DIR, "terminal_usage.log")

@app.route('/api/ready')
def api_ready():
    """ネットワーク接続確認エンドポイント（起動ローディング画面から呼ばれる）。"""
    network_ok = False
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.settimeout(2)
        s.connect(('8.8.8.8', 53))
        s.close()
        network_ok = True
    except OSError:
        pass
    return jsonify({
        "ready": network_ok and bool(API_KEY),
        "network": network_ok,
        "api_key": bool(API_KEY),
    })


# ==========================================
# 音声の文字起こし（Web Speech API が使えないときの保険）
# ==========================================
_TRANSCRIBE_PROMPT = (
    "音声の文字起こしをします。\n"
    "この音声に人の話し声が入っているかを先に判断してください。\n"
    "入っていない場合（無音・雑音・音楽・機械音・電子音だけ）は、必ず #なし# とだけ答えます。"
    "内容を想像して答えを作ることは絶対にしないでください。\n"
    "人の話し声が入っている場合は、話した言葉だけを日本語で1行書きます"
    "（説明・かぎかっこ・句読点は書かない）。\n"
    "話しているのは日本の子どもで、AIに描いてほしいものを言っています。"
)


def _transcribe_model() -> str:
    """文字起こしに使うモデル。速くて安い flash 系を優先する。"""
    if "gemini-2.5-flash" in GEMINI_MODELS:
        return "gemini-2.5-flash"
    for m in GEMINI_MODELS:
        low = m.lower()
        if "flash" in low and "lite" not in low:
            return m
    return GEMINI_MODEL


@app.route('/api/transcribe', methods=['POST'])
def api_transcribe():
    """ブラウザで録音した WAV を Gemini に送って文字起こしする。

    WebView2 の Web Speech API は Microsoft/Google の音声サーバーに接続するため、
    通信が不安定だと 'network' エラーでマイクが使えなくなる。
    そのときにフロント側からこのエンドポイントへ切り替える。
    """
    if not API_KEY:
        return jsonify({"error": "GEMINI_API_KEY が設定されていません"}), 500

    data = request.get_json(silent=True) or {}
    audio_b64 = (data.get("audio") or "").strip()
    mime = (data.get("mime") or "audio/wav").strip()
    if not audio_b64:
        return jsonify({"error": "音声データがありません"}), 400
    # base64 で約14MB（音声20秒程度）を超えるものは受け付けない
    if len(audio_b64) > 14_000_000:
        return jsonify({"error": "音声が長すぎます"}), 413

    model = _transcribe_model()
    gen_config = {"temperature": 0, "maxOutputTokens": 128}
    # 文字起こしは「考えさせない」方が正確。考えさせると声が無くても言葉を作ってしまう
    if any(model.lower().startswith(pfx) for pfx in ("gemini-3", "gemini-2.5")):
        gen_config["thinkingConfig"] = {"thinkingBudget": 0}

    payload = {
        "contents": [{"parts": [
            {"inlineData": {"mimeType": mime, "data": audio_b64}},
            {"text": _TRANSCRIBE_PROMPT},
        ]}],
        "generationConfig": gen_config,
    }
    url = (
        f"https://generativelanguage.googleapis.com/v1beta/models/"
        f"{model}:generateContent"
    )

    last_error = ""
    for i in range(2):
        try:
            res = requests.post(
                url, json=payload,
                headers=_gemini_headers(),
                timeout=25,
            )
            body = res.json()
        except Exception as e:
            last_error = f"通信エラー: {e}"
            print(f"[Transcribe] {last_error}")
            if i == 0:
                time.sleep(1)
                continue
            break

        if isinstance(body, dict) and body.get('error'):
            err = body['error']
            last_error = err.get('message', str(err))
            print(f"[Transcribe] APIエラー model={model}: {last_error}")
            if err.get('code') == 503 or err.get('status') == 'UNAVAILABLE':
                if i == 0:
                    time.sleep(1)
                    continue
            break

        if not res.ok:
            last_error = f"HTTP {res.status_code}: {res.text[:200]}"
            print(f"[Transcribe] {last_error}")
            if i == 0:
                time.sleep(1)
                continue
            break

        candidates = body.get('candidates', []) if isinstance(body, dict) else []
        text = ''
        if candidates:
            for p in (candidates[0].get('content', {}).get('parts', []) or []):
                if isinstance(p, dict) and 'text' in p and not p.get('thought'):
                    text = p['text']
                    break
        text = (text or '').strip().strip('。「」"\'')
        # 話し声が無いときのしるし。そのまま返すと変な言葉が絵になってしまう
        if '#なし#' in text or text in ('なし', '無し'):
            text = ''
        text = ' '.join(text.split())
        print(f"[Transcribe] model={model} 結果=「{text}」")
        return jsonify({"text": text, "model": model})

    return jsonify({"error": last_error or "文字起こしに失敗しました"}), 502



@app.route('/api/log', methods=['POST'])
def receive_log():
    data = request.json
    log_type = data.get("type", "unknown")
    log_data = data.get("data", "")
    timestamp = data.get("time", datetime.now().isoformat())
    log_entry = f"[{timestamp}] [{log_type.upper()}] {log_data}\n"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(log_entry)
    except Exception as e:
        print(f"Log write error: {e}")
    # 管理者画面用に最新ログを保持
    short_time = timestamp[11:19] if len(timestamp) >= 19 else timestamp
    with _recent_logs_lock:
        _recent_logs.append({"time": short_time, "type": log_type.upper(), "data": log_data})
        if len(_recent_logs) > MAX_RECENT_LOGS:
            _recent_logs.pop(0)
    try:
        log_to_xlsx(log_type.upper(), detail=str(log_data))
    except Exception as e:
        print(f"[XLSX] frontend log write skipped: {e}")
    return jsonify({"status": "ok"})

def _is_white_ish(color: str) -> bool:
    """#rrggbb または #rgb が白系（明度90%以上）かどうかを判定する。"""
    c = color.strip().lstrip('#')
    try:
        if len(c) == 3:
            c = c[0]*2 + c[1]*2 + c[2]*2
        if len(c) != 6:
            return False
        r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
        return r >= 230 and g >= 230 and b >= 230
    except Exception:
        return False

def _filter_cmds(text: str) -> str:
    """AIが出力したコマンド列から禁止コマンドを除去して返す。"""
    import re
    clean = re.sub(r'```json|```', '', text).strip()
    try:
        cmds = json.loads(clean)
    except Exception:
        return text  # パース失敗時はそのまま返す（JS側でエラーにする）

    filtered = []
    for cmd in cmds:
        if not isinstance(cmd, dict):
            continue
        c = cmd.get('cmd', '')
        # background / gradient_rect は除去
        if c in ('background', 'gradient_rect'):
            continue
        # 白系半透明（テカリ・ハイライト）を除去：alpha < 0.85 かつ color が白系
        alpha = cmd.get('alpha', 1.0)
        color = cmd.get('color', '#000000')
        if alpha < 0.85 and _is_white_ish(color):
            continue
        filtered.append(cmd)

    return json.dumps(filtered, ensure_ascii=False)

# ==========================================
# 入力バリデーション
# ==========================================
_NG_WORDS = [
    'ちんちん','おちんちん','ちんぽ','ちんぽこ','まんこ','おまんこ',
    'ちくび','やりまん','おかま',
    'せっくす','セックス','えっち','エッチ','えろ','エロ',
    'ぽるの','ポルノ','あだると','アダルト',
    'しこしこ','ふぇら','くんに','クンニ',
    'sex','porn','naked','nude','fuck','dick','pussy','cock','hentai',
]

def _contains_ng_word(text: str) -> bool:
    t = text.lower().replace(' ', '').replace('　', '')
    return any(w.lower() in t for w in _NG_WORDS)

# --- 日本語として成立しないかな並びの判定用テーブル ---
# 小書き文字（ゃゅょぁぃぅぇぉゎ）の直前に来てよい文字。
# 「ぶぃ」「はゃ」のような、キーボード連打でしか出ない並びを弾くために使う。
_SMALL_KANA_PREV = {
    'ゃ': set('きしちにひみりぎじぢびぴ'),
    'ゅ': set('きしちにひみりぎじぢびぴてでふゔ'),
    'ょ': set('きしちにひみりぎじぢびぴ'),
    'ぁ': set('ふうゔつくぐ'),
    'ぃ': set('ふうゔつてですず'),
    'ぅ': set('とどうゔ'),
    'ぇ': set('しちじふうゔついでき'),
    'ぉ': set('ふうゔつくぐ'),
    'ゎ': set('くぐ'),
}
# 「っ」の次に来てよくない文字（母音・ん・長音・小書き）
_SOKUON_NG_NEXT = set('あいうえおんーぁぃぅぇぉゃゅょっゎ')

def _kana_sequence_broken(core: str) -> bool:
    """かなの並びが日本語として成立しているかを調べる。
    成立しない並びが1つでもあれば True（でたらめ入力とみなす）。"""
    # カタカナはひらがなに寄せてから判定する
    s = ''.join(chr(ord(c) - 0x60) if 'ァ' <= c <= 'ヶ' else c for c in core)
    for i, c in enumerate(s):
        prev = s[i - 1] if i > 0 else ''
        # 小書き文字：直前が正しい組み合わせでなければ NG（先頭に来るのも NG）
        if c in _SMALL_KANA_PREV:
            if not prev or prev not in _SMALL_KANA_PREV[c]:
                return True
        # 促音「っ」：語末・語頭、および母音や「ん」の直前は NG
        elif c == 'っ':
            nxt = s[i + 1] if i + 1 < len(s) else ''
            if not prev or not nxt or nxt in _SOKUON_NG_NEXT:
                return True
        # 「ん」「ー」が先頭に来るのは NG
        elif c in 'んー' and i == 0:
            return True
    return False

def _is_gibberish(text: str) -> bool:
    """意味のないランダム入力を検出する。"""
    import re, unicodedata
    from collections import Counter
    t = text.strip()
    if not t:
        return False
    if len(t) <= 1:
        return False
    core = t.replace(' ', '').replace('　', '')
    # 同じ文字のみ（aaaa / あああ）
    if len(set(core)) == 1:
        return True
    # 4文字以上連続して同じ文字（ああああ 等）
    # ※スペースは消さずに判定する。「かわいい いぬ」を繋げると「いいい」になり、
    #   3連続だと正常な入力まで弾いてしまうため。
    if re.search(r'(.)\1{3,}', t):
        return True
    # 長すぎる（お題として現実的でない）
    if len(core) > 40:
        return True
    # 小文字かな（っぁぃぅぇぉゃゅょゎ）が半分以上 → でたらめ連打
    small_kana = set('っぁぃぅぇぉゃゅょゎッァィゥェォャュョヮ')
    if len(core) >= 3 and sum(1 for c in core if c in small_kana) / len(core) >= 0.5:
        return True
    # 日本語として成立しないかなの並び（ぶぃ / はゃ / っ の位置がおかしい 等）
    if _kana_sequence_broken(core):
        return True
    # かなとアルファベットが混ざった連打（IME未確定のまま打ち続けた状態）
    #   ・1文字だけのアルファベットが2つ以上（ｆなｆはｋ…）
    #   ・アルファベットが合計4文字以上（あsdfghjkl）
    if re.search(r'[ぁ-ゖァ-ヶ]', core):
        letters = re.findall(r'[A-Za-zＡ-Ｚａ-ｚ]', core)
        isolated = re.findall(r'(?<![A-Za-zＡ-Ｚａ-ｚ])[A-Za-zＡ-Ｚａ-ｚ](?![A-Za-zＡ-Ｚａ-ｚ])', core)
        if len(isolated) >= 2 or len(letters) >= 4:
            return True
    # 記号だらけ（意味のある文字が30%未満）
    meaningful = sum(1 for c in t if unicodedata.category(c) in
                     ('Lo','Ll','Lu','Nd','Lm') or '぀' <= c <= '鿿')
    if len(t) >= 3 and meaningful / len(t) < 0.3:
        return True
    # 1種類の文字が70%以上
    counts = Counter(core)
    if counts and counts.most_common(1)[0][1] / len(core) > 0.7:
        return True
    return False

# ==========================================
# Gemini AI 描画API
# ==========================================
@app.route('/api/draw', methods=['POST'])
def draw_api():
    if not API_KEY:
        return jsonify({"error": "GEMINI_API_KEY が設定されていません"}), 500

    data = request.json
    user_prompt = data.get('prompt', '')
    w, h = data.get('width', 800), data.get('height', 600)
    _draw_t0 = time.time()

    # 入力内容による拒否は "blocked" を付けて返す。
    # 画面側はこれを見て、エラー画面ではなく子ども向けの案内を出す。
    if _contains_ng_word(user_prompt):
        return jsonify({"error": "そのことばは使えません。べつのことばをためしてね！",
                        "blocked": "ng_word"}), 400

    if _is_gibberish(user_prompt):
        return jsonify({"error": "なにをかきたいか、もうすこしくわしくおしえてね！",
                        "blocked": "gibberish"}), 400

    current_model = GEMINI_MODEL
    # Flash系は高速化のためコマンド数・トークン数を削減
    is_flash = "flash" in current_model.lower()
    cmd_range  = "18〜24" if is_flash else "24〜36"
    cmd_min    = 12       if is_flash else 18
    max_tokens = 3072     if is_flash else 8192
    detail_parts = "4〜6個" if is_flash else "5〜8個"

    prompt_text = (
    f"あなたはHTML Canvas 2D APIの描画コマンドだけを使って、子供向けの『ひと目で何かわかる、シンプルでかわいい絵』を描く一流のAIアーティストです。\n"
    f"線画のラフではなく、色をたくさん塗り重ねた『仕上がった絵』を作ってください。\n"
    f"キャンバスは最初から白です。背景・影・テカリ（白系半透明）は一切描かない。\n"
    f"\n"
    f"==============================\n"
    f"【絶対禁止（これをやると失敗）】\n"
    f"==============================\n"
    f'- "background" コマンドを使わない ✗\n'
    f'- "gradient_rect" コマンドを使わない ✗\n'
    f"- 影：黒・グレー・暗色の半透明（alpha < 0.7 で color が暗い）コマンドを使わない ✗\n"
    f"- テカリ・ハイライト：白系（#fff, #ffffff, #ffe など明度90%以上）の半透明 ellipse / circle / rect を使わない ✗\n"
    f"- キャンバス全体やその半分以上を覆う大きな rect / ellipse を使わない ✗\n"
    f"- 丸い果物（りんご・みかん・もも・ぶどう粒・さくらんぼ等）の本体を curve や polygon で描かない ✗\n"
    f"  → 必ず circle か ellipse コマンドだけで丸く描くこと。これは絶対ルール。\n"
    f"- 出力は JSON配列のみ。前後の説明・コードブロック ``` は一切禁止。\n"
    f"- JSONは構文的に必ず正しく、最後の要素のあとにカンマを置かない。\n"
    f"\n"
    f"==============================\n"
    f"【絶対ルール】\n"
    f"==============================\n"
    f"- キャンバスサイズ: 幅 {w}px × 高さ {h}px。すべての座標はこの範囲内。\n"
    f"- 必ず {cmd_range} 個のコマンドで描く。\n"
    f"- 主役は『お題そのもの』。お題が「いぬ」なら必ず犬とわかる絵にする。抽象画にしないこと。\n"
    f"\n"
    f"==============================\n"
    f"【描く順番（必ずこの順番で並べる）】\n"
    f"==============================\n"
    f"1. 装飾・添景（任意）：主役を引き立てる小さな要素（花・星・葉など）。3〜5個まで。\n"
    f"2. 主役の体：画面中央付近、サイズは画面の 40〜60%。複数の ellipse / circle / curve を重ねてはっきり描く。\n"
    f"3. 主役のディテール：目・口・鼻・模様・しっぽ・羽・葉・へた等。小さなパーツを {detail_parts}。輪郭は curve で明確に描く。\n"
    f"\n"
    f"==============================\n"
    f"【上手に見せるテクニック】\n"
    f"==============================\n"
    f"- 主役は大きく・はっきり・中央に。子供が一瞬で何かわかるシルエットにする。\n"
    f"- 輪郭は curve を 3〜5本使って、対象物の形がくっきりわかるように描く。\n"
    f"- 【重要】果物・食べ物・乗り物・植物・建物など『生き物でないもの』には目・顔・表情を一切つけない。\n"
    f"- 動物・人・キャラクターの場合のみ目を描く（白目→黒目の 2層構造、白系 ellipse は使わない）。\n"
    f"- 完全な左右対称はぎこちないので、左右の大きさ・位置をわずかにずらす。\n"
    f"- 立体感は『同じ形をやや小さく＋やや明るい色で内側に重ねる』で表現する（白系半透明は使わない）。\n"
    f"\n"
    f"==============================\n"
    f"【色の使い方】\n"
    f"==============================\n"
    f"- 主役・差し色・輪郭の 3系統を最初に決め、その範囲内で微妙に変化させる。\n"
    f"- 子供向けなのでビビッドで明るい色を中心に。くすみ色やグレー単独は使わない。\n"
    f"- 動物の場合は実際の色（犬=茶系/白系、ぞう=灰系、ライオン=黄〜オレンジ等）を基本にしつつ、彩度を上げて元気な印象にする。\n"
    f"- 果物・食べ物は実物の色と形を最優先で守る。\n"
    f"  りんご/みかん/もも/ぶどう粒 → 必ず `circle` か `ellipse` コマンドで丸く描く（curve/polygon 禁止）。\n"
    f"  ばなな → curve コマンドで黄色い弓型を描く。りんごの赤+ヘタ必須、ぶどうは丸粒の集まり、すいかは緑楕円+縞。\n"
    f"- 果物・食べ物に目・顔・口などキャラクター要素を加えない（あくまで食べ物として描く）。\n"
    f"\n"
    f"==============================\n"
    f"【利用できるコマンド】\n"
    f"==============================\n"
    f"必ずこの形式の JSON オブジェクトを使う。記載のないキーは付けない。\n"
    f'- {{"cmd":"rect","x":0,"y":0,"w":100,"h":100,"color":"#hex","alpha":1}}\n'
    f'- {{"cmd":"circle","cx":200,"cy":200,"r":50,"color":"#hex","alpha":1}}\n'
    f'- {{"cmd":"ellipse","cx":200,"cy":200,"rx":80,"ry":40,"color":"#hex","alpha":1}}\n'
    f'- {{"cmd":"line","x1":0,"y1":0,"x2":100,"y2":100,"color":"#hex","width":2,"alpha":1}}\n'
    f'- {{"cmd":"curve","points":[[x1,y1],[cx1,cy1],[cx2,cy2],[x2,y2]],"color":"#hex","width":4,"fill":false,"alpha":1}}\n'
    f'- {{"cmd":"polygon","points":[[x,y],[x,y],[x,y]],"color":"#hex","fill":true,"alpha":1}}\n'
    f'- {{"cmd":"arc","cx":200,"cy":200,"r":50,"start":0,"end":3.14,"color":"#hex","width":3,"alpha":1}}\n'
    f"\n"
    f"==============================\n"
    f"【座標と大きさの目安】\n"
    f"==============================\n"
    f"- 中心は ({w//2}, {h//2})。\n"
    f"- 主役の体はだいたい cx=({w//2}±{w//10}), cy=({h//2}+{h//12}), 半径や rx は {min(w,h)//4} 〜 {min(w,h)//3}。\n"
    f"- 動物・キャラクターの目は体の中心より上 ({h//12} 程度上)に、左右に {min(w,h)//12} ほど離して配置。\n"
    f"- 果物の例：ばなな=黄色の弓型、りんご=赤い丸、みかん=オレンジの丸、すいか=緑と黒の縞+赤い断面。\n"
    f"\n"
    f"==============================\n"
    f"【避けるべき失敗】\n"
    f"==============================\n"
    f"- コマンド数が {cmd_min} 個以下になる ✗\n"
    f"- 主役が小さく(画面の 25% 未満)、寂しく見える ✗\n"
    f"- お題と全く違うものを描く ✗\n"
    f"- 線画のみで色塗りがない ✗\n"
    f"- 背景・影・白系テカリを描く ✗\n"
    f"- 何を描いたかわかりにくい抽象的な形にする ✗（子供が一目でわかること）\n"
    f"- 果物・食べ物・乗り物・植物などに目・顔・口をつける ✗（生き物でないものに表情は不要）\n"
    f"- お題と色・形が明らかに違うものを描く ✗（ばなな=黄色い弓型、りんご=赤い丸形 を必ず守る）\n"
    f"- りんご・みかん等の丸い果物を curve や polygon で描いてガタガタにする ✗（必ず circle/ellipse を使う）\n"
    f"\n"
    f"==============================\n"
    f"【お題】\n"
    f"==============================\n"
    f"「{user_prompt}」\n"
    f"\n"
    f"上記すべてを守って、JSON 配列のみを出力してください。"
)

    print(f"[draw_api] model={current_model} flash={is_flash} tokens={max_tokens} cmds={cmd_range}")

    gen_config = {
        "maxOutputTokens": max_tokens,
        "temperature": app_settings["ai_temperature"],
        "responseMimeType": "application/json",
    }
    tc = _build_thinking_config(current_model, user_prompt)
    if tc is not None:
        gen_config["thinkingConfig"] = tc
        print(f"[draw_api] thinkingBudget={tc['thinkingBudget']} (お題: {user_prompt})")

    payload = {
        "contents": [{"parts": [{"text": prompt_text}]}],
        "generationConfig": gen_config,
    }
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{current_model}:generateContent"

    # 各モデルあたり2回ずつ。proは90s、flashは18sタイムアウト
    req_timeout = 18 if is_flash else 90

    # フォールバックチェーン: current_model(2回) → 次のpro(2回) → flash(2回)
    attempt_models = [current_model]
    if not is_flash:
        try:
            start_idx = GEMINI_MODELS.index(current_model)
        except ValueError:
            start_idx = 0
        # 次の非flashモデル（2番目のpro）
        for m in GEMINI_MODELS[start_idx + 1:]:
            if 'flash' not in m.lower() and 'lite' not in m.lower():
                attempt_models.append(m)
                break
        # 最初のflashモデル（最終フォールバック）
        for m in GEMINI_MODELS:
            if 'flash' in m.lower() and 'lite' not in m.lower():
                attempt_models.append(m)
                break

    last_error = ""
    total_tries = 0

    for model_attempt, attempt_model in enumerate(attempt_models):
        attempt_url = (
            f"https://generativelanguage.googleapis.com/v1beta/models/"
            f"{attempt_model}:generateContent"
        )
        # モデルごとに thinkingConfig・タイムアウトを設定
        if model_attempt > 0:
            attempt_gc = dict(gen_config)
            tc_fb = _build_thinking_config(attempt_model, user_prompt)
            if tc_fb is not None:
                attempt_gc['thinkingConfig'] = tc_fb
            elif 'thinkingConfig' in attempt_gc:
                del attempt_gc['thinkingConfig']
            attempt_payload = {**payload, 'generationConfig': attempt_gc}
            attempt_timeout = 18 if 'flash' in attempt_model.lower() else req_timeout
            print(f"[Gemini] フォールバック: {current_model} → {attempt_model}")
        else:
            attempt_payload = payload
            attempt_timeout = req_timeout

        attempt_retries = 2

        for i in range(attempt_retries):
            # リトライ間隔: 1回目失敗後に3秒待つ（flashは1秒）
            wait = 1 if 'flash' in attempt_model.lower() else 3
            total_tries += 1
            try:
                response = requests.post(
                    attempt_url,
                    json=attempt_payload,
                    headers=_gemini_headers(),
                    timeout=attempt_timeout
                )

                # 503 → リトライ
                if response.status_code == 503:
                    last_error = f"503 Service Unavailable ({attempt_model} try {i+1}/{attempt_retries})"
                    print(f"[Gemini] {last_error} {wait}秒後に再試行")
                    if i < attempt_retries - 1:
                        time.sleep(wait)
                        continue
                    break

                # レスポンスをJSONとして解釈
                try:
                    gemini_res = response.json()
                except Exception as je:
                    last_error = f"JSONパース失敗: {je} / body先頭: {response.text[:200]}"
                    print(f"[Gemini] {last_error}")
                    if i < attempt_retries - 1:
                        time.sleep(wait)
                        continue
                    break

                # Gemini側のエラー構造（プロンプトブロックなど）
                if isinstance(gemini_res, dict) and gemini_res.get('error'):
                    err = gemini_res['error']
                    code = err.get('code')
                    status = err.get('status')
                    msg = err.get('message', '')
                    # 503/UNAVAILABLE はリトライ
                    if code == 503 or status == 'UNAVAILABLE':
                        last_error = f"{status or code}: {msg}"
                        print(f"[Gemini] {last_error} {wait}秒後に再試行 ({i+1}/{attempt_retries})")
                        if i < attempt_retries - 1:
                            time.sleep(wait)
                            continue
                        break
                    # それ以外の API エラーは即返す
                    print(f"[Gemini] APIエラー model={attempt_model}: {err}")
                    log_to_xlsx("DRAW_ERROR", prompt=user_prompt,
                                duration_sec=time.time() - _draw_t0,
                                canvas_w=w, canvas_h=h,
                                retries=total_tries,
                                model=attempt_model,
                                error=msg or str(err))
                    return jsonify({"error": f"[{attempt_model}] {msg or str(err)}"}), 500

                # HTTPエラーコード
                if not response.ok:
                    last_error = f"HTTP {response.status_code}: {response.text[:200]}"
                    print(f"[Gemini] {last_error}")
                    if i < attempt_retries - 1:
                        time.sleep(wait)
                        continue
                    break

                # parts を走査して thought: true を除外した text を拾う
                candidates = gemini_res.get('candidates', [])
                if not candidates:
                    last_error = f"候補なし: {gemini_res}"
                    print(f"[Gemini] {last_error}")
                    if i < attempt_retries - 1:
                        time.sleep(wait)
                        continue
                    break

                parts = candidates[0].get('content', {}).get('parts', []) or []
                text = ''
                for p in parts:
                    if isinstance(p, dict) and 'text' in p and not p.get('thought'):
                        text = p['text']
                        break

                # finishReason をログ（MAX_TOKENS / SAFETY などを検知）
                finish_reason = candidates[0].get('finishReason', '')
                print(f"[Gemini] model={attempt_model} status={response.status_code} "
                      f"finishReason={finish_reason} textLen={len(text)}")

                if not text:
                    last_error = f"空応答 finishReason={finish_reason}"
                    print(f"[Gemini] {last_error} full={gemini_res}")
                    if i < attempt_retries - 1:
                        time.sleep(wait)
                        continue
                    log_to_xlsx("DRAW_ERROR", prompt=user_prompt,
                                duration_sec=time.time() - _draw_t0,
                                canvas_w=w, canvas_h=h,
                                retries=total_tries,
                                model=attempt_model,
                                error=last_error,
                                detail="empty response")
                    return jsonify({"text": "[]", "warning": last_error})

                cleaned = _filter_cmds(text)
                raw_n = count_commands(text)
                filtered_n = count_commands(cleaned)
                removed_n = (raw_n - filtered_n) if raw_n is not None and filtered_n is not None else None
                usage = gemini_res.get('usageMetadata', {}) if isinstance(gemini_res, dict) else {}
                tok_in = usage.get('promptTokenCount')
                tok_out = usage.get('candidatesTokenCount')
                tok_thought = usage.get('thoughtsTokenCount', 0) or 0
                model_used = gemini_res.get('modelVersion') or attempt_model
                log_to_xlsx("DRAW_SUCCESS", prompt=user_prompt,
                            duration_sec=time.time() - _draw_t0,
                            canvas_w=w, canvas_h=h,
                            cmd_raw=raw_n,
                            cmd_filtered=filtered_n,
                            removed_bg=removed_n,
                            retries=total_tries,
                            model=model_used,
                            tokens_in=tok_in,
                            tokens_out=tok_out,
                            tokens_thought=tok_thought,
                            cost_jpy=estimate_cost_jpy(model_used, tok_in, tok_out, tok_thought),
                            detail=f"finishReason={finish_reason}")
                return jsonify({"text": cleaned})

            except requests.exceptions.RequestException as e:
                last_error = f"通信エラー: {e}"
                print(f"[Gemini] {last_error}")
                if i < attempt_retries - 1:
                    time.sleep(wait)
                    continue
                break
            except Exception as e:
                last_error = f"例外: {e}"
                print(f"[Gemini] {last_error}")
                traceback.print_exc()
                if i < attempt_retries - 1:
                    time.sleep(wait)
                    continue
                break
        # このモデルは全リトライ失敗 → 次のモデルへ

    # 全モデル・全リトライ後もダメ
    log_to_xlsx("DRAW_ERROR", prompt=user_prompt,
                duration_sec=time.time() - _draw_t0,
                canvas_w=w, canvas_h=h,
                retries=total_tries,
                model=current_model,
                error=last_error,
                detail="max retries exceeded")
    return jsonify({"error": f"APIが混雑しています。もう一度「かく」ボタンを押してください。\n({last_error})"}), 503

# ==========================================
# 印刷機能
# ==========================================
# A5縦 (300dpi想定) : 1748 x 2480 px
A5_WIDTH, A5_HEIGHT = 1748, 2480

# 画像エリアを「正方形・最大サイズ」に (A5 横幅 1748 を一辺とする)
IMAGE_AREA_SIZE = A5_WIDTH          # = 1748 (正方形)
IMAGE_AREA_PAD = 10                 # 画像エリアの外余白（極小）
TEXT_AREA_Y = IMAGE_AREA_SIZE       # 画像の下
TEXT_AREA_H = A5_HEIGHT - TEXT_AREA_Y  # = 2480 - 1748 = 732 px

MARGIN = 40                          # テキスト上下余白
FONT_TITLE_SIZE = 80                 # タイトルフォント（大きめに）

def load_font(size: int):
    candidates = ["C:/Windows/Fonts/msgothic.ttc", "C:/Windows/Fonts/meiryo.ttc"]
    for path in candidates:
        if os.path.exists(path): return ImageFont.truetype(path, size)
    return ImageFont.load_default()

def get_usb_printer():
    import win32print
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 2)
    for p in printers:
        if "USB" in p.get("pPortName", "").upper(): return p.get("pPrinterName")
    return win32print.GetDefaultPrinter()

def build_print_image(drawing_image, prompt_text):
    """
    A5縦 (1748x2480) のうち、
    ・上 1748x1748 → 絵を正方形で最大サイズに貼り付け
    ・下 1748x732  → タイトル（プロンプト）
    """
    page = Image.new("RGB", (A5_WIDTH, A5_HEIGHT), color=(255, 255, 255))
    draw = ImageDraw.Draw(page)

    # ── 絵を正方形エリア一杯にフィット（アスペクト比は維持） ──
    img = drawing_image.convert("RGB")
    img_w, img_h = img.size
    fit_box = IMAGE_AREA_SIZE - IMAGE_AREA_PAD * 2  # 1748 - 20 = 1728
    ratio = min(fit_box / img_w, fit_box / img_h)
    new_w = int(img_w * ratio)
    new_h = int(img_h * ratio)
    img_resized = img.resize((new_w, new_h), Image.LANCZOS)

    # 正方形エリアの中央に配置
    x_off = (IMAGE_AREA_SIZE - new_w) // 2
    y_off = (IMAGE_AREA_SIZE - new_h) // 2
    page.paste(img_resized, (x_off, y_off))

    # 区切り線
    draw.line([(0, TEXT_AREA_Y), (A5_WIDTH, TEXT_AREA_Y)], fill=(180, 180, 180), width=4)

    # ── テキストエリア（プロンプトをセンタリング） ──
    font_title = load_font(FONT_TITLE_SIZE)
    title_text = f"「{prompt_text}」" if prompt_text else "AI お絵描き作品"
    lines = textwrap.wrap(title_text, width=16) or [title_text]

    # 行ごとの高さを測ってセンタリング
    line_h_list = []
    for line in lines:
        bbox = draw.textbbox((0, 0), line, font=font_title)
        line_h_list.append(bbox[3] - bbox[1])
    total_h = sum(line_h_list) + (len(lines) - 1) * 20

    # テキスト領域内で縦中央
    curr_y = TEXT_AREA_Y + max(MARGIN, (TEXT_AREA_H - total_h) // 2)
    for line, lh in zip(lines, line_h_list):
        bbox = draw.textbbox((0, 0), line, font=font_title)
        text_w = bbox[2] - bbox[0]
        draw.text(((A5_WIDTH - text_w) // 2, curr_y), line, font=font_title, fill=(40, 40, 40))
        curr_y += lh + 20

    return page

@app.route('/print', methods=['POST'])
def print_image():
    if SERVER_MODE:
        return jsonify({"status": "error", "error": "サーバー版はブラウザ（Chrome）で印刷します"}), 400
    """印刷ジョブを送る。

    主な改善:
    - DEVMODEの書き換えは「できればやる」ベストエフォート方式（失敗してもクラッシュしない）
    - DEVMODE非対応のプリンターでも、用紙の中央にA5サイズで描画して印刷
    - すべての段階で例外をログに残す
    """
    _print_t0 = time.time()
    _print_prompt = ""
    _print_image_rel = ""
    try:
        data = request.json
        _print_prompt = data.get('prompt', '') if isinstance(data, dict) else ''
        try:
            _print_image_rel = save_drawing_image(data.get('image', ''), _print_prompt)
        except Exception as _e_si:
            print(f"[XLSX] print image save skipped: {_e_si}")
            _print_image_rel = ""
        img_data = base64.b64decode(data['image'].split(',')[1])
        drawing_image = Image.open(io.BytesIO(img_data)).convert("RGB")
        print_page = build_print_image(drawing_image, data.get('prompt', ''))

        import win32ui, win32con, win32print
        from PIL import ImageWin

        DMPAPER_A5 = 11
        DM_PAPERSIZE = 0x00000002

        try:
            printer_name = get_usb_printer()
        except Exception as e:
            print(f"[Print][WARN] プリンター取得失敗: {e} → デフォルトプリンターを使用")
            printer_name = win32print.GetDefaultPrinter()
        print(f"[Print] printer={printer_name}")

        # ── ① ベストエフォートでDEVMODEをA5に変更（失敗しても続行） ──
        try:
            h_printer = win32print.OpenPrinter(printer_name)
            try:
                level2 = win32print.GetPrinter(h_printer, 2)
                devmode = level2.get("pDevMode") if isinstance(level2, dict) else None
                if devmode is not None:
                    try:
                        devmode.PaperSize = DMPAPER_A5
                        try:
                            devmode.Fields = devmode.Fields | DM_PAPERSIZE
                        except Exception:
                            pass
                        print("[Print] DEVMODE PaperSize を A5 (11) にセット")
                    except Exception as e_dm:
                        print(f"[Print][WARN] DEVMODE書換不可: {e_dm}")
            finally:
                win32print.ClosePrinter(h_printer)
        except Exception as e_op:
            print(f"[Print][WARN] OpenPrinter失敗: {e_op}（DEVMODE変更スキップ）")

        # ── ② プリンタDCを作成 ──
        hDC = win32ui.CreateDC()
        try:
            hDC.CreatePrinterDC(printer_name)
        except Exception as e_dc:
            print(f"[Print][ERROR] CreatePrinterDC失敗: {e_dc}")
            return jsonify({"error": f"プリンターに接続できません: {e_dc}"}), 500

        try:
            hDC.SetMapMode(win32con.MM_TEXT)
            hDC.StartDoc("AI Print")
            hDC.StartPage()

            # ── ③ プリンタ実DPIでA5サイズ計算→用紙中央に描画 ──
            try:
                dpi_x = hDC.GetDeviceCaps(win32con.LOGPIXELSX) or 300
                dpi_y = hDC.GetDeviceCaps(win32con.LOGPIXELSY) or 300
                horz_res = hDC.GetDeviceCaps(win32con.HORZRES)
                vert_res = hDC.GetDeviceCaps(win32con.VERTRES)
            except Exception as e_caps:
                print(f"[Print][WARN] DeviceCaps取得失敗: {e_caps} → 既定値300dpiで進行")
                dpi_x = dpi_y = 300
                horz_res = vert_res = 0

            if horz_res > 0 and vert_res > 0:
                a5_w_px = int(round(148.0 / 25.4 * dpi_x))
                a5_h_px = int(round(210.0 / 25.4 * dpi_y))
                scale = min(horz_res / a5_w_px, vert_res / a5_h_px, 1.0)
                draw_w = int(a5_w_px * scale)
                draw_h = int(a5_h_px * scale)
                x0 = (horz_res - draw_w) // 2
                y0 = (vert_res - draw_h) // 2
                print(f"[Print] DPI=({dpi_x},{dpi_y}) HORZ={horz_res} VERT={vert_res} A5=({a5_w_px},{a5_h_px}) draw=({draw_w},{draw_h})")
                rect = (x0, y0, x0 + draw_w, y0 + draw_h)
            else:
                # フォールバック：プリンタ用紙いっぱいに描画
                horz_res = horz_res or 1748
                vert_res = vert_res or 2480
                rect = (0, 0, horz_res, vert_res)
                print(f"[Print] フォールバック: 用紙いっぱい描画 rect={rect}")

            ImageWin.Dib(print_page).draw(hDC.GetHandleOutput(), rect)
            hDC.EndPage()
            hDC.EndDoc()
        finally:
            try:
                hDC.DeleteDC()
            except Exception:
                pass

        log_to_xlsx("PRINT_SUCCESS", prompt=_print_prompt,
                    duration_sec=time.time() - _print_t0,
                    image_file=_print_image_rel)
        return jsonify({"status": "ok"})
    except Exception as e:
        traceback.print_exc()
        log_to_xlsx("PRINT_ERROR", prompt=_print_prompt,
                    duration_sec=time.time() - _print_t0,
                    error=str(e),
                    image_file=_print_image_rel)
        return jsonify({"error": str(e)}), 500

# ==========================================
# デスクトップアプリ起動
# ==========================================
APP_HOST = "127.0.0.1"
APP_PORT = 5000
APP_URL = f"http://{APP_HOST}:{APP_PORT}/"
ADMIN_URL = f"http://{APP_HOST}:{APP_PORT}/admin.html"

# 起動方法：browser = Chrome のキオスク（全画面ロック）/ window = 従来のアプリウィンドウ(pywebview)
#   .env の LAUNCH_MODE か、起動引数 --browser / --window で切り替え
def _launch_mode() -> str:
    args = [a.lower() for a in sys.argv[1:]]
    if "--window" in args:
        return "window"
    if "--browser" in args:
        return "browser"
    v = (os.environ.get("LAUNCH_MODE", "browser") or "browser").strip().lower()
    return "window" if v == "window" else "browser"

def _env_flag(name: str, default: str = "1") -> bool:
    return (os.environ.get(name, default) or "").strip().lower() not in ("0", "false", "no", "off")

LAUNCH_MODE = _launch_mode()   # 今回の起動方法（引数 > .env）
LAUNCH_MODE_SETTING = "window" if (os.environ.get("LAUNCH_MODE", "browser") or "").strip().lower() == "window" else "browser"
CHROME_PATH = (os.environ.get("CHROME_PATH", "") or "").strip()
KIOSK_BLOCK_KEYS = _env_flag("KIOSK_BLOCK_KEYS", "1")
# Chrome 専用プロファイル（普段使いの Chrome とは別。デスクトップを散らかさないよう AppData に置く）
KIOSK_PROFILE_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA") or exe_dir, "AIDrawing", "chrome-profile")

_kiosk = None         # kiosk.KioskBrowser
_keyblocker = None    # kiosk.KeyboardBlocker

def _shutdown_everything():
    """管理画面の「AIソフトを閉じる」：キオスクブラウザも閉じてからアプリ終了。"""
    try:
        if _keyblocker is not None:
            _keyblocker.stop()
        if _kiosk is not None:
            _kiosk.close()
    except Exception as e:
        print(f"[shutdown] {e}")
    os._exit(0)

def _startup_delay():
    """--delay=秒：Windows 起動直後はネットやデスクトップの準備待ちに少し待つ。"""
    for a in sys.argv[1:]:
        if a.lower().startswith("--delay="):
            try:
                sec = max(0, min(120, int(a.split("=", 1)[1])))
                print(f"[startup] {sec} 秒待ってから起動します")
                time.sleep(sec)
            except ValueError:
                pass

def run_browser_mode():
    """Chrome をキオスク（全画面ロック）で開き、サーバーを動かし続ける。"""
    global _kiosk, _keyblocker
    _kiosk = kiosk.KioskBrowser(APP_URL, KIOSK_PROFILE_DIR, CHROME_PATH,
                                on_give_up=lambda: os._exit(1))
    if not _kiosk.start():
        _fatal_box("AI お絵描き",
                   "Google Chrome が見つかりませんでした。\n"
                   "Chrome をインストールするか、.env に CHROME_PATH=chrome.exeのパス を書いてください。")
        os._exit(1)
    if KIOSK_BLOCK_KEYS:
        _keyblocker = kiosk.KeyboardBlocker(lambda: _kiosk.pid if _kiosk else None)
        _keyblocker.start()
    print(f"[OK] ブラウザ（キオスク）で起動しました: {APP_URL}")

def start_server():
    """Flask サーバーをバックグラウンドで起動。"""
    _load_xlsx_history_into_buffer()
    app.run(host=APP_HOST, port=APP_PORT, debug=False, use_reloader=False, threaded=True)

def wait_until_server_ready(timeout=15):
    """サーバーが応答するまで待つ（最大 timeout 秒）。"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            r = requests.get(APP_URL, timeout=1)
            if r.status_code in (200, 304):
                return True
        except Exception:
            pass
        time.sleep(0.2)
    return False

def _wants_fullscreen() -> bool:
    """全画面起動を要求されているかをチェック。
    1) コマンドライン引数 --fullscreen / -f / --kiosk
    2) .env の FULLSCREEN=0 / false / no / off で解除
    """
    args = [a.lower() for a in sys.argv[1:]]
    if any(a in ('--fullscreen', '-f', '--kiosk', '/f') for a in args):
        return True
    val = (os.environ.get('FULLSCREEN', '1') or '').strip().lower()
    return val not in ('0', 'false', 'no', 'off')

def open_native_window():
    """pywebview でネイティブウィンドウを開く。失敗時は Edge --app にフォールバック。"""
    fullscreen = _wants_fullscreen()

    # ★WebView2 / Edge にマイク使用許可ダイアログを出させない設定
    #   - --use-fake-ui-for-media-stream : メディアアクセス許可ダイアログを抑止し常に許可
    #   - --autoplay-policy=no-user-gesture-required : 念のため自動再生もブロックしない
    media_flags = '--use-fake-ui-for-media-stream --autoplay-policy=no-user-gesture-required'
    # WebView2 はこの環境変数を読んでブラウザ引数を追加する
    os.environ['WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS'] = media_flags

    try:
        import webview  # pywebview
    except ImportError:
        print("[WARN] pywebview が見つからないため Edge --app で起動します")
        try:
            edge_args = [
                'start', 'msedge',
                '--app=' + APP_URL,
                '--use-fake-ui-for-media-stream',
                '--autoplay-policy=no-user-gesture-required',
            ]
            if fullscreen:
                # Edge アプリモードでも --start-fullscreen で全画面
                edge_args.insert(3, '--start-fullscreen')
            subprocess.Popen(edge_args, shell=True)
        except Exception as e:
            print(f"[ERROR] Edge 起動失敗: {e}")
        # Flask サーバーを生かし続けるため、ここではブロックしない
        return False

    # アイコンを探す（任意）
    icon_path = None
    for cand in ("icon.ico", "icon.png"):
        p = os.path.join(exe_dir, cand)
        if os.path.exists(p):
            icon_path = p
            break

    window = webview.create_window(
        title=APP_WINDOW_TITLE,
        url=APP_URL,
        width=1280,
        height=860,
        min_size=(900, 640),
        resizable=True,
        fullscreen=fullscreen,   # ★全画面で起動
        confirm_close=False,
        background_color='#fff5e6',
    )
    admin_state = {"window": None}

    # F11 で全画面トグル / Esc で解除する仕組みを JS 側から
    # window.toggle_fullscreen() を呼べるようにする
    def _on_loaded():
        try:
            window.evaluate_js("""
                document.addEventListener('keydown', function(e){
                    if (e.key === 'F11') {
                        e.preventDefault();
                        if (window.pywebview && window.pywebview.api && window.pywebview.api.toggle_fs) {
                            window.pywebview.api.toggle_fs();
                        }
                    }
                });
            """)
        except Exception:
            pass

    class _AdminApi:
        """管理者ウィンドウから window.pywebview.api.xxx() で呼ばれる"""

        def get_settings(self):
            with _settings_lock:
                return {
                    "model": GEMINI_MODEL,
                    "models": GEMINI_MODELS,
                    **app_settings,
                }

        def set_model(self, model):
            global GEMINI_MODEL
            model = str(model).strip()
            if not model:
                return {"error": "model is empty"}
            with _settings_lock:
                if model not in GEMINI_MODELS:
                    GEMINI_MODELS.append(model)
                GEMINI_MODEL = model
                _save_env_value("GEMINI_MODEL", model)
            return {"status": "ok", "model": GEMINI_MODEL}

        def set_bgm(self, enabled, volume):
            with _settings_lock:
                app_settings["bgm_enabled"] = bool(enabled)
                app_settings["bgm_volume"] = max(0, min(100, int(volume)))
            return {"status": "ok"}

        def set_speed(self, speed):
            with _settings_lock:
                app_settings["draw_speed"] = max(0, min(4, int(speed)))
            return {"status": "ok"}

        def set_temperature(self, temp):
            with _settings_lock:
                app_settings["ai_temperature"] = max(0.1, min(2.0, float(temp)))
            return {"status": "ok"}

        def get_logs(self):
            with _recent_logs_lock:
                return list(_recent_logs)

        def open_debug(self):
            subprocess.Popen(
                ["cmd", "/c", "start", "", "http://127.0.0.1:5000/debug"],
                shell=False,
            )
            return {"status": "ok"}

        def shutdown(self):
            def _quit():
                time.sleep(0.25)
                os._exit(0)
            threading.Thread(target=_quit, daemon=True).start()
            return {"status": "closing"}

    class _Api:
        def toggle_fs(self):
            try:
                window.toggle_fullscreen()
            except Exception:
                pass

        def open_admin(self, token=""):
            try:
                if admin_state["window"] is not None:
                    return
                w = webview.create_window(
                    title="\u7ba1\u7406\u8005\u753b\u9762",
                    url=ADMIN_URL + ("?token=" + token if token else ""),
                    width=480,
                    height=700,
                    resizable=True,
                    confirm_close=False,
                )
                admin_state["window"] = w
                w.events.closed += lambda: admin_state.update({"window": None})
                # pywebview API \u3092\u7ba1\u7406\u8005\u30a6\u30a3\u30f3\u30c9\u30a6\u306b\u516c\u958b
                adm = _AdminApi()
                w.expose(
                    adm.get_settings,
                    adm.set_model,
                    adm.set_bgm,
                    adm.set_speed,
                    adm.set_temperature,
                    adm.get_logs,
                    adm.open_debug,
                    adm.shutdown,
                )
            except Exception:
                subprocess.Popen(
                    ["cmd", "/c", "start", "", "msedge", "--app=" + ADMIN_URL, "--window-size=480,700"],
                    shell=False,
                )

    api = _Api()
    window.expose(api.toggle_fs, api.open_admin)
    window.events.loaded += _on_loaded

    # webview.start() はブロッキング。Windows では EdgeChromium を優先的に使う。
    try:
        webview.start(gui='edgechromium')
    except Exception:
        # GUI 指定が効かない環境は自動選択にフォールバック
        webview.start()
    return True

def _fatal_box(title: str, msg: str):
    """致命的エラーをメッセージボックスで表示（--windowed でも見える）。"""
    print(f"[{title}] {msg}")
    try:
        ctypes.windll.user32.MessageBoxW(0, msg, title, 0x10)  # MB_ICONERROR
    except Exception:
        pass

# gunicorn など WSGI サーバーから読み込まれたとき（本番）も、過去ログを読み込んでおく
if SERVER_MODE and __name__ != "__main__":
    try:
        _load_xlsx_history_into_buffer()
    except Exception as _e:
        print(f"[XLSX] 履歴の読み込みをスキップ: {_e}")

def run_server_mode():
    """サーバー公開モード（Mac での開発確認用。本番は gunicorn で起動）。"""
    _load_xlsx_history_into_buffer()
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "8080"))
    print(f"🎨 AIお絵描き サーバーモード: http://{'localhost' if host in ('127.0.0.1', '0.0.0.0') else host}:{port}/")
    print(f"[data] {DATA_DIR}")
    print(f"[APIキー] {'設定OK' if API_KEY else '★未設定★（管理画面の環境設定から入力できます）'}")
    print(f"[ログイン] {'ID/パスワード設定済み' if _app_password() else 'APP_PASSWORD 未設定（localhost からのみ利用可）'}")
    app.run(host=host, port=port, debug=False, use_reloader=False, threaded=True)

def _main():
    args = [a.lower() for a in sys.argv[1:]]
    if any(a in ('--debug-only', '--debug-server') for a in args):
        start_server()
        return
    if SERVER_MODE:
        run_server_mode()
        return

    # ★最優先：多重起動チェック。既に起動中ならここで終了する
    ensure_single_instance()
    _startup_delay()

    # 念のため、ミューテックスが効かない環境用にポートも見ておく
    if is_port_in_use(APP_HOST, APP_PORT):
        # ポートが塞がっている = 既存インスタンス or 別アプリが使用中
        _bring_existing_window_to_front()
        print(f"[ERROR] ポート {APP_PORT} が既に使われています。"
              + ("（Mac は「AirPlayレシーバー」が 5000 番を使うので OFF にしてください）" if sys.platform == "darwin" else ""))
        try:
            ctypes.windll.user32.MessageBoxW(
                0,
                f"ポート {APP_PORT} が既に使われています。\n"
                "AI お絵描きがすでに起動していないか確認してください。",
                "AI お絵描き",
                0x30,  # MB_ICONWARNING
            )
        except Exception:
            pass
        sys.exit(0)

    print("🎨 AIお絵描き 起動中...")
    # 起動時チェック（原因切り分け用）
    env_path = os.path.join(DATA_DIR, ".env")
    print(f"[.env] path={env_path} exists={os.path.exists(env_path)}")
    if API_KEY:
        print(f"[APIキー] 設定OK (先頭4文字: {API_KEY[:4]}...)")
    else:
        print("[APIキー] ★未設定★ 管理画面（タイトル長押し→暗証番号）の「環境設定」から入力してください")
    print(f"[index.html] path={os.path.join(base_dir, 'index.html')} exists={os.path.exists(os.path.join(base_dir, 'index.html'))}")

    # 1) Flask を別スレッドで起動
    server_thread = threading.Thread(target=start_server, daemon=True)
    server_thread.start()

    # 2) サーバーが応答可能になるまで待機
    if not wait_until_server_ready():
        print("[ERROR] サーバーが起動しませんでした (タイムアウト)")
        sys.exit(1)
    print(f"[OK] サーバー応答確認: {APP_URL}")

    # 3-a) ブラウザ（Chrome キオスク）モード
    if LAUNCH_MODE == "browser":
        run_browser_mode()
        try:
            server_thread.join()
        except KeyboardInterrupt:
            _shutdown_everything()
        return

    # 3-b) ネイティブウィンドウを開く（閉じられるまでブロック）
    used_pywebview = open_native_window()

    if used_pywebview:
        # pywebview ウィンドウが閉じられたらアプリ全体を終了する
        print("ウィンドウが閉じられました。終了します。")
        os._exit(0)
    else:
        # フォールバック時はサーバーを生かす
        print("Edge アプリモードで起動しました。サーバーは Ctrl+C で停止します。")
        try:
            server_thread.join()
        except KeyboardInterrupt:
            os._exit(0)


if __name__ == '__main__':
    # --windowed ビルドだと print された例外がどこにも出ないので、
    # ここで全例外を捕まえて MessageBox とログファイルに残す。
    try:
        _main()
    except SystemExit:
        raise
    except BaseException:
        tb = traceback.format_exc()
        # ログにも書く
        try:
            print("[FATAL]\n" + tb)
        except Exception:
            pass
        try:
            log_path = os.path.join(exe_dir, "logs", "fatal.log")
            os.makedirs(os.path.dirname(log_path), exist_ok=True)
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(f"\n----- {datetime.now().isoformat()} -----\n{tb}\n")
        except Exception:
            pass
        # 画面にも見せる
        _fatal_box(
            "AI お絵描き — 起動エラー",
            "起動中にエラーが発生しました。\n\n"
            f"{tb[-1200:]}\n\n"
            f"詳細は logs フォルダの fatal.log を確認してください。"
        )
        sys.exit(1)
