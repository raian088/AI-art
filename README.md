# 🎨 AI お絵描き（AI-art）

子どもが言った・入力したお題を、AI（Google Gemini）が一筆ずつアニメーションで描くアプリです。
Windows PC で Google Chrome を**全画面ロック（キオスク）**で開いて動かします。

## フォルダ構成

```
AI-art/
├── desktopaap.py        # 本体（Flask サーバー + 起動処理）
├── kiosk.py             # Chrome キオスク起動・キー制限
├── index.html           # お絵描き画面
├── admin.html           # 管理者画面
├── assets/              # ロゴ画像
├── BGM/                 # BGM（BGM2.mp3 = 完成音、それ以外は考え中にランダム再生）
├── build.bat            # exe 化（dist\AIDrawing\ に出力）
├── setup/               # Windows へのインストール・自動起動の登録
│   ├── install.bat
│   └── uninstall.bat
├── tools/
│   ├── open_debug_log.bat   # デバッグログ画面を開く（Windows）
│   └── start_mac.command    # Mac 用の起動スクリプト
├── docs/
│   └── LICENSE_AUDIT.md     # ライブラリの商用利用可否の調査
└── .env.example         # 設定ファイルの見本（コピーして .env にする）
```

## Windows PC へのセットアップ

前提: Windows 10/11、Google Chrome、Python 3.12（ビルドする PC のみ）

1. **APIキーを設定**: `.env.example` をコピーして `.env` にし、`GEMINI_API_KEY=` にキーを書く
2. **ビルド**: `build.bat` をダブルクリック → `dist\AIDrawing\AIDrawing.exe` ができる
3. **インストール**: `dist\AIDrawing\setup\install.bat` をダブルクリック
   - アプリ一式をデスクトップの `AIDrawing` フォルダに配置
   - デスクトップにショートカット「AIお絵描き」を作成
   - 「PC起動時に自動で全画面起動しますか？」→ Y でスタートアップに登録
   - （任意）スリープ・画面オフの無効化

別の PC に入れるときは、`dist\AIDrawing` フォルダごとコピーして `setup\install.bat` を実行するだけです（Python は不要）。

### 自動起動の ON / OFF

Windows の **設定 > アプリ > スタートアップ** に「AIお絵描き」が表示され、そこで切り替えできます。
完全に解除するときは `setup\uninstall.bat` を実行します。

> サインイン画面で止まると自動起動されません。展示用 PC は自動サインインの設定もおすすめです。

## 画面ロック（キオスク）

起動すると Chrome が全画面で開き、アドレスバーやタブは表示されません。

| 操作 | ロック中 |
|---|---|
| アプリのボタン・お題の入力・マイク | ✅ 使える |
| 右クリック・文字選択・ドラッグ・拡大（ピンチ / Ctrl+ホイール） | ❌ 無効 |
| 更新（F5）・戻る・印刷・検索・開発者ツールなどのブラウザ操作 | ❌ 無効 |
| Alt+Tab / Alt+F4 / Alt+Esc / Ctrl+W / Ctrl+N / Ctrl+Shift+Esc | ❌ 無効 |
| Windows キー | ✅ 使える（指定どおり残しています） |
| Ctrl+Alt+Del | ✅ 使える（Windows の仕様上ブロック不可） |

- Chrome が落ちたり閉じられたりしても、自動で開き直します
- Chrome は普段使いとは別の専用プロファイル（`%LOCALAPPDATA%\AIDrawing\chrome-profile`）で起動します

### 管理画面・終了のしかた

1. 画面上部のタイトル「🎨 AI お絵描き ✨」を **2秒長押し**
2. 暗証番号を入力（初期値 **0001**）
3. 管理画面が開く → BGM・スピード・AIモデル・ログ確認・デバッグ・**「AIソフトを閉じる」**

管理画面は 3 分間さわらないと自動で閉じます。暗証番号を 5 回間違えると 30 秒ロックされます。

## 設定（.env）

| 項目 | 初期値 | 内容 |
|---|---|---|
| `GEMINI_API_KEY` | — | Gemini API キー（必須） |
| `GEMINI_MODEL` | — | 使うモデル（管理画面からも変更可） |
| `LAUNCH_MODE` | `browser` | `browser` = Chrome キオスク / `window` = 従来のアプリウィンドウ |
| `ADMIN_PIN` | `0001` | 管理画面の暗証番号 |
| `KIOSK_BLOCK_KEYS` | `1` | `0` にすると Alt+Tab などのキー制限をしない |
| `CHROME_PATH` | 自動検出 | Chrome が標準の場所にないときだけ指定 |

起動引数 `--browser` / `--window` で一時的に切り替えもできます。`--delay=10` で起動を10秒遅らせます（自動起動で使用）。

## Mac での起動

Mac でも動かせます（展示用の本番は Windows を想定しています）。

1. Google Chrome と Python 3（https://www.python.org/ ）をインストール
2. **AirPlay レシーバーを OFF**：システム設定 > 一般 > AirDrop と Handoff > AirPlay レシーバー
   （Mac はこの機能が 5000 番ポートを使うため、ONのままだと起動できません）
3. `tools/start_mac.command` をダブルクリック
   - 初回は `.env` が作られるので APIキーを書いて保存 → もう一度ダブルクリック
   - 初回はライブラリのインストールに少し時間がかかります
   - 「開発元を確認できません」と出たら、右クリック >「開く」で起動

Mac での違い：

- **印刷はできません**（Windows のプリンター機能を使っているため）
- Cmd+Tab・Cmd+Q などのキー制限はありません（Cmd+Q で閉じても Chrome は自動で開き直します）
- 終了は管理画面の「AIソフトを閉じる」、またはターミナルで Ctrl+C

## 開発時の起動

```bat
pip install flask flask-cors requests python-dotenv pillow pywin32 openpyxl pywebview
python desktopaap.py
```

ブラウザで http://127.0.0.1:5000/ を開いても使えます。

## 注意

- `.env`（APIキー）・`logs/`・`debug.xlsx`・`ratings.json` は **Git に入れないでください**（`.gitignore` で除外済み）。ログにはお客様の入力したお題や絵が含まれます。
- 子どもの音声・入力を扱うため、Gemini API は有料枠の利用を推奨します（詳しくは `docs/LICENSE_AUDIT.md`）。
