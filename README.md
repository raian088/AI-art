# 🎨 AI お絵描き（AI-art）

子どもが言った・入力したお題を、AI（Google Gemini）が一筆ずつアニメーションで描くアプリです。

**運用の流れ**

```
 開発（Mac）          公開（サーバー）               展示（Windows PC）
 修正・動作確認  ──▶  GitHub に push で自動反映  ──▶  Chrome 全画面で URL を開くだけ
 tools/dev_mac        Render など + 独自ドメイン        windows/setup.bat
```

- APIキーはサーバーにだけ置くので、展示 PC にキーは入りません
- 修正を push すれば、すべての展示 PC が次に開いたときから新しい版になります（PC 側の作業不要）
- 展示 PC の設定・APIキー・暗証番号は、展示 PC の画面（管理画面）から変更できます

## フォルダ構成

```
AI-art/
├── desktopaap.py        # 本体（Flask サーバー）
├── index.html           # お絵描き画面
├── admin.html           # 管理者画面
├── login.html           # ログイン画面（サーバー版）
├── kiosk.py             # PC 版（exe）用：Chrome 全画面起動・キー制限
├── assets/  BGM/        # ロゴ画像・BGM（BGM2.mp3 = 完成音）
├── Dockerfile  requirements.txt  render.yaml   # サーバー公開用
├── tools/dev_mac.command    # Mac で開発・動作確認
├── windows/                 # 展示用 Windows PC の設定（サーバー版）
├── offline/  build.bat      # PC 版（exe・オフライン）
├── docs/LICENSE_AUDIT.md    # ライブラリの商用利用可否の調査
└── .env.example             # 設定項目の見本
```

---

## 1. Mac で開発する

前提: Python 3（https://www.python.org/ ）

1. `tools/dev_mac.command` をダブルクリック（「開発元を確認できません」と出たら右クリック →「開く」）
2. ブラウザで http://localhost:8080/ が開きます
3. APIキーは、起動画面の「⚙ APIキーを設定する」→ 暗証番号 0001 → 入力（`data/.env` に保存）
4. ファイルを修正したら、ターミナルで Ctrl+C → もう一度ダブルクリック

Mac の開発環境ではログインなしで使えます（`localhost` のときだけ）。印刷ボタンを押すと Mac の印刷ダイアログが出ます。

## 2. サーバーに公開する（ドメイン）

おすすめは **Render**（GitHub につなぐだけで、push のたびに自動で公開。https も自動）。

1. https://render.com にログイン → **New > Blueprint** → この GitHub リポジトリ（raian088/AI-art）を選ぶ
   - `render.yaml` の設定（Docker・データ保存用ディスク 1GB・Starter プラン）が読み込まれます
2. 環境変数を入力
   - `APP_USER` / `APP_PASSWORD` … 展示 PC の Chrome でログインする ID・パスワード（**必須**）
   - `GEMINI_API_KEY` … 空でもOK（公開後に管理画面から入力できます）
3. デプロイが終わると `https://ai-art-xxxx.onrender.com` で開けます
4. 独自ドメイン：Render の **Settings > Custom Domains** でドメインを追加 → 表示された CNAME をドメイン管理画面（お名前.com など）に設定

> 費用の目安：Render Starter 約 $7/月 + ディスク 1GB 約 $0.25/月。無料プランはディスクが使えず、15分使わないとスリープして起動に時間がかかるので展示には向きません。

**他のサービスに置く場合**：`Dockerfile` があるので、Docker が動く所ならどこでも動きます。

- 環境変数 `APP_MODE=server`、`DATA_DIR=/data`、`APP_USER`、`APP_PASSWORD` を設定し、`/data` に永続ディスクをつなぐ
- Google Cloud Run など永続ディスクが無いサービスでは、管理画面で変えた設定（APIキー・暗証番号など）が再起動で消えます。その場合は環境変数で設定してください

**更新のしかた**：Mac で修正 → `git push` → Render が自動で公開（数分）。展示 PC は「最初から」ボタンか再起動で新しい版になります。

## 3. 展示用 Windows PC を設定する

前提: Windows 10/11、Google Chrome、プリンター（印刷する場合）

1. `windows` フォルダを PC にコピー（USB メモリなど）
   - `windows/kiosk_url.txt` に公開した URL を書いておくと、入力の手間が省けます
2. `windows/setup.bat` をダブルクリック
   - URL を入力（kiosk_url.txt に書いてあれば不要）
   - デスクトップにショートカット「AIお絵描き」を作成
   - 「PC起動時に自動で全画面起動しますか？」→ Y でスタートアップに登録
3. 「今すぐ起動しますか？」→ Y → 全画面でログイン画面が出るので ID・パスワードを入力（初回だけ。1年間ログインしたまま）

| やりたいこと | 方法 |
|---|---|
| 起動 | デスクトップの「AIお絵描き」 |
| 閉じる | **Alt+F4** |
| 管理画面（BGM・AIモデル・APIキー・暗証番号など） | タイトルを **2秒長押し** → 暗証番号（初期値 **0001**） |
| 自動起動の ON/OFF | Windows の 設定 > アプリ > スタートアップ |
| 設定を元に戻す | `windows/uninstall.bat` |

- PC 起動時はサーバーにつながるまで待ってから Chrome を開きます（最大90秒）
- Chrome は普段使いとは別の専用プロファイルで開きます

### 印刷

「印刷する」ボタンで、**確認ダイアログなしで Windows の既定のプリンター**に印刷します（A5 縦：上に絵、下にお題）。

- 設定 > Bluetooth とデバイス > プリンターとスキャナー で、使うプリンターを「既定」に
- 同じ画面の「Windows で通常使うプリンターを管理する」は **オフ**（勝手に既定が変わるのを防ぐ）
- プリンターの用紙サイズを **A5** に

### 画面ロック

全画面でアドレスバー・タブは表示されません。ページ内では右クリック・文字選択・拡大・更新（F5）・戻る などを無効にしています。
Alt+F4（閉じる）と Windows キーは使えます。

---

## 管理画面

タイトル「🎨 AI お絵描き ✨」を 2秒長押し → 暗証番号。3分さわらないと自動で閉じ、暗証番号を5回間違えると30秒ロックされます。

- BGM・描画スピード・AIモデル・個性（温度）
- **🔑 環境設定**：Gemini APIキー（接続テスト付き）・暗証番号・ログインID/パスワード
- 最近のログ・デバッグ画面

サーバー版の設定は全 PC 共通です（1台で変えると全部に反映）。

## 設定項目（.env / 環境変数）

| 項目 | 初期値 | 内容 |
|---|---|---|
| `GEMINI_API_KEY` | — | Gemini API キー |
| `GEMINI_MODEL` | — | 使うモデル |
| `ADMIN_PIN` | `0001` | 管理画面の暗証番号 |
| `APP_USER` / `APP_PASSWORD` | — | サーバー版のログイン（公開時は必須） |
| `APP_MODE` | — | `server` でサーバー版 |
| `DATA_DIR` | — | 設定・ログ・描いた絵の保存先（サーバーでは永続ディスク） |
| `SECRET_KEY` | 自動生成 | ログイン状態の暗号化キー |

## PC 版（exe・オフライン）

サーバーを使わず、PC 1台で完結させる従来の方法も残しています（Windows・ネット接続は Gemini のみ必要）。

1. `.env.example` をコピーして `.env` に（APIキーは後から管理画面でも入力可）
2. `build.bat` → `dist\AIDrawing\AIDrawing.exe`
3. `dist\AIDrawing\setup\install.bat` → デスクトップ配置・自動起動
   - Alt+Tab・Alt+F4 などのキー制限、USB プリンターへの直接印刷つき
   - 終了は管理画面の「AIソフトを閉じる」

## 注意

- `.env`・`data/`・`logs/`・`debug.xlsx` は **Git に入れないでください**（`.gitignore` で除外済み）。お客様が入力したお題や絵が含まれます
- 子どもの音声・入力を扱うため、Gemini API は有料枠の利用を推奨します（`docs/LICENSE_AUDIT.md`）
- マイクは https の URL でないと使えません
