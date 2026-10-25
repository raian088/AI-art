# AI お絵描き — ライブラリ商用利用可否調査レポート

調査日: 2026-05-04

> **2026-10-25 追記**: README に記載されていた APIキーは削除済み（リポジトリには含めていません）。旧キーは Google AI Studio で失効させてください。README の「Anthropic」表記も Gemini に修正済み。

対象: `index.html` / `server.js` / `app.py` / `desktopaap.py` / `package.json` / `*.spec` / `build.bat`

---

## 結論

**商用利用「不可」のライブラリは見つかりませんでした。** 全て商用利用可能なライセンス（MIT / BSD / Apache 2.0 / SIL OFL / GPL with exception）です。
ただし、いくつか **注意・確認が必要な項目** があります（後述）。

---

## 1. Python 依存（`app.py` / `desktopaap.py` / `build.bat` で `pip install` されるもの）

| ライブラリ | ライセンス | 商用利用 | 備考 |
|---|---|---|---|
| Flask | BSD-3-Clause | ✅ 可 | 表示義務あり（LICENSEファイル同梱推奨） |
| Flask-CORS | MIT | ✅ 可 | |
| Werkzeug / Jinja2 / Click / itsdangerous / MarkupSafe | BSD-3-Clause | ✅ 可 | Flaskの依存 |
| blinker | MIT | ✅ 可 | |
| Pillow (PIL) | MIT-CMU (HPND) | ✅ 可 | |
| python-dotenv | BSD-3-Clause | ✅ 可 | |
| requests | Apache-2.0 | ✅ 可 | NOTICEファイルがあれば同梱 |
| urllib3 | MIT | ✅ 可 | |
| certifi | MPL-2.0 | ✅ 可 | ファイル単位コピーレフト。改変しなければ問題なし |
| idna | BSD-3-Clause | ✅ 可 | |
| charset-normalizer | MIT | ✅ 可 | |
| **pywin32** (win32print/ui/con/gui, pywintypes) | PSF-2.0 | ✅ 可 | Python Software Foundationライセンス |
| **pywebview** | BSD-3-Clause | ✅ 可 | |
| pythonnet / clr_loader | MIT | ✅ 可 | |
| **PyInstaller** | GPL-2.0+ ( **bootloader例外あり** ) | ⚠️ 条件付き可 | ★下記参照 |

### ⚠️ PyInstaller の注意点（重要）

PyInstaller本体はGPL-2.0+ですが、「ブートローダー例外条項」があり、**PyInstallerでバンドルしたexeはGPLにする必要はありません。** 商用クローズドソース配布は問題ありません。
公式ガイド: <https://pyinstaller.org/en/stable/license.html>

ただし、PyInstallerの **ソースコード自体を改変して再配布** する場合はGPLが適用されます。本プロジェクトは標準のpipインストール版を使う想定なので問題なし。

### ⚠️ UPX の注意点

`.spec` ファイルに `upx=True` の指定があります。UPXはGPL-2.0+ですが「stub linking exception」があり、**圧縮されたバイナリ自体はGPLになりません。** 商用利用OK。
公式: <https://upx.github.io/upx-license.html>

---

## 2. Node.js 依存（`package.json`）

| パッケージ | バージョン | ライセンス | 商用利用 |
|---|---|---|---|
| dotenv | ^17.4.0 | BSD-2-Clause | ✅ 可 |

`node_modules/` を確認したところ、依存は `dotenv` のみでした。問題なし。

---

## 3. フロントエンド（`index.html`）の外部リソース

| 項目 | ライセンス | 商用利用 | 備考 |
|---|---|---|---|
| **Google Fonts: Zen Maru Gothic** | SIL Open Font License 1.1 | ✅ 可 | 商用Webサイト・アプリ・印刷物すべてOK。フォント単体販売のみNG |
| HTML5 Canvas API | 標準仕様 | ✅ 可 | |
| Fetch API | 標準仕様 | ✅ 可 | |
| Web Speech API (`webkitSpeechRecognition`) | ブラウザ実装 | ⚠️ ★下記 | |

### ⚠️ Web Speech API（音声認識）の注意点

Chromeの `webkitSpeechRecognition` は **音声データをGoogleのサーバーに送信して認識** しています。
- ライセンス上の制限はないが、**Googleの利用規約に従う必要** がある
- W3C正式仕様ではなくChrome独自実装に近く、**プロダクション用途として明文化された保証はない**
- 子ども（未成年）の音声を扱うため、**プライバシーポリシー・保護者同意** の整備推奨
- 商用イベント等で使う場合は、ユーザー（保護者）へ「音声がGoogleに送信される」ことを告知することを推奨

---

## 4. 外部APIサービス

| サービス | 商用利用 | 注意点 |
|---|---|---|
| **Gemini API** (`generativelanguage.googleapis.com`) | ✅ 可（有料枠） | ★下記 |

### ⚠️ Gemini API の注意点（最重要）

`server.js` で使用中のモデル: `gemini-3.0-flash`（※実際には `gemini-3-flash-preview` 等のモデル名要確認）

- **無料枠（Free tier）**: 入力データがGoogleのモデル改善に使用される可能性あり。**子どもの入力（プロンプト）が学習データに使われる懸念** があるため、商用イベント・継続運用では推奨しない
- **有料枠（Paid tier / Pay-as-you-go）**: データは学習に使用されない。商用利用OK
- 生成された絵（描画コマンド経由）の **著作権・利用規約** は Google AI の使用許諾に従う
- 公式利用規約: <https://ai.google.dev/gemini-api/terms>

**推奨**: 商用イベント等で使うなら **必ず有料枠（請求先設定済みプロジェクト）** を使ってください。

---

## 5. その他見つけた懸念事項（ライセンス外）

### `.env` のAPIキーが README に記載されている

`README.md` 内に以下のような記述があります:
```
$env:GEMINI_API_KEY="AIzaSy********************************"
```

**これは実際に使えるAPIキーが含まれているように見えます。** すぐに以下の対応をしてください:
1. Google AI StudioでこのAPIキーを **失効（revoke）**
2. 新しいキーを発行して `.env` のみに保存
3. README から実キーの記述を削除（プレースホルダ `AIzaXXXX...` に置換）
4. Gitリポジトリに含まれている場合は履歴からも削除（`git filter-repo` 等）

### README と実装の不一致

READMEは「Anthropic APIキー」と書かれていますが、実装は **Gemini API** を使っています。配布前に修正推奨。

### `.env` のパーミッション

調査した範囲では `.env` の中身は読み取れませんでしたが、配布する `dist/AIDrawing/` フォルダに `.env` を含めると **APIキーが利用者に渡る** ことになります。商用配布する場合は、APIキーをサーバー側に置く構成（プロキシ）への変更を強く推奨。

---

## 6. 商用化チェックリスト（推奨アクション）

- [ ] Gemini API を **有料枠** に切り替え（請求先プロジェクトを使う）
- [ ] README に書かれている古いAPIキーを **revoke** し、新しいキーを発行
- [ ] README の「Anthropic」表記を「Gemini」に修正
- [ ] 配布物に **第三者ライブラリのライセンス一覧** を同梱（`THIRD_PARTY_LICENSES.txt` 等）
- [ ] 子どもの音声・プロンプト入力に関する **プライバシーポリシー** を整備
- [ ] エンドユーザー（保護者）に「音声がGoogleに送信される」ことを告知
- [ ] APIキーを利用者に配布せずに済む構成（中継サーバー）への変更検討

---

## まとめ

ライセンス的に **「商用利用不可」** のものは無し。すべてMIT/BSD/Apache/SIL OFL/GPL（例外条項付き）の範囲で、商用配布可能です。

**実質的なリスクは「ライセンス」より「外部API利用規約とAPIキー管理」** の側にあります。特に:
1. README記載のAPIキー失効
2. Gemini API有料枠への切り替え
3. APIキーをエンドユーザーに渡さない構成

の3点を優先対応することをお勧めします。
