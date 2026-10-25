# AI お絵描き セットアップ
#  1) アプリ一式をデスクトップの「AIDrawing」フォルダに配置
#  2) デスクトップにショートカット「AIお絵描き」を作成
#  3) （任意）スタートアップに登録して、PC起動時に自動で全画面起動
#  4) （任意）スリープ・画面オフを無効化（展示用PC向け）
#
#  使い方: dist\AIDrawing\setup\install.bat をダブルクリック
#  配置先を変えたいとき: install.bat -Target "D:\AIDrawing"
param(
    [string]$Target = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'AIDrawing'),
    [switch]$Yes   # 質問せずに全部「はい」（自動起動ON・スリープ無効はOFFのまま）
)

$ErrorActionPreference = 'Stop'
$AppName = 'AIお絵描き'

function Info($m) { Write-Host "  $m" }
function Ok($m)   { Write-Host "  [OK] $m" -ForegroundColor Green }
function Warn($m) { Write-Host "  [注意] $m" -ForegroundColor Yellow }
function Ask($q, $defaultYes) {
    if ($Yes) { return $defaultYes }
    $suffix = if ($defaultYes) { '[Y/n]' } else { '[y/N]' }
    $a = Read-Host "  $q $suffix"
    if ([string]::IsNullOrWhiteSpace($a)) { return $defaultYes }
    return $a -match '^[yY]'
}

Write-Host ''
Write-Host '============================================================'
Write-Host "  $AppName セットアップ"
Write-Host '============================================================'

# --- コピー元（このスクリプトの1つ上 = AIDrawing.exe があるフォルダ） ---
$src = Split-Path -Parent $PSScriptRoot
if (-not (Test-Path (Join-Path $src 'AIDrawing.exe'))) {
    Write-Host "  [エラー] AIDrawing.exe が見つかりません: $src" -ForegroundColor Red
    Write-Host '  build.bat でビルドしてから dist\AIDrawing\setup\install.bat を実行してください。'
    exit 1
}

# --- 起動中なら止める ---
$running = Get-Process -Name 'AIDrawing' -ErrorAction SilentlyContinue
if ($running) {
    Info '起動中の AI お絵描きを終了します...'
    $running | Stop-Process -Force
    Start-Sleep -Seconds 2
}

# --- 1) 配置 ---
$srcFull = (Resolve-Path $src).Path.TrimEnd('\')
$tgtFull = [IO.Path]::GetFullPath($Target).TrimEnd('\')
if ($srcFull -ieq $tgtFull) {
    Info "すでに配置先にあります: $tgtFull"
} else {
    Info "配置先: $tgtFull"
    New-Item -ItemType Directory -Force -Path $tgtFull | Out-Null
    # .env・ログ・評価データは上書きしない（更新インストールでも設定や記録が消えないように）
    robocopy $srcFull $tgtFull /E /NFL /NDL /NJH /NJS /NP /R:2 /W:1 `
        /XF .env debug.xlsx ratings.json /XD logs | Out-Null
    if ($LASTEXITCODE -ge 8) { throw "コピーに失敗しました (robocopy $LASTEXITCODE)" }
    $srcEnv = Join-Path $srcFull '.env'
    $tgtEnv = Join-Path $tgtFull '.env'
    if ((Test-Path $srcEnv) -and -not (Test-Path $tgtEnv)) { Copy-Item $srcEnv $tgtEnv }
    Ok 'アプリを配置しました'
}

# ダウンロードした zip 由来の「ブロック」を解除
Get-ChildItem -Path $tgtFull -Recurse -File | Unblock-File -ErrorAction SilentlyContinue

# --- .env（APIキー） ---
$envPath = Join-Path $tgtFull '.env'
if (-not (Test-Path $envPath)) {
    $example = Join-Path $tgtFull '.env.example'
    if (Test-Path $example) { Copy-Item $example $envPath } else { Set-Content -Path $envPath -Value 'GEMINI_API_KEY=' -Encoding UTF8 }
    Warn '.env を作成しました。メモ帳が開くので GEMINI_API_KEY= のあとに APIキーを書いて保存してください。'
    Start-Process notepad.exe $envPath -Wait
} else {
    Ok '.env（APIキー設定）はそのまま使います'
}

# --- Chrome の確認 ---
$chromeCands = @(
    "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
    "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
)
if ($chromeCands | Where-Object { $_ -and (Test-Path $_) }) {
    Ok 'Google Chrome を確認しました'
} else {
    Warn 'Google Chrome が見つかりません。Chrome をインストールしてください（無い場合は Edge で起動します）。'
}

# --- 2) ショートカット ---
$exe = Join-Path $tgtFull 'AIDrawing.exe'
$shell = New-Object -ComObject WScript.Shell
function New-AppShortcut([string]$lnkPath, [string]$arguments) {
    $s = $shell.CreateShortcut($lnkPath)
    $s.TargetPath = $exe
    $s.Arguments = $arguments
    $s.WorkingDirectory = $tgtFull
    $s.IconLocation = "$exe,0"
    $s.Description = $AppName
    $s.Save()
}

$desktopLnk = Join-Path ([Environment]::GetFolderPath('Desktop')) "$AppName.lnk"
New-AppShortcut $desktopLnk '--browser'
Ok "デスクトップにショートカットを作成: $AppName"

# --- 3) スタートアップ（自動起動） ---
$startupLnk = Join-Path ([Environment]::GetFolderPath('Startup')) "$AppName.lnk"
if (Ask 'PC起動（サインイン）時に自動で全画面起動しますか？' $true) {
    # 起動直後はネットワークの準備に時間がかかるので 10 秒待ってから起動
    New-AppShortcut $startupLnk '--browser --delay=10'
    Ok 'スタートアップに登録しました（設定 > アプリ > スタートアップ でON/OFFできます）'
} elseif (Test-Path $startupLnk) {
    Remove-Item $startupLnk -Force
    Info 'スタートアップ登録を解除しました'
}

# --- 4) スリープ・画面オフ無効（任意） ---
if (Ask '展示用に、スリープと画面オフを無効にしますか？（電源設定を変更します）' $false) {
    powercfg /change standby-timeout-ac 0
    powercfg /change monitor-timeout-ac 0
    powercfg /change standby-timeout-dc 0
    powercfg /change monitor-timeout-dc 0
    Ok 'スリープ・画面オフを無効にしました'
}

Write-Host ''
Write-Host '============================================================'
Write-Host '  セットアップ完了'
Write-Host "  ・起動: デスクトップの「$AppName」をダブルクリック"
Write-Host '  ・管理画面: タイトルを2秒長押し → 暗証番号（初期値 0001）'
Write-Host '  ・暗証番号の変更: .env に ADMIN_PIN=1234 のように書く'
Write-Host '============================================================'

if (Ask '今すぐ起動しますか？' $false) {
    Start-Process -FilePath $exe -ArgumentList '--browser' -WorkingDirectory $tgtFull
}
