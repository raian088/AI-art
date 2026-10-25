# AI お絵描き アンインストール（ショートカットと自動起動の解除）
#  アプリ本体のフォルダは消しません（ログやAPIキー設定が入っているため）。
#  不要ならフォルダごと手動で削除してください。
$ErrorActionPreference = 'Continue'
$AppName = 'AIお絵描き'

Write-Host ''
Write-Host "  $AppName の自動起動とショートカットを解除します"

Get-Process -Name 'AIDrawing' -ErrorAction SilentlyContinue | Stop-Process -Force

$targets = @(
    (Join-Path ([Environment]::GetFolderPath('Startup')) "$AppName.lnk"),
    (Join-Path ([Environment]::GetFolderPath('Desktop')) "$AppName.lnk")
)
foreach ($t in $targets) {
    if (Test-Path $t) { Remove-Item $t -Force; Write-Host "  [OK] 削除: $t" -ForegroundColor Green }
}

# キオスク用 Chrome プロファイル（普段の Chrome とは別物）
$kioskProfile = Join-Path $env:LOCALAPPDATA 'AIDrawing'
if (Test-Path $kioskProfile) {
    $a = Read-Host '  キオスク用の Chrome プロファイルも削除しますか？ [y/N]'
    if ($a -match '^[yY]') { Remove-Item $kioskProfile -Recurse -Force; Write-Host '  [OK] 削除しました' -ForegroundColor Green }
}

Write-Host ''
Write-Host "  完了。アプリ本体のフォルダ（$(Split-Path -Parent $PSScriptRoot)）は残しています。"
