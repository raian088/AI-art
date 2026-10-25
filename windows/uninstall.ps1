# AI お絵描き — Windows PC の設定を元に戻す（ショートカット・自動起動・専用プロファイル）
$ErrorActionPreference = 'Continue'
$AppName = 'AIお絵描き'
$Home2   = Join-Path $env:LOCALAPPDATA 'AIDrawingKiosk'

Write-Host ''
Write-Host "  $AppName の設定を解除します"
foreach ($t in @(
    (Join-Path ([Environment]::GetFolderPath('Startup')) "$AppName.lnk"),
    (Join-Path ([Environment]::GetFolderPath('Desktop')) "$AppName.lnk"))) {
    if (Test-Path $t) { Remove-Item $t -Force; Write-Host "  [OK] 削除: $t" -ForegroundColor Green }
}
if (Test-Path $Home2) {
    $a = Read-Host '  起動スクリプトと専用 Chrome プロファイル（ログイン状態）も削除しますか？ [y/N]'
    if ($a -match '^[yY]') {
        Remove-Item $Home2 -Recurse -Force
        Write-Host '  [OK] 削除しました（Chrome が開いていると消せないので、閉じてから実行してください）' -ForegroundColor Green
    }
}
Write-Host '  完了'
