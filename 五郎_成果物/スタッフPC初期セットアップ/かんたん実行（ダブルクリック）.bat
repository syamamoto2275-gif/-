@echo off
chcp 65001 > nul
rem ===== スタッフPC 初期セットアップ 実行用 =====
rem 同じフォルダにある .ps1 を管理者権限で起動します
setlocal
set "PS="
for %%f in ("%~dp0*.ps1") do set "PS=%%f"
if "%PS%"=="" (
  echo [エラー] 同じフォルダに .ps1 が見つかりません。
  pause
  exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process powershell -Verb RunAs -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-File','\"%PS%\"'"
exit /b 0
