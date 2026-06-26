# =====================================================================
#  スタッフPC 初期セットアップ
#  Python / Node.js / pandoc / docx を「未導入なら」自動インストール
#  ※ 何回実行しても安全（入っているものは飛ばします）
#  株式会社SEED ／ 作成:五郎(Claude Code)
# =====================================================================

$ErrorActionPreference = 'Continue'

function Refresh-Path {
  $env:Path = [Environment]::GetEnvironmentVariable('Path','Machine') + ';' +
              [Environment]::GetEnvironmentVariable('Path','User')
}

function Has-Cmd($name) {
  Refresh-Path
  $c = Get-Command $name -ErrorAction SilentlyContinue
  if (-not $c) { return $false }
  # Microsoft Store版の「空のPython」を正規導入済みと誤認しないための対策
  if ($name -eq 'python' -and $c.Source -like '*WindowsApps*') { return $false }
  return $true
}

Write-Host ''
Write-Host '==============================================' -ForegroundColor Cyan
Write-Host '   スタッフPC 初期セットアップ を開始します'      -ForegroundColor Cyan
Write-Host '==============================================' -ForegroundColor Cyan
Write-Host ''

# --- winget（Windows標準のインストーラー）確認 ---
if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
  Write-Host '[エラー] winget が見つかりません。' -ForegroundColor Red
  Write-Host '        Microsoft Store で「アプリ インストーラー」を入れてから、もう一度実行してください。' -ForegroundColor Red
  Read-Host 'Enterキーで終了します'
  exit 1
}

# --- 導入対象ツール一覧 ---
$tools = @(
  @{ Name='Python';  Id='Python.Python.3.13';    Scope='user';    Check={ (Has-Cmd 'py') -or (Has-Cmd 'python') } }
  @{ Name='Node.js'; Id='OpenJS.NodeJS.LTS';     Scope='machine'; Check={ Has-Cmd 'node' } }
  @{ Name='pandoc';  Id='JohnMacFarlane.Pandoc'; Scope='user';    Check={ Has-Cmd 'pandoc' } }
)

foreach ($t in $tools) {
  if (& $t.Check) {
    Write-Host ("[OK]   {0} は既に導入済みです。" -f $t.Name) -ForegroundColor Green
    continue
  }
  Write-Host ("[導入] {0} をインストールします..." -f $t.Name) -ForegroundColor Yellow
  $wingetArgs = @('install','--id',$t.Id,
                  '--accept-package-agreements','--accept-source-agreements',
                  '--silent','--disable-interactivity')
  if ($t.Scope -eq 'user') { $wingetArgs += @('--scope','user') }
  winget @wingetArgs | Out-Null
  Refresh-Path
  if (& $t.Check) {
    Write-Host ("[完了] {0} を導入しました。" -f $t.Name) -ForegroundColor Green
  } else {
    Write-Host ("[要確認] {0} の導入が確認できませんでした。PCを再起動して再実行してください。" -f $t.Name) -ForegroundColor Red
  }
}

# --- docx（Word自動生成ライブラリ／npmで導入） ---
Refresh-Path
if (Has-Cmd 'npm') {
  $hasDocx = npm ls -g docx 2>$null | Select-String 'docx@'
  if ($hasDocx) {
    Write-Host '[OK]   docx (Word生成ライブラリ) は既に導入済みです。' -ForegroundColor Green
  } else {
    Write-Host '[導入] docx (Word生成ライブラリ) をインストールします...' -ForegroundColor Yellow
    npm install -g docx | Out-Null
    if (npm ls -g docx 2>$null | Select-String 'docx@') {
      Write-Host '[完了] docx を導入しました。' -ForegroundColor Green
    } else {
      Write-Host '[要確認] docx の導入が確認できませんでした。' -ForegroundColor Red
    }
  }
} else {
  Write-Host '[保留] npm が見つからないため docx は後回しです。' -ForegroundColor Yellow
  Write-Host '       一度PCを再起動し、このセットアップをもう一度実行してください。' -ForegroundColor Yellow
}

Write-Host ''
Write-Host '==============================================' -ForegroundColor Cyan
Write-Host '   セットアップ完了'                              -ForegroundColor Cyan
Write-Host '==============================================' -ForegroundColor Cyan
Write-Host ''
Write-Host '※ 入れたばかりのツールは、Claude Code やターミナルを' -ForegroundColor White
Write-Host '   一度閉じて開き直すと、確実に使えるようになります。'   -ForegroundColor White
Write-Host ''
Read-Host 'Enterキーで閉じます'
