---
name: reference-this-pc
description: 五郎は2台（会社PC/自宅PC）。パスで自己判別する。会話ログにどちらか明記
metadata: 
  node_type: memory
  type: reference
  originSessionId: 61100090-3a5f-457e-b24c-2fa6569ba368
---

**このファイルは共有リポジトリにあり、会社PC・自宅PC両方の五郎が読む。だから「このPCは○○」と決め打ちしない。必ず自分の作業パスで自己判別すること**（2026-06-30 社長の訂正を反映）。

## PCの見分け方（パスで判別）
- **会社PC**：`C:\Users\User\Desktop\新しいフォルダー`（ユーザー名 `User`、OneDriveなし）
- **自宅PC**：`C:\Users\山本　史朗\OneDrive\Desktop\新しいフォルダー`（ユーザー名 `山本　史朗`、OneDrive配下）

自分の作業パスを見て、どちらの五郎かを判断する。会話ログ（`五郎_会話ログ_日付.md`）には「会社PC」「自宅PC」を明記する。2台は GitHubリポジトリ `syamamoto2275-gif/-` 経由で同期する。**自分が編集・作成したものをpushして共有する相手は、もう片方のPC**（会社PCならpush先は自宅PC、自宅PCならpush先は会社PC）。`.claude/memory` はPC間で同期されないため、PC共有が必要な事項はリポジトリ内ファイルに書いてpushする。

共通ルール詳細は [[reference_github_repo]] のリポジトリ内 `五郎_共通ルール.md` を参照。

## 自宅PCのツール環境（2026-06-28 整備完了）
Word/Excel/PDF/同期すべて作業可能。
- git 2.54 ／ Node v24.18＋npm 11.16＋docx 9.7（Word生成）／ pandoc 3.10 ／ Python 3.13.14
- Python本体は `C:\Users\山本　史朗\AppData\Local\Programs\Python\Python313\python.exe`。当初はWindowsストアのダミーがPATHを奪っていたため、本物のPython313とScriptsをユーザーPATH先頭に追加して解消済み。
- Python部品導入済み：openpyxl・pandas・python-docx・pypdf・pdfplumber・reportlab・Pillow。
- 未導入：gh（GitHub CLI。git本体で代替できるため任意）。

## 会社PCのツール環境
- Python本体は `C:\Users\User\AppData\Local\Programs\Python\Python313\python.exe`（Windowsストア版ダミーとは別物）。openpyxl・oletools導入済み。
