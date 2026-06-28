---
name: reference-this-pc
description: このPCは「自宅PC」。会話ログ記録時はPCを書き分ける
metadata: 
  node_type: memory
  type: reference
  originSessionId: 61100090-3a5f-457e-b24c-2fa6569ba368
---

このPC（C:\Users\山本　史朗\OneDrive\Desktop\新しいフォルダー）の五郎は**自宅PCの五郎**。

会話ログ（`五郎_会話ログ_日付.md`）を書くときは「自宅PC」と明記する。デスクトップPC＝PC1（会社/別拠点）の五郎とはGitHubリポジトリ `syamamoto2275-gif/-` 経由で同期する。`.claude/memory` はPC間で同期されないため、PC共有が必要な事項はリポジトリ内ファイルに書いてpushする。

共通ルール詳細は [[reference_github_repo]] のリポジトリ内 `五郎_共通ルール.md` を参照。

**自宅PCのツール環境（2026-06-28 整備完了）：** Word/Excel/PDF/同期すべて作業可能。
- git 2.54 ／ Node v24.18＋npm 11.16＋docx 9.7（Word生成）／ pandoc 3.10 ／ Python 3.13.14
- Python本体は `C:\Users\山本　史朗\AppData\Local\Programs\Python\Python313\python.exe`。当初はWindowsストアのダミーがPATHを奪っていたため、本物のPython313とScriptsをユーザーPATH先頭に追加して解消済み。
- Python部品導入済み：openpyxl・pandas・python-docx・pypdf・pdfplumber・reportlab・Pillow。
- 未導入：gh（GitHub CLI。git本体で代替できるため任意）。
