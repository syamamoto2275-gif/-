# .xlsm からVBAを取り出す手順（救出）

Excelの `.xlsm` は中身がZIP。VBAは `xl/vbaProject.bin`（OLE形式・圧縮）に入っているため、
そのままテキストでは読めない。`oletools` の `olevba` を使って取り出す。

## 前提ツール

- Python（本物のインストール版。Microsoft Store版のスタブ `WindowsApps\python.exe` ではなく、
  例: `C:\Users\User\AppData\Local\Programs\Python\Python313\python.exe`）
- 初回のみ: `python -m pip install oletools`

## 取り出しスクリプト

各モジュールを `.bas` として書き出す最小スクリプト（文字化け対策で警告は抑制）:

```python
import sys, warnings, os
warnings.simplefilter("ignore")
from oletools.olevba import VBA_Parser

path, outdir, label = sys.argv[1], sys.argv[2], sys.argv[3]
vp = VBA_Parser(path)
for (fn, stream, vba_name, vba_code) in vp.extract_macros():
    if not vba_code or not vba_code.strip():
        continue
    safe = vba_name.replace("/", "_")
    with open(os.path.join(outdir, f"{label}_{safe}.bas"), "w", encoding="utf-8") as f:
        f.write(vba_code)
    print(f"  {vba_name}: {vba_code.count(chr(10))+1}行")
vp.close()
```

実行（PowerShell例）:

```powershell
$py = "C:\Users\User\AppData\Local\Programs\Python\Python313\python.exe"
& $py extract_vba.py "対象の.xlsm" "出力フォルダ" "ラベル" 2>$null
```

## 取り出した後の整理

- **標準モジュール（Module1, Module2, …）** に実処理が入っている。これを型に沿って整理する。
- **クラスモジュール（Sheet1.cls, ThisWorkbook.cls 等）** は中身が空（9行程度のヘッダのみ）なら無視。
- olevbaの出力は先頭に `Attribute VB_Name = "..."` が二重に付くことがある → 1つに直す。
- コンソール表示が文字化けしても、**書き出した.basファイル自体はUTF-8で正しい**（表示だけの問題）。

## 最新版の見分け方（重要）

同名で複数版があることが多い。**ファイル名の「最新」「完成版」を鵜呑みにしない**。
- 更新日時が新しいものを候補にする。
- 取り出して**行数・モジュール数を比較**する。手順モジュールが増えている方が本当の最新。
  （例：楽天では新版にだけ Module9＝実行ボタン構成と手順2b自動振り分けが存在した）

## Excelへ戻すとき

整理した `.bas` をExcelに取り込むには、VBE（Alt+F11）→ ファイル → ファイルのインポート、
または対象モジュールを削除してからインポートし直す。ボタンには対応する `実行1/2/3` を割り当てる。
取り込み後は必ず実データ1件で実行1→2→3を通して動作確認する。
