"""
.xlsm からVBAモジュールを .bas として書き出す。
使い方: python extract_vba.py <対象.xlsm> <出力フォルダ> <ラベル>
依存: pip install oletools （本物のPythonで）
"""
import sys, warnings, os
warnings.simplefilter("ignore")
from oletools.olevba import VBA_Parser

if len(sys.argv) < 4:
    print("使い方: python extract_vba.py <対象.xlsm> <出力フォルダ> <ラベル>")
    sys.exit(1)

path, outdir, label = sys.argv[1], sys.argv[2], sys.argv[3]
os.makedirs(outdir, exist_ok=True)
vp = VBA_Parser(path)
count = 0
for (fn, stream, vba_name, vba_code) in vp.extract_macros():
    if not vba_code or not vba_code.strip():
        continue
    count += 1
    safe = vba_name.replace("/", "_")
    out = os.path.join(outdir, f"{label}_{safe}.bas")
    with open(out, "w", encoding="utf-8") as f:
        f.write(vba_code)
    print(f"  {vba_name}: {vba_code.count(chr(10))+1}行 -> {os.path.basename(out)}")
print(f"合計モジュール数: {count}")
vp.close()
