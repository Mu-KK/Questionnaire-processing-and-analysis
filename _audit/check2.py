# -*- coding: utf-8 -*-
"""① 题头出现≥2个字母型括号（可能取错） ② 答案在非题头位置 的文件"""
import os, re, json, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import openpyxl
RAW = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\_audit\raw_extract.json"
SNAP = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\_audit\snapshot_已校对.xlsx"
QCOLS = ['1','2','3','4','5','6','7','8','9-1','9-2','9-3','9-4','10-1','10-2',
         '11','12','13','14-1','14-2','14-3','15','16-1','16-2','16-3','16-4',
         '17','18','19']
raw0 = json.load(open(RAW, encoding='utf-8'))
raw = {os.path.splitext(k)[0]: v for k, v in raw0.items()}
wb = openpyxl.load_workbook(SNAP, data_only=True)
ws = wb['Sheet1']; rows = list(ws.iter_rows(values_only=True)); hdr = list(rows[0])
excel = {}
for r in rows[1:]:
    if r[0] is None: continue
    excel[str(r[0])] = {k: ('' if v is None else str(v).strip()) for k, v in zip(hdr, r)}


def let(s):
    t = re.sub(r'[\s,、，;；/\\\.。\-_\|]', '', s).upper()
    return bool(t) and bool(re.fullmatch(r'[A-F]{1,6}', t))


print('=== ① 题头有 >=2 个字母型括号 ===')
for fid, blocks in raw.items():
    if fid.endswith('.pdf'): continue
    for q in QCOLS:
        b = blocks.get(q)
        if not b: continue
        cands = [c for c in b['header_brackets'] if let(c)]
        if len(cands) >= 2:
            print(f'  {fid[:40]:<40} Q{q:<5} excel={excel.get(fid,{}).get(q)!r} 候选={cands}')
            print(f'        header={b["header"][:150]}')

print()
print('=== ② 答案出现在“非题头”位置（题头无字母括号，但选项行里有） ===')
for fid, blocks in raw.items():
    if fid.endswith('.pdf'): continue
    for q in QCOLS:
        b = blocks.get(q)
        if not b: continue
        hc = [c for c in b['header_brackets'] if let(c)]
        nc = [c for c in b['nonheader_brackets'] if let(c)]
        if not hc and nc:
            print(f'  {fid[:40]:<40} Q{q:<5} excel={excel.get(fid,{}).get(q)!r} 非题头候选={nc}')
            for ln in b['lines']:
                if any(c in ln for c in nc):
                    print(f'        > {ln[:120]}')
