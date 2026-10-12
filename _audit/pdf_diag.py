# -*- coding: utf-8 -*-
"""PDF 专项：列出每题所有 A-F 型括号候选 + 末尾单字母，供人工判读。"""
import os, re, json, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import openpyxl

SRC = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\2026秋问卷回收（88份）"
RAW = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\_audit\raw_extract.json"
QCOLS = ['1','2','3','4','5','6','7','8','9-1','9-2','9-3','9-4','10-1','10-2',
         '11','12','13','14-1','14-2','14-3','15','16-1','16-2','16-3','16-4',
         '17','18','19']
raw0 = json.load(open(RAW, encoding='utf-8'))
raw = {os.path.splitext(k)[0]: v for k, v in raw0.items()}

wb = openpyxl.load_workbook(os.path.join(SRC, 'processed', '问卷汇总结果_已校对.xlsx'), data_only=True)
ws = wb['Sheet1']
rows = list(ws.iter_rows(values_only=True))
hdr = list(rows[0])
excel = {}
for r in rows[1:]:
    d = {k: ('' if v is None else str(v).strip()) for k, v in zip(hdr, r)}
    excel[d['ID']] = d


def letterish(s):
    t = re.sub(r'[\s,、，;；/\\\.。\-_\|]', '', s).upper()
    return bool(t) and bool(re.fullmatch(r'[A-F]{1,6}', t))


for fid_ext in sorted(raw0):
    if not fid_ext.lower().endswith('.pdf'):
        continue
    fid = os.path.splitext(fid_ext)[0]
    print('=' * 80)
    print(fid)
    blocks = raw[fid]
    for q in QCOLS:
        b = blocks.get(q)
        if not b:
            print(f'  Q{q:<5} 无题块   excel={excel[fid][q]!r}')
            continue
        joined = '\n'.join(b['lines'])
        cands = []
        for m in re.finditer(r'[（(]\s*([^）)]{1,12}?)\s*[）)]', joined):
            c = m.group(1)
            if letterish(c):
                cands.append((c, joined[max(0, m.start()-25):m.end()].replace('\n', '⏎')))
        tails = []
        for ln in b['lines']:
            t = ln.strip()
            m = re.match(r'^([A-F])[\s\.．、].*?\s([A-F])\s*$', t)
            if m and m.group(1) != m.group(2):
                tails.append(t)
        print(f'  Q{q:<5} excel={excel[fid][q]!r:<14} 括号候选={[c[0] for c in cands]} 末尾字母={tails[:2]}')
        for c, ctx in cands:
            print(f'          · {c!r}  …{ctx}')
