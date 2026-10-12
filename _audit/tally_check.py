# -*- coding: utf-8 -*-
"""核对 问卷选项统计结果.xlsx 是否与已校对数据一致。"""
import os, re, sys, io, collections
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import openpyxl

SNAP = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\_audit\snapshot_已校对.xlsx"
SRC = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\2026秋问卷回收（88份）\processed"
QCOLS = ['1','2','3','4','5','6','7','8','9-1','9-2','9-3','9-4','10-1','10-2',
         '11','12','13','14-1','14-2','14-3','15','16-1','16-2','16-3','16-4',
         '17','18','19']


def load(p):
    wb = openpyxl.load_workbook(p, data_only=True); ws = wb['Sheet1']
    rows = list(ws.iter_rows(values_only=True)); hdr = list(rows[0])
    return hdr, [dict(zip(hdr, r)) for r in rows[1:] if r[0] is not None]


h, data = load(SNAP)
cnt = collections.defaultdict(collections.Counter)
for d in data:
    for q in QCOLS:
        v = str(d[q] or '')
        for ch in set(re.findall(r'[A-Z]', v.upper())):
            cnt[q][ch] += 1

for fn in ['问卷选项统计结果.xlsx']:
    p = os.path.join(SRC, fn)
    if not os.path.exists(p):
        continue
    hh, rows = load(p)
    print('=====', fn, hh)
    bad = 0
    for r in rows:
        q = str(r[hh[0]]); opt = str(r[hh[1]]); n = r[hh[2]]
        mine = cnt.get(q, {}).get(opt, 0)
        if int(n or 0) != mine:
            bad += 1
            print(f'  差异: Q{q} {opt}  表内={n}  复算={mine}')
    print('   总行数:', len(rows), ' 差异数:', bad)
