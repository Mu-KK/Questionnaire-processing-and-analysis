# -*- coding: utf-8 -*-
import os, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import openpyxl

SRC = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\2026秋问卷回收（88份）"
QCOLS = ['1','2','3','4','5','6','7','8','9-1','9-2','9-3','9-4','10-1','10-2',
         '11','12','13','14-1','14-2','14-3','15','16-1','16-2','16-3','16-4',
         '17','18','19']


def load(p):
    wb = openpyxl.load_workbook(p, data_only=True)
    ws = wb['Sheet1']
    rows = list(ws.iter_rows(values_only=True))
    hdr = list(rows[0])
    d = {}
    for r in rows[1:]:
        if r[0] is None:
            continue
        d[str(r[0])] = {k: ('' if v is None else str(v).strip()) for k, v in zip(hdr, r)}
    return d


a = load(os.path.join(SRC, 'processed', '问卷汇总结果.xlsx'))
b = load(r"C:\Users\LingchenKang\Desktop\ycl问卷处理\_audit\snapshot_已校对.xlsx")

print('问卷汇总结果.xlsx 行数:', len(a))
print('已校对.xlsx 行数      :', len(b))
print()
print('=== 只在已校对中出现的行（新增，预计是4份PDF） ===')
for k in b:
    if k not in a:
        print('  +', k)
print()
print('=== 相同 ID 下的单元格差异（手工改动） ===')
n = 0
for k in a:
    if k not in b:
        print('  - 已校对缺失:', k)
        continue
    for q in QCOLS:
        if a[k][q] != b[k][q]:
            n += 1
            print(f'  {k[:44]:<44} Q{q:<5} 汇总={a[k][q]!r:<20} 已校对={b[k][q]!r}')
print('差异数:', n)
