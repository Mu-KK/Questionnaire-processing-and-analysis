# -*- coding: utf-8 -*-
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import openpyxl
SRC = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\2026秋问卷回收（88份）"
for xls in ["问卷汇总结果.xlsx", "问卷汇总结果_已校对.xlsx"]:
    p = os.path.join(SRC, "processed", xls)
    wb = openpyxl.load_workbook(p, data_only=True)
    ws = wb['Sheet1']
    rows = list(ws.iter_rows(values_only=True))
    hdr = list(rows[0])
    key = sys.argv[1]
    print('#####', xls)
    for r in rows[1:]:
        if r[0] and key in str(r[0]):
            for k, v in zip(hdr, r):
                print(f'   {k:<5} = {v}')
