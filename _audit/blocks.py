# -*- coding: utf-8 -*-
"""打印指定文件/题号的原始题块。用法: blocks.py 文件名关键字 题号 [题号...]"""
import os, re, json, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
RAW = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\_audit\raw_extract.json"
raw0 = json.load(open(RAW, encoding='utf-8'))
raw = {os.path.splitext(k)[0]: v for k, v in raw0.items()}
key = sys.argv[1]
qs = sys.argv[2:]
for fid in raw:
    if key in fid:
        print('=' * 70)
        print(fid)
        qlist = qs if qs else ['1','2','3','4','5','6','7','8','9-1','9-2','9-3','9-4',
                               '10-1','10-2','11','12','13','14-1','14-2','14-3','15',
                               '16-1','16-2','16-3','16-4','17','18','19']
        for q in qlist:
            b = raw[fid].get(q)
            print(f'--- Q{q} ---')
            if not b:
                print('  (无题块)')
                continue
            for ln in b['lines']:
                print('   ', ln)
