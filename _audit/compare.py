# -*- coding: utf-8 -*-
import os, re, json, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import openpyxl

SRC = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\2026秋问卷回收（88份）"
XLS = os.path.join(SRC, "processed", "问卷汇总结果_已校对.xlsx")
RAW = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\_audit\raw_extract.json"

QCOLS = ['1','2','3','4','5','6','7','8','9-1','9-2','9-3','9-4','10-1','10-2',
         '11','12','13','14-1','14-2','14-3','15','16-1','16-2','16-3','16-4',
         '17','18','19']
SINGLE = ['1','2','3','4','5','6','7','8','9-1','9-2','9-3','9-4','10-1','10-2',
          '11','12','13','14-1','14-2','14-3','15','16-1','16-2','16-3','16-4']
MULTI = ['17','18','19']

raw = json.load(open(RAW, encoding='utf-8'))

wb = openpyxl.load_workbook(XLS, data_only=True)
ws = wb['Sheet1']
rows = list(ws.iter_rows(values_only=True))
header = list(rows[0])
assert header[0] == 'ID'
excel = {}
order = []
for r in rows[1:]:
    if r[0] is None:
        continue
    d = {k: ('' if v is None else str(v).strip()) for k, v in zip(header, r)}
    excel[d['ID']] = d
    order.append(d['ID'])

print('Excel rows:', len(order))
print('raw files  :', len(raw))

ids_excel = set(excel)
ids_files = set(os.path.splitext(f)[0] for f in raw)
print('\n== 只在文件里、不在Excel ==')
for x in sorted(ids_files - ids_excel):
    print('  ', x)
print('== 只在Excel里、没有文件 ==')
for x in sorted(ids_excel - ids_files):
    print('  ', x)


def norm(s):
    s = s.replace('、', ',').replace('，', ',').replace('；', ',').replace(';', ',')
    s = re.sub(r'\s+', '', s)
    return s.upper()


def letters_only(s):
    s = norm(s.replace('/', '').replace('\\', ''))
    s = re.sub(r'[^A-Z]', '', s)
    return s


def cmp_letters(a, b):
    la = sorted(set(letters_only(a)))
    lb = sorted(set(letters_only(b)))
    return la, lb


problems = []
suspicious = []

# ---- 单元格层面检查 ----
print('\n\n================ 单元格层面可疑检查 ================')
for fid in order:
    row = excel[fid]
    for q in QCOLS:
        v = row.get(q, '')
        if v == '' or v is None:
            problems.append((fid, q, 'EMPTY', '空单元格', ''))
            continue
        L = letters_only(v)
        if q in SINGLE:
            # 单选：正常应为 1 个字母
            if len(set(letters_only(v))) > 1:
                suspicious.append((fid, q, 'MULTI-IN-SINGLE', v, ''))
            elif len(L) == 0:
                suspicious.append((fid, q, 'NO-LETTER', v, ''))
        if q in MULTI:
            if len(set(letters_only(v))) == 0:
                suspicious.append((fid, q, 'NO-LETTER', v, ''))
            elif len(set(letters_only(v))) == 1:
                suspicious.append((fid, q, 'ONE-IN-MULTI', v, ''))

print('\n--- 空单元格（问题） ---')
for p in problems:
    print('  ', p[0][:46], '| Q', p[1], '|', p[3])

print('\n--- 单选出现多字母 / 多选只有单字母 / 无字母（可疑） ---')
for p in suspicious:
    print('  ', p[0][:46], '| Q', p[1], '|', p[2], '|', repr(p[3]))

print()
print('空单元格数:', len(problems))
print('可疑单元格数:', len(suspicious))
