# -*- coding: utf-8 -*-
"""综合体检：基于快照，输出所有问题/可疑项。"""
import os, re, json, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import openpyxl

SNAP = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\_audit\snapshot_已校对.xlsx"
QCOLS = ['1','2','3','4','5','6','7','8','9-1','9-2','9-3','9-4','10-1','10-2',
         '11','12','13','14-1','14-2','14-3','15','16-1','16-2','16-3','16-4',
         '17','18','19']
SINGLE = ['1','2','3','4','5','6','7','8','9-1','9-2','9-3','9-4','10-1','10-2',
          '11','12','13','14-1','14-2','14-3','15','16-1','16-2','16-3']
MULTI = ['17','18','19']
VALID = {'4': set('ABCDF'), '14-3': set('ABCD'), '16-2': set('ABCD'),
         '15': set('ABCDF')}
for q in QCOLS:
    VALID.setdefault(q, set('ABCDEF'))

wb = openpyxl.load_workbook(SNAP, data_only=True)
ws = wb['Sheet1']; rows = list(ws.iter_rows(values_only=True)); hdr = list(rows[0])
data = []
for r in rows[1:]:
    if r[0] is None: continue
    data.append({k: ('' if v is None else str(v).strip()) for k, v in zip(hdr, r)})


def letters(s):
    return ''.join(sorted(set(re.findall(r'[A-F]', s.upper()))))


print(f'共 {len(data)} 行，{len(QCOLS)} 题')
print()
print('########## 1. 空单元格 ##########')
for d in data:
    for q in QCOLS:
        if d[q] == '':
            print(f'  {d["ID"]}  Q{q}')

print()
print('########## 2. 含非字母字符的值（未按选项字母填写 / 附加文字） ##########')
for d in data:
    for q in QCOLS:
        v = d[q]
        if v == '':
            continue
        if not re.fullmatch(r'[A-Fa-f,、，;；/\s\.]+', v):
            print(f'  {d["ID"][:40]:<40} Q{q:<5} = {v!r}')

print()
print('########## 3. 单选出现多字母 ##########')
for d in data:
    for q in SINGLE:
        L = letters(d[q])
        if len(L) > 1:
            print(f'  {d["ID"][:40]:<40} Q{q:<5} = {d[q]!r}')

print()
print('########## 4. 答案含该题不存在的选项字母 ##########')
for d in data:
    for q in QCOLS:
        for ch in set(re.findall(r'[A-F]', d[q].upper())):
            if ch not in VALID[q]:
                print(f'  {d["ID"][:40]:<40} Q{q:<5} = {d[q]!r}  (无 {ch} 选项)')

print()
print('########## 5. 多选题只选了 1 个选项 ##########')
for d in data:
    for q in MULTI:
        L = letters(d[q])
        if len(L) == 1:
            print(f'  {d["ID"][:40]:<40} Q{q:<5} = {d[q]!r}')

print()
print('########## 6. 多选/16-4 用 “/” 或空白表示未答 ##########')
for d in data:
    for q in QCOLS:
        v = d[q]
        if v in ('/', '-', '—', '无', '未答'):
            print(f'  {d["ID"][:40]:<40} Q{q:<5} = {v!r}')

print()
print('########## 7. 多选答案分隔符写法不一致（样本） ##########')
import collections
seps = collections.Counter()
for d in data:
    for q in MULTI + ['16-4']:
        v = d[q]
        if not v:
            continue
        if re.fullmatch(r'[A-F]{2,6}', v):
            seps['无分隔(ABCDE)'] += 1
        elif ',' in v or '，' in v or '、' in v:
            seps['逗号分隔(A,B)'] += 1
        elif ' ' in v.strip() and re.fullmatch(r'[A-F\s]+', v):
            seps['空格分隔(A B)'] += 1
        else:
            seps['其他'] += 1
for k, v in seps.items():
    print(f'  {k}: {v}')
