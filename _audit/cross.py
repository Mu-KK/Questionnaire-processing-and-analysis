# -*- coding: utf-8 -*-
"""把每份原始问卷的“题头括号内容”与 Excel 值对照，列出不一致处。"""
import os, re, json, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import openpyxl

SRC = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\2026秋问卷回收（88份）"
XLS = os.path.join(SRC, "processed", "问卷汇总结果_已校对.xlsx")
RAW = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\_audit\raw_extract.json"

QCOLS = ['1','2','3','4','5','6','7','8','9-1','9-2','9-3','9-4','10-1','10-2',
         '11','12','13','14-1','14-2','14-3','15','16-1','16-2','16-3','16-4',
         '17','18','19']

raw = json.load(open(RAW, encoding='utf-8'))
wb = openpyxl.load_workbook(XLS, data_only=True)
ws = wb['Sheet1']
rows = list(ws.iter_rows(values_only=True))
header = list(rows[0])
excel = {}
order = []
for r in rows[1:]:
    d = {k: ('' if v is None else str(v).strip()) for k, v in zip(header, r)}
    excel[d['ID']] = d
    order.append(d['ID'])


def clean(s):
    return re.sub(r'[\s\u200c\.。]', '', s)


def letters(s):
    s = s.replace('、', ',').replace('，', ',').replace('；', ',').replace(';', ',')
    s = s.replace('/', ',').replace('\\', ',').replace('|', ',').replace('和', ',')
    return ''.join(sorted(set(re.findall(r'[A-Za-z]', s.upper()))))


def is_letterish(s):
    t = s.strip()
    if not t:
        return False
    if not re.search(r'[A-Za-z]', t):
        return False
    # 只允许字母、空格、标点、下划线
    return bool(re.fullmatch(r'[A-Za-z\s,、，;；/\\\.。\-_()（）:：]*', t))


print('=== 题头括号与 Excel 不一致（潜在错误） ===')
count = 0
for fid in order:
    if fid not in raw:
        continue
    blocks = raw[fid]
    for q in QCOLS:
        v_excel = excel[fid][q]
        b = blocks.get(q)
        if not b:
            print(f'[缺题块] {fid[:40]} Q{q}')
            continue
        cands = [c for c in b['header_brackets'] if is_letterish(c)]
        hb_ans = cands[-1] if cands else ''
        # 忽略纯文本型答案的比较
        if not hb_ans:
            continue
        if letters(hb_ans) != letters(v_excel):
            count += 1
            print(f'  {fid[:44]:<44} Q{q:<5} excel={v_excel!r:<24} 题头括号={hb_ans!r}')
            print(f'       header: {b["header"][:150]}')
            if b['nonheader_brackets']:
                print(f'       非题头括号: {b["nonheader_brackets"]}')
print('不一致数:', count)
