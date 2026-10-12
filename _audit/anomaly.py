# -*- coding: utf-8 -*-
"""列出所有“答案位置异常”的题块：题头无括号/空括号/非字母答案/选项行有括号 等。"""
import os, re, json, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

RAW = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\_audit\raw_extract.json"
QCOLS = ['1','2','3','4','5','6','7','8','9-1','9-2','9-3','9-4','10-1','10-2',
         '11','12','13','14-1','14-2','14-3','15','16-1','16-2','16-3','16-4',
         '17','18','19']
raw = json.load(open(RAW, encoding='utf-8'))


def is_letterish(s):
    t = s.strip()
    if not t or not re.search(r'[A-Za-z]', t):
        return False
    return bool(re.fullmatch(r'[A-Za-z\s,、，;；/\\\.。\-_()（）:：]*', t))


print('############ A. 题块里“非题头括号”包含字母（脚本可能取错） ############')
for fid, blocks in raw.items():
    for q in QCOLS:
        b = blocks.get(q)
        if not b:
            continue
        nh = [c for c in b['nonheader_brackets'] if is_letterish(c)]
        if nh:
            print(f'{fid[:44]:<44} Q{q:<5} 非题头括号字母={nh}')

print()
print('############ B. 题头没有任何字母型括号 ############')
for fid, blocks in raw.items():
    miss = []
    for q in QCOLS:
        b = blocks.get(q)
        if not b:
            miss.append(q + '(无题块)')
            continue
        cands = [c for c in b['header_brackets'] if is_letterish(c)]
        if not cands:
            miss.append(f'{q}{b["header_brackets"]}')
    if miss:
        print(f'--- {fid}')
        for m in miss:
            print('     ', m)

print()
print('############ C. 缺失的题块（应有28题） ############')
for fid, blocks in raw.items():
    have = [q for q in QCOLS if q in blocks]
    if len(have) != 28:
        missing = [q for q in QCOLS if q not in blocks]
        extra = [k for k in blocks if k not in QCOLS]
        print(f'{fid[:50]:<50} 有{len(have)}题 缺={missing} 额外={extra}')
