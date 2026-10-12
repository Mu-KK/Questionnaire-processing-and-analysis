# -*- coding: utf-8 -*-
"""独立逐题抽取 + 与 _已校对 对比。输出到 verify_out.txt / verify_out.json"""
import os, re, json, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import openpyxl
from docx import Document
import pdfplumber

SRC = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\2026秋问卷回收（88份）"
XLS = os.path.join(SRC, "processed", "问卷汇总结果_已校对.xlsx")
RAW = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\_audit\raw_extract.json"

QCOLS = ['1','2','3','4','5','6','7','8','9-1','9-2','9-3','9-4','10-1','10-2',
         '11','12','13','14-1','14-2','14-3','15','16-1','16-2','16-3','16-4',
         '17','18','19']

raw0 = json.load(open(RAW, encoding='utf-8'))
raw = {os.path.splitext(k)[0]: v for k, v in raw0.items()}

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


def is_letterish(s):
    t = re.sub(r'[\s,、，;；/\\\.。\-_\|]', '', s).upper()
    if not t:
        return False
    return bool(re.fullmatch(r'[A-F]{1,6}', t))


def letters(s):
    s = (s.replace('、', ',').replace('，', ',').replace('；', ',')
          .replace(';', ',').replace('/', ',').replace('\\', ','))
    return ''.join(sorted(set(re.findall(r'[A-F]', s.upper()))))


def detect(fid, q, block):
    """返回 (答案, 来源说明)"""
    lines = block['lines']
    header = lines[0]
    hb = block['header_brackets']
    # 1. 题头里 letter-ish 括号
    cand = [c for c in hb if is_letterish(c)]
    if cand:
        return cand[-1], 'header'
    # 2. 题头里非空文本括号（可能是文字答案）
    ne = [c for c in hb if c.strip()]
    # 3. 选项行里的 letter-ish 括号
    optb = []
    for ln in lines[1:]:
        for c in re.findall(r'[（(]\s*([^）)]*?)\s*[）)]', ln):
            if is_letterish(c):
                optb.append((c, ln))
    if optb:
        return optb[-1][0], 'opt-bracket:' + optb[-1][1][:40]
    # 4. 选项行末尾（或 √ 标记）
    marked = []
    for ln in lines[1:]:
        t = ln.strip()
        if not t:
            continue
        if '√' in t or '✓' in t or '✔' in t:
            m = re.match(r'^\s*([A-F])[\s\.．、]', t)
            if m:
                marked.append(m.group(1))
            else:
                m2 = re.findall(r'([A-F])', t)
                marked.extend(m2)
        # 选项行末尾单独字母  "F. 其他    C"
        m = re.match(r'^\s*([A-F])[\s\.．、].*?\s([A-F])\s*$', t)
        if m and m.group(1) != m.group(2):
            marked.append(m.group(2))
    if marked:
        return ','.join(sorted(set(marked))), 'opt-mark'
    if ne:
        return ne[-1], 'header-text'
    return '', 'none'


out = {}
print('================ 逐份逐题核对 ================')
issues = []

# 鄢萍：纯答案文档 “第 X 题：Y”
yanping = {}
yp_path = os.path.join(SRC, '鄢萍 选择题答案.docx')
if os.path.exists(yp_path):
    for p in Document(yp_path).paragraphs:
        m = re.match(r'第\s*(\d+(?:-\d+)?)\s*题\s*[：:]\s*(.+)', p.text.strip())
        if m:
            yanping[m.group(1)] = m.group(2).strip()

for fid in order:
    blocks = raw.get(fid, {})
    rec = {}
    for q in QCOLS:
        b = blocks.get(q)
        vx = excel[fid].get(q, '')
        if fid.startswith('鄢萍'):
            det, src = yanping.get(q, ''), 'yanping'
        elif b is None:
            det, src = '', 'no-block'
        else:
            det, src = detect(fid, q, b)
        rec[q] = {'excel': vx, 'detect': det, 'src': src}
        same = letters(det) == letters(vx)
        if not same:
            issues.append((fid, q, vx, det, src))
    out[fid] = rec

json.dump(out, open(r"C:\Users\LingchenKang\Desktop\ycl问卷处理\_audit\verify_out.json", 'w', encoding='utf-8'),
          ensure_ascii=False, indent=1)

print(f'\n共 {len(issues)} 处不一致：')
for fid, q, vx, det, src in issues:
    print(f'  {fid[:46]:<46} Q{q:<5} excel={vx!r:<28} detect={det!r:<20} ({src[:30]})')
