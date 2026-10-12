# -*- coding: utf-8 -*-
"""找出所有“填空”痕迹：选项行里的下划线中/其他处填了内容，并对照 Excel 值。"""
import os, re, json, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import openpyxl

SRC = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\2026秋问卷回收（88份）"
RAW = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\_audit\raw_extract.json"
SNAP = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\_audit\snapshot_已校对.xlsx"
QCOLS = ['1','2','3','4','5','6','7','8','9-1','9-2','9-3','9-4','10-1','10-2',
         '11','12','13','14-1','14-2','14-3','15','16-1','16-2','16-3','16-4',
         '17','18','19']

raw0 = json.load(open(RAW, encoding='utf-8'))
raw = {os.path.splitext(k)[0]: v for k, v in raw0.items()}

wb = openpyxl.load_workbook(SNAP, data_only=True)
ws = wb['Sheet1']
rows = list(ws.iter_rows(values_only=True))
hdr = list(rows[0])
excel = {}
for r in rows[1:]:
    if r[0] is None:
        continue
    excel[str(r[0])] = {k: ('' if v is None else str(v).strip()) for k, v in zip(hdr, r)}


def user_filled(line):
    """返回该选项行被用户填入的额外内容（模板下划线之外的文字），若无则 None。"""
    if '_' not in line:
        return None
    # 去掉常规选项前缀
    seg = re.sub(r'^\s*[A-F][\.．、]?', '', line)
    # 找出下划线区间内夹着的非下划线内容
    found = []
    for m in re.finditer(r'_{1,}([^_\s][^_]*?[^_\s])?_{1,}', seg):
        inside = m.group(1)
        if inside and re.search(r'[\u4e00-\u9fff0-9%.]', inside):
            found.append(inside.strip())
    if found:
        return ' | '.join(found)
    return None


print('=== 选项“填空”被填写，且该题 Excel 值不是纯字母（可能信息丢失/口径不一致） ===')
for fid in excel:
    if fid.startswith('鄢萍'):
        continue
    for q in QCOLS:
        b = raw.get(fid, {}).get(q)
        if not b:
            continue
        fills = []
        for ln in b['lines'][1:]:
            c = user_filled(ln)
            if c:
                fills.append((ln.strip()[:60], c))
        if not fills:
            continue
        vx = excel[fid].get(q, '')
        pure = bool(re.fullmatch(r'[A-F,、，\s]+', vx))
        mark = '' if pure else '   <<< 非纯字母'
        print(f'  {fid[:40]:<40} Q{q:<5} excel={vx!r:<18}{mark}')
        for ln, c in fills:
            print(f'        - {ln}   ⇒ 填入: {c!r}')
