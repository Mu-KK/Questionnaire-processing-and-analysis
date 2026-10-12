# -*- coding: utf-8 -*-
"""扫描所有 docx：勾选符号 / 加粗 / 高亮 / 下划线 标记答案的情况。"""
import os, re, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from docx import Document

SRC = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\2026秋问卷回收（88份）"
MARK = ['√', '✓', '✔', '☑', '■', '●', '◆', '×', '✗']

hits = {}
bold_hits = {}
for f in sorted(os.listdir(SRC)):
    if not f.lower().endswith('.docx'):
        continue
    doc = Document(os.path.join(SRC, f))
    m_lines = []
    b_lines = []
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        if any(k in t for k in MARK):
            m_lines.append(t[:70])
        # 加粗/高亮/下划线
        if re.match(r'^\s*[A-F][\.．、]', t):
            for r in p.runs:
                if not r.text.strip():
                    continue
                bold = r.bold
                hl = None
                try:
                    hl = r.font.highlight_color
                except Exception:
                    hl = None
                ul = r.font.underline
                if bold or hl or ul:
                    b_lines.append(f'{t[:50]} || run={r.text.strip()[:20]!r} bold={bold} hl={hl} ul={ul}')
                    break
    if m_lines:
        hits[f] = m_lines
    if b_lines:
        bold_hits[f] = b_lines

print('############ 含勾选/记号 的文件 ############')
for f, ls in hits.items():
    print('---', f)
    for l in ls[:40]:
        print('   ', l)

print()
print('############ 选项行有加粗/高亮/下划线 的文件 ############')
for f, ls in bold_hits.items():
    print('---', f, f'({len(ls)} 行)')
    for l in ls[:30]:
        print('   ', l)
