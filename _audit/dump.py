# -*- coding: utf-8 -*-
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from docx import Document
import pdfplumber

SRC = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\2026秋问卷回收（88份）"
name = sys.argv[1]
path = os.path.join(SRC, name)
if path.lower().endswith('.docx'):
    doc = Document(path)
    for i, p in enumerate(doc.paragraphs):
        t = p.text
        if t.strip():
            print(i, repr(t))
else:
    with pdfplumber.open(path) as pdf:
        for pi, page in enumerate(pdf.pages):
            print(f'----- page {pi+1} -----')
            print(page.extract_text())
