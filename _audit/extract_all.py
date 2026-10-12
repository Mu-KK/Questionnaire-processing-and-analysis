# -*- coding: utf-8 -*-
"""独立解析 88 份问卷（docx/pdf），输出 JSON 以便与 _已校对.xlsx 对比。"""
import os, re, json, sys, io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from docx import Document
import pdfplumber

SRC = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\2026秋问卷回收（88份）"
OUT = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\_audit\raw_extract.json"

SPLIT_KEYWORDS = ["（单选或填空）", "（多选或填空）",
                  "9-1", "9-2", "9-3", "9-4",
                  "10-1", "10-2",
                  "14-1", "14-2", "14-3",
                  "16-1", "16-2", "16-3", "16-4"]
MAIN_TO_FILTER = ['9', '10', '14', '16']
LOOSE_KW = re.compile(r'[（(][^）)]*?(单选|多选)或填空[^）)]*?[）)]')
BRACKET = re.compile(r'[（(]\s*([^）)]*?)\s*[）)]')


def get_lines(path):
    if path.lower().endswith('.docx'):
        doc = Document(path)
        return [p.text for p in doc.paragraphs]
    # pdf
    lines = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ''
            lines.extend(t.split('\n'))
    return lines


def is_noise(line):
    if not line.strip():
        return True
    if re.match(r'^\d+\s*/\s*\d+$', line.strip()):
        return True
    if 'EID·CMR' in line:
        return True
    if line.strip() in ('2026年中国宏观经济形势和政策问卷调查',
                        '2026年秋季中国宏观经济形势与政策调查问卷(总第27次）'):
        return True
    return False


def split_blocks(lines):
    blocks = []
    cur = []
    cur_num = None
    skip = False
    for raw in lines:
        line = raw.strip()
        if is_noise(line):
            continue
        has_kw = any(k in line for k in SPLIT_KEYWORDS) or bool(LOOSE_KW.search(line))
        if has_kw:
            if cur and not skip:
                blocks.append({'number': cur_num, 'lines': cur})
            cur = [line]
            skip = False
            m = re.match(r'^(\d+(?:-\d+)?)\s*[\.\．\（\s]', line)
            if m:
                cur_num = m.group(1)
                if cur_num in MAIN_TO_FILTER:
                    skip = True
                    cur = []
            else:
                for k in SPLIT_KEYWORDS[2:]:
                    if k in line:
                        cur_num = k
                        break
        else:
            if cur and not skip:
                cur.append(line)
    if cur and not skip:
        blocks.append({'number': cur_num, 'lines': cur})
    return blocks


def main():
    files = sorted(os.listdir(SRC))
    files = [f for f in files if f.lower().endswith(('.docx', '.pdf')) and not f.startswith('~$')]
    out = {}
    for f in files:
        path = os.path.join(SRC, f)
        try:
            lines = get_lines(path)
            blocks = split_blocks(lines)
        except Exception as e:
            out[f] = {'error': str(e)}
            continue
        rec = {}
        for b in blocks:
            if not b['number']:
                continue
            header = b['lines'][0] if b['lines'] else ''
            hb = BRACKET.findall(header)
            ab = []
            for ln in b['lines']:
                ab.extend(BRACKET.findall(ln))
            nonheader_brackets = []
            for ln in b['lines'][1:]:
                nonheader_brackets.extend(BRACKET.findall(ln))
            rec[b['number']] = {
                'lines': b['lines'],
                'header': header,
                'header_brackets': hb,
                'block_brackets': ab,
                'nonheader_brackets': nonheader_brackets,
                'n_lines': len(b['lines']),
            }
        out[f] = rec
    with open(OUT, 'w', encoding='utf-8') as fh:
        json.dump(out, fh, ensure_ascii=False, indent=1)
    print('files:', len(files))
    print('written:', OUT)


if __name__ == '__main__':
    main()
