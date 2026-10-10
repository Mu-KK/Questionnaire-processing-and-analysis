# ==============================
# generate_report.py
# 1. 标准化数据：大写、合并多选、处理Q16-4特殊值
# 2. 生成问卷选项统计结果Excel
# 3. 按参考格式生成Word文档
# ==============================

import pandas as pd
import re, os, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# ===============================
# 路径配置
# ===============================
INPUT_EXCEL = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\2026秋问卷回收（88份）\processed\问卷汇总结果_已校对.xlsx"
OUTPUT_STATS = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\2026秋问卷回收（88份）\processed\问卷选项统计结果_标准化.xlsx"
OUTPUT_WORD = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\2026秋问卷回收（88份）\processed\问卷调查统计_88份.docx"
OUTPUT_STATS_RESULT = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\结果\2026秋_问卷选项统计结果_标准化.xlsx"
OUTPUT_WORD_RESULT = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\结果\2026秋_问卷调查统计_88份.docx"

# 题目定义（顺序及类型）
# 类型: 'single' = 单选或填空, 'multi' = 多选或填空
QUESTION_META = [
    ('1', 'single'), ('2', 'single'), ('3', 'single'), ('4', 'single'),
    ('5', 'single'), ('6', 'single'), ('7', 'single'), ('8', 'single'),
    ('9-1', 'single'), ('9-2', 'single'), ('9-3', 'single'), ('9-4', 'single'),
    ('10-1', 'single'), ('10-2', 'single'),
    ('11', 'single'), ('12', 'single'), ('13', 'single'),
    ('14-1', 'single'), ('14-2', 'single'), ('14-3', 'single'),
    ('15', 'single'),
    ('16-1', 'single'), ('16-2', 'single'), ('16-3', 'single'), ('16-4', 'single'),
    ('17', 'multi'), ('18', 'multi'), ('19', 'multi'),
]


# ===============================
# 标准化函数
# ===============================
def normalize_answer(value, qnum='', question_type='single'):
    """标准化单个答案值"""
    if pd.isna(value) or str(value).strip() == '':
        return [], [], ''  # (options_list, fill_text_list, raw)

    raw = str(value).strip()

    # 统一转大写
    normalized = raw.upper()

    # D/E → DE
    normalized = normalized.replace('/', ',')

    # 中文逗号、顿号 → 英文逗号
    normalized = normalized.replace('，', ',').replace('、', ',')
    # 多个空格 → 单个空格
    normalized = re.sub(r'\s+', ' ', normalized)

    # 移除空格（仅对纯选项字母）
    # 但保留填空文字中的空格

    # 把逗号和空格统一为逗号
    # 先处理连续字母（如 "AC" → "A,C"）用于多选和Q16-4
    if question_type == 'multi' or qnum in ('16-4',):
        # 插入逗号在连续大写字母之间
        normalized = re.sub(r'([A-F])([A-F])', r'\1,\2', normalized)
    parts = re.split(r'[\s,]+', normalized)
    parts = [p.strip() for p in parts if p.strip()]

    # 分离纯选项和填空文字
    options = []
    fill_texts = []

    for p in parts:
        # 如果是纯 A-F 单个字母
        if re.match(r'^[A-F]$', p):
            options.append(p)
        elif re.match(r'^[A-F].+', p):
            # 以A-F开头但后面有内容（如 "A25个基点"）
            opt_letter = p[0]
            fill = p[1:].strip()
            options.append(opt_letter)
            if fill:
                fill_texts.append(fill)
        else:
            # 纯填空文字
            fill_texts.append(p)

    # 去重
    options = list(dict.fromkeys(options))
    # 过滤 fills 中纯选项字母
    fill_texts = [f for f in fill_texts if not re.match(r'^[A-F]$', f.strip())]

    return options, fill_texts, raw


# ===============================
# 读取并标准化
# ===============================
df = pd.read_excel(INPUT_EXCEL)

print(f"读取 {len(df)} 份问卷")

# 统计结构: {题号: {选项: (count, [fill_texts])}}
stats = {}

# 逐行处理
for idx, row in df.iterrows():
    rid = row['ID']
    for qnum, qtype in QUESTION_META:
        v = row.get(qnum, '')
        opts, fills, raw = normalize_answer(v, qnum=qnum, question_type=qtype)

        if qnum not in stats:
            stats[qnum] = {}

        if qtype == 'single':
            # 单选题：一般只有一个选项
            if opts:
                opt = opts[0]  # 主选项
                # 其他选项也计入（如 Q16-4 有人选多个）
                for o in opts:
                    if o not in stats[qnum]:
                        stats[qnum][o] = {'count': 0, 'fills': []}
                    stats[qnum][o]['count'] += 1
                    stats[qnum][o]['fills'].extend(fills)
            elif fills:
                # 纯填空（如F的其他）
                if 'F' not in stats[qnum]:
                    stats[qnum]['F'] = {'count': 0, 'fills': []}
                stats[qnum]['F']['count'] += 1
                stats[qnum]['F']['fills'].extend(fills)
            else:
                # 空值
                if '_MISSING' not in stats[qnum]:
                    stats[qnum]['_MISSING'] = {'count': 0, 'fills': []}
                stats[qnum]['_MISSING']['count'] += 1

        elif qtype == 'multi':
            # 多选题：每个选项独立计数
            if opts:
                for o in opts:
                    if o not in stats[qnum]:
                        stats[qnum][o] = {'count': 0, 'fills': []}
                    stats[qnum][o]['count'] += 1
            if fills:
                if 'F' not in stats[qnum]:
                    stats[qnum]['F'] = {'count': 0, 'fills': []}
                stats[qnum]['F']['fills'].extend(fills)
            if not opts and not fills:
                if '_MISSING' not in stats[qnum]:
                    stats[qnum]['_MISSING'] = {'count': 0, 'fills': []}
                stats[qnum]['_MISSING']['count'] += 1


# ===============================
# 生成统计Excel
# ===============================
print("\n生成统计Excel...")
all_rows = []
for qnum, qtype in QUESTION_META:
    st = stats.get(qnum, {})
    total_respondents = len(df) - st.get('_MISSING', {}).get('count', 0)

    for opt in ['A', 'B', 'C', 'D', 'E', 'F']:
        if opt in st:
            cnt = st[opt]['count']
            fill_str = '; '.join(st[opt]['fills']) if st[opt]['fills'] else ''
            all_rows.append({
                '题号': qnum,
                '选项': opt,
                '人数': cnt,
                '填空': fill_str
            })
        else:
            all_rows.append({
                '题号': qnum,
                '选项': opt,
                '人数': 0,
                '填空': ''
            })

    # 总有效人次
    all_rows.append({
        '题号': qnum,
        '选项': '本题有效填写人次',
        '人数': total_respondents,
        '填空': ''
    })

stats_df = pd.DataFrame(all_rows)
stats_df.to_excel(OUTPUT_STATS, index=False)
stats_df.to_excel(OUTPUT_STATS_RESULT, index=False)
print(f"统计Excel已保存")


# ===============================
# 生成Word文档
# ===============================
print("\n生成Word文档...")

doc = Document()

# 设置默认字体
style = doc.styles['Normal']
font = style.font
font.name = '宋体'
font.size = Pt(10.5)
style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

for qnum, qtype in QUESTION_META:
    # --- 标题段落 ---
    st = stats.get(qnum, {})
    total = len(df) - st.get('_MISSING', {}).get('count', 0)

    # 标题格式
    title_text = f"{qnum}   "
    if qtype == 'single':
        title_text += "[单选或填空]"
    else:
        title_text += "[多选或填空]"

    p_title = doc.add_paragraph()
    run = p_title.add_run(title_text)
    run.font.size = Pt(10.5)
    run.font.name = '宋体'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    # 如果有漏题备注
    missing_count = st.get('_MISSING', {}).get('count', 0)
    if missing_count > 0:
        p_note = doc.add_paragraph()
        run_note = p_note.add_run(f"备注：{missing_count}人漏题")
        run_note.font.size = Pt(9)
        run_note.font.name = '宋体'
        run_note._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    # --- 表格 ---
    if qtype == 'single':
        # 3列: 选项 | 小计 | 填空
        table = doc.add_table(rows=8, cols=3, style='Table Grid')
        table.autofit = True

        headers = ['选项', '小计', '填空']
        for ci, h in enumerate(headers):
            cell = table.cell(0, ci)
            cell.text = h
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(10.5)
                    r.font.name = '宋体'
                    r._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
                    r.font.bold = True

        for ri, opt in enumerate(['A', 'B', 'C', 'D', 'E', 'F']):
            row = table.rows[ri + 1]
            # 选项
            row.cells[0].text = opt

            # 小计
            cnt = st.get(opt, {}).get('count', 0)
            row.cells[1].text = str(cnt)

            # 填空
            fills = st.get(opt, {}).get('fills', [])
            fill_text = '; '.join(fills) if fills else ''
            row.cells[2].text = fill_text

            # 设置字体
            for ci in range(3):
                for p in row.cells[ci].paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(10.5)
                        r.font.name = '宋体'
                        r._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

        # 最后一行：本题有效填写人次
        last_row = table.rows[7]
        last_row.cells[0].text = '本题有效填写人次'
        last_row.cells[1].text = str(total)
        last_row.cells[2].text = ''
        for ci in range(3):
            for p in last_row.cells[ci].paragraphs:
                for r in p.runs:
                    r.font.size = Pt(10.5)
                    r.font.name = '宋体'
                    r._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    else:
        # 多选: 2列: 选项 | 小计
        num_rows = 7  # A-F + header
        # 检查F是否有填空
        f_fills = st.get('F', {}).get('fills', [])
        extra_rows = 0
        for opt in ['A', 'B', 'C', 'D', 'E', 'F']:
            fills = st.get(opt, {}).get('fills', [])
            if fills:
                extra_rows += 1

        table = doc.add_table(rows=8 + extra_rows, cols=2, style='Table Grid')
        table.autofit = True

        headers = ['选项', '小计']
        for ci, h in enumerate(headers):
            cell = table.cell(0, ci)
            cell.text = h
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(10.5)
                    r.font.name = '宋体'
                    r._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
                    r.font.bold = True

        current_row = 1
        for opt in ['A', 'B', 'C', 'D', 'E', 'F']:
            cnt = st.get(opt, {}).get('count', 0)
            fills = st.get(opt, {}).get('fills', [])

            if fills:
                # 有填空：加额外行
                # 选项行
                row = table.rows[current_row]
                label = opt
                if opt == 'F':
                    label = 'F（填空）'
                row.cells[0].text = label
                row.cells[1].text = str(cnt)
                for ci in range(2):
                    for p in row.cells[ci].paragraphs:
                        for r in p.runs:
                            r.font.size = Pt(10.5)
                            r.font.name = '宋体'
                            r._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
                current_row += 1

                # 填空行
                fill_row = table.rows[current_row]
                fill_row.cells[0].text = ''
                fill_row.cells[1].text = '; '.join(fills)
                for ci in range(2):
                    for p in fill_row.cells[ci].paragraphs:
                        for r in p.runs:
                            r.font.size = Pt(10.5)
                            r.font.name = '宋体'
                            r._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
                current_row += 1
            else:
                row = table.rows[current_row]
                label = opt
                if opt == 'F':
                    label = 'F（填空）'
                row.cells[0].text = label
                row.cells[1].text = str(cnt)
                for ci in range(2):
                    for p in row.cells[ci].paragraphs:
                        for r in p.runs:
                            r.font.size = Pt(10.5)
                            r.font.name = '宋体'
                            r._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
                current_row += 1

        # 本题有效填写人次
        last_row = table.rows[current_row]
        last_row.cells[0].text = '本题有效填写人次'
        last_row.cells[1].text = str(total)
        for ci in range(2):
            for p in last_row.cells[ci].paragraphs:
                for r in p.runs:
                    r.font.size = Pt(10.5)
                    r.font.name = '宋体'
                    r._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    # 空行
    doc.add_paragraph()

doc.save(OUTPUT_WORD)
doc.save(OUTPUT_WORD_RESULT)
print(f"Word文档已保存")

# ===============================
# 打印摘要
# ===============================
print(f"\n=== 统计摘要 ===")
for qnum, qtype in QUESTION_META:
    st = stats.get(qnum, {})
    total = len(df) - st.get('_MISSING', {}).get('count', 0)
    missing = st.get('_MISSING', {}).get('count', 0)
    parts = []
    for opt in ['A', 'B', 'C', 'D', 'E', 'F']:
        cnt = st.get(opt, {}).get('count', 0)
        if cnt > 0:
            fill_info = ''
            fills = st.get(opt, {}).get('fills', [])
            if fills:
                fill_info = f'(填空:{"; ".join(fills[:2])})'
            parts.append(f'{opt}:{cnt}{fill_info}')
    missing_str = f'  [缺{missing}]' if missing else ''
    type_str = '单选' if qtype == 'single' else '多选'
    sep = ' | '
    print(f'Q{qnum}({total}人{missing_str}): {sep.join(parts)}')