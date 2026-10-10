# ==============================
# 处理4份PDF问卷，追加到已校对的Excel
# ==============================

import re
import os, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import pdfplumber
import pandas as pd
from docx import Document

PDF_DIR = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\2026秋问卷回收（88份）"
EXCEL_PATH = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\2026秋问卷回收（88份）\processed\问卷汇总结果_已校对.xlsx"

# ==============================
# 分割题目（复用v2逻辑）
# ==============================
# PDF中可能出现 (X. 单选或填空) 或 (X.（单选或填空）) 等变体
# 用更宽松的正则匹配
split_keywords = [
    "（单选或填空）",
    "（多选或填空）",
    "9-1", "9-2", "9-3", "9-4",
    "10-1", "10-2",
    "14-1", "14-2", "14-3",
    "16-1", "16-2", "16-3", "16-4",
]
# 额外的宽松匹配：处理 "X（. 单选或填空）" 这种PDF变体
loose_kw_pattern = re.compile(r'[（(].*?单选或填空.*?[）)]|[（(].*?多选或填空.*?[）)]')
main_numbers_to_filter = ['9', '10', '14', '16']


def split_questions_text(lines):
    """从文本行列表中分割题目"""
    questions = []
    current_question = []
    current_number = None
    skip_current = False

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # 去掉页码和页眉噪声
        if re.match(r'^\d+/\d+$', line):
            continue
        if 'EID·CMR' in line or '2026年中国宏观经济形势和政策问卷调查' in line:
            continue

        contains_keyword = any(keyword in line for keyword in split_keywords) or bool(loose_kw_pattern.search(line))

        if contains_keyword:
            if current_question and not skip_current:
                questions.append({
                    'number': current_number,
                    'text': '\n'.join(current_question)
                })

            current_question = []
            skip_current = False
            current_question.append(line)

            # PDF中可能格式: "1（. 单选" 或 "1.（单选" 或 "14-1. xxx"
            match = re.match(r'^(\d+(?:-\d+)?)[\.\（\s]', line)
            if match:
                current_number = match.group(1)
                if current_number in main_numbers_to_filter:
                    skip_current = True
                    current_question = []
            else:
                for keyword in split_keywords[2:]:
                    if keyword in line:
                        current_number = keyword
                        break
        else:
            if current_question and not skip_current:
                current_question.append(line)

    if current_question and not skip_current:
        questions.append({
            'number': current_number,
            'text': '\n'.join(current_question)
        })

    return questions


def extract_options_from_questions(questions):
    results = {}
    for q in questions:
        number = q['number']
        text = q['text']
        bracket_pattern = r'[（(]\s*([^）)]+?)\s*[）)]'
        matches = re.findall(bracket_pattern, text)
        if matches:
            option_value = matches[-1].strip()
            option_value = re.sub(r'\s+', ' ', option_value)
            option_value = option_value.replace('、', ',')
            results[number] = option_value
        else:
            results[number] = ""
    return results


# ==============================
# 处理PDF文件
# ==============================
# 读入现有Excel
df = pd.read_excel(EXCEL_PATH)

pdf_files = sorted([f for f in os.listdir(PDF_DIR) if f.endswith('.pdf')])
print(f"找到 {len(pdf_files)} 个PDF文件")

for pf in pdf_files:
    path = os.path.join(PDF_DIR, pf)
    print(f"\n处理: {pf}")

    with pdfplumber.open(path) as pdf:
        full_text = ''
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                full_text += t + '\n'

    lines = full_text.split('\n')
    questions = split_questions_text(lines)
    results = extract_options_from_questions(questions)

    file_id = os.path.splitext(pf)[0]

    print(f"  解析到 {len(results)} 题")
    for qn, ans in sorted(results.items(), key=lambda x: (x[0].split('-')[0], x[0])):
        if ans:
            print(f"  Q{qn}: {ans}")
        else:
            print(f"  Q{qn}: 【空】")

    # 追加到DataFrame
    row = {'ID': file_id}
    row.update(results)
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)

# ==============================
# 保存
# ==============================
# 确保ID列在第一位
cols = df.columns.tolist()
if 'ID' in cols:
    cols.remove('ID')
    cols = ['ID'] + cols
    df = df[cols]

df.to_excel(EXCEL_PATH, index=False)

print(f"\n=== 完成 ===")
print(f"原有 {len(df) - len(pdf_files)} 份 + 新增 {len(pdf_files)} 份PDF = 共 {len(df)} 份")
print(f"已保存到: {EXCEL_PATH}")

# 快速检查缺失
qcols = [c for c in df.columns if c != 'ID']
print(f"\n=== 最终缺失情况 ===")
for c in qcols:
    missing = df[c].isna().sum() + (df[c] == '').sum()
    if missing > 0:
        print(f"  Q{c}: {missing}/{len(df)} 缺失")