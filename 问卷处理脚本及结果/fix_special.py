# ==============================
# 修复脚本：处理特殊格式的问卷
# 在新版queations_statistic_v2.py结果基础上
#   - 鄢萍: "第 X 题：Y" 纯答案格式
#   - 贺力平: 用"√"标记答案
#   - 王曦: 答案在选项末尾（手动处理）
# ==============================

import pandas as pd
import re
import os, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

INPUT = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\2026秋问卷回收（88份）\processed\问卷汇总结果.xlsx"
OUTPUT_CORRECTED = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\2026秋问卷回收（88份）\processed\问卷汇总结果_已校对.xlsx"

df = pd.read_excel(INPUT)
qcols = [c for c in df.columns if c != 'ID']

# ==========================================
# 1. 修复鄢萍 - 纯答案格式
# ==========================================
print("=== 修复鄢萍 ===")
from docx import Document

yanping_file = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\2026秋问卷回收（88份）\鄢萍 选择题答案.docx"
doc = Document(yanping_file)
yanping_answers = {}
for p in doc.paragraphs:
    t = p.text.strip()
    m = re.match(r'第\s*(\d+(?:-\d+)?)\s*题[：:]\s*(.+)', t)
    if m:
        qnum = m.group(1)
        ans = m.group(2).strip()
        yanping_answers[qnum] = ans

print(f"  解析到 {len(yanping_answers)} 题答案")

# 找到鄢萍的行并填充
mask = df['ID'].str.contains('鄢萍', na=False)
if mask.any():
    for qnum, ans in yanping_answers.items():
        if qnum in df.columns:
            df.loc[mask, qnum] = ans
    print(f"  已填充到: {list(yanping_answers.keys())}")


# ==========================================
# 2. 修复贺力平 - √标记格式
# ==========================================
print("=== 修复贺力平 ===")

helping_file = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\2026秋问卷回收（88份）\贺力平 2026年秋季中国宏观经济形势与政策调查问卷(总第27次）.docx"
doc = Document(helping_file)

# 构建分段结构：找到每个以 "(单选或填空)" 或 "(多选或填空)" 或子题号开头的段落作为分割点
helping_questions = {}
current_q = None
current_opts = []

split_kw = ["（单选或填空）", "（多选或填空）",
            "9-1", "9-2", "9-3", "9-4",
            "10-1", "10-2",
            "14-1", "14-2", "14-3",
            "16-1", "16-2", "16-3", "16-4"]

for i, para in enumerate(doc.paragraphs):
    t = para.text.strip()
    if not t:
        continue

    # 检查是否是新的题目开始
    is_new_q = False
    q_num = None
    for kw in split_kw:
        if kw in t:
            is_new_q = True
            # extract question number
            m = re.match(r'^(\d+(?:-\d+)?)[\.\）)]', t)
            if m:
                q_num = m.group(1)
            else:
                # check if kw itself is the q number
                for k in split_kw[2:]:
                    if k in t:
                        q_num = k
                        break
            break

    if is_new_q:
        # 保存上一题
        if current_q and current_opts:
            # find √ in options
            for opt in reversed(current_opts):
                if '√' in opt:
                    m = re.match(r'([A-F])[\.\s]', opt)
                    if m:
                        if current_q in helping_questions:
                            # multi-select: accumulate
                            if m.group(1) not in helping_questions[current_q]:
                                helping_questions[current_q] += ',' + m.group(1)
                        else:
                            helping_questions[current_q] = m.group(1)
                        break

        current_q = q_num
        current_opts = [t]
    else:
        if current_q:
            current_opts.append(t)

# 最后一题
if current_q and current_opts:
    for opt in reversed(current_opts):
        if '√' in opt:
            m = re.match(r'([A-F])[\.\s]', opt)
            if m:
                if current_q in helping_questions:
                    if m.group(1) not in helping_questions[current_q]:
                        helping_questions[current_q] += ',' + m.group(1)
                else:
                    helping_questions[current_q] = m.group(1)
                break

print(f"  解析到 {len(helping_questions)} 题答案: {helping_questions}")

# 填充到 dataframe
mask = df['ID'].str.contains('贺力平', na=False)
if mask.any():
    for qnum, ans in helping_questions.items():
        if qnum in df.columns:
            df.loc[mask, qnum] = ans
    print(f"  已填充")


# ==========================================
# 3. 王曦 - 尝试修复（答案在选项末尾）
# ==========================================
print("=== 尝试修复王曦 ===")

wangxi_file = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\2026秋问卷回收（88份）\王曦 2026年秋季中国宏观经济形势与政策调查问卷(总第27次）.docx"
doc = Document(wangxi_file)

wangxi_answers = {}
current_q = None
current_opts = []

for i, para in enumerate(doc.paragraphs):
    t = para.text.strip()
    if not t:
        continue

    is_new_q = False
    q_num = None
    for kw in split_kw:
        if kw in t:
            is_new_q = True
            m = re.match(r'^(\d+(?:-\d+)?)[\.\）)]', t)
            if m:
                q_num = m.group(1)
            else:
                for k in split_kw[2:]:
                    if k in t:
                        q_num = k
                        break
            break

    if is_new_q:
        if current_q and current_opts:
            # 找最后一个选项行末尾是否有额外字母
            for opt in reversed(current_opts):
                # 检查是否像 "F. 其他    C" 这样在选项后有单独字母
                # 先匹配选项格式，再看后面是否有多余内容
                m_opt = re.match(r'([A-F])[\.\s、]\s*(.+?)\s+([A-F]\b)$', opt)
                if m_opt:
                    # 选项后明确有答案如 "F. 其他    C"
                    wangxi_answers[current_q] = m_opt.group(3)
                    break
                m_opt2 = re.match(r'([A-F])[\.\s、]', opt)
                if m_opt2:
                    # 正常选项，检查后面内容
                    rest = opt[len(m_opt2.group(0)):].strip()
                    # 如果剩余部分是单个大写字母
                    if re.match(r'^[A-F]\s*$', rest):
                        wangxi_answers[current_q] = rest.strip()
                        break

        current_q = q_num
        current_opts = [t]
    else:
        if current_q:
            current_opts.append(t)

if current_q and current_opts:
    for opt in reversed(current_opts):
        m_opt = re.match(r'([A-F])[\.\s、]\s*(.+?)\s+([A-F]\b)$', opt)
        if m_opt:
            wangxi_answers[current_q] = m_opt.group(3)
            break

print(f"  解析到 {len(wangxi_answers)} 题答案: {wangxi_answers}")

mask = df['ID'].str.contains('王曦', na=False)
if mask.any():
    for qnum, ans in wangxi_answers.items():
        if qnum in df.columns:
            df.loc[mask, qnum] = ans
    print(f"  已填充")


# ==========================================
# 4. 保存校对结果
# ==========================================
df.to_excel(OUTPUT_CORRECTED, index=False)

# 输出统计
print()
print("=== 修复后缺失统计 ===")
for c in qcols:
    missing = df[c].isna().sum() + (df[c] == '').sum()
    if missing > 0:
        print(f"  Q{c}: {missing} 缺失")

print(f"\n已保存到: {OUTPUT_CORRECTED}")