# ==============================
# 改编版：适配2026秋季问卷(总第27次)
# 将Word问卷提取为 id-题号 格式的Excel
# 
# 与原版差异：
#   1. 题号结构变化 (8→9, 9→10, 13→14, 15→16, 新增第19题)
#   2. 全部.docx格式，保留.doc兼容
#   3. PDF需手动统计
# ==============================

import re
import pandas as pd
from docx import Document
import os
import sys

# 修复 Windows GBK 终端编码问题
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# win32com 仅用于 .doc → .docx 转换
try:
    import win32com.client as win32
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False


# ==============================
# doc 转 docx (保留兼容)
# ==============================
def convert_doc_to_docx(word, doc_path, output_folder):
    doc = word.Documents.Open(doc_path)
    filename = os.path.basename(doc_path)
    new_name = filename + "x"
    new_path = os.path.join(output_folder, new_name)
    doc.SaveAs(new_path, FileFormat=16)
    doc.Close(False)
    return new_path


# ==============================
# 分割题目 — 适配2026秋季问卷
# ==============================
def split_questions(docx_path):
    """
    新版问卷结构:
      1-8:   单选题
      9-1~9-4:  父题9 (固定资产投资等)
      10-1,10-2: 父题10 (国有/民间投资)
      11-13:  单选题
      14-1~14-3: 父题14 (进出口)
      15:     单选题
      16-1~16-4: 父题16 (宏观政策/M2等)
      17-19:  多选题
    """
    doc = Document(docx_path)
    paragraphs = [para.text.strip() for para in doc.paragraphs if para.text.strip()]

    # 拆分关键词: 单选/多选标记 + 子题号
    split_keywords = [
        "（单选或填空）",
        "（多选或填空）",
        "9-1", "9-2", "9-3", "9-4",
        "10-1", "10-2",
        "14-1", "14-2", "14-3",
        "16-1", "16-2", "16-3", "16-4",
    ]

    # 这些是父题号，不提取答案（由其子题代替）
    main_numbers_to_filter = ['9', '10', '14', '16']

    questions = []
    current_question = []
    current_number = None
    skip_current = False

    for para in paragraphs:
        contains_keyword = any(keyword in para for keyword in split_keywords)

        if contains_keyword:
            if current_question and not skip_current:
                questions.append({
                    'number': current_number,
                    'text': '\n'.join(current_question)
                })

            current_question = []
            skip_current = False
            current_question.append(para)

            match = re.match(r'^(\d+(?:-\d+)?)\.', para)
            if match:
                current_number = match.group(1)
                if current_number in main_numbers_to_filter:
                    skip_current = True
                    current_question = []
            else:
                for keyword in split_keywords[2:]:
                    if keyword in para:
                        current_number = keyword
                        break
        else:
            if current_question and not skip_current:
                current_question.append(para)

    if current_question and not skip_current:
        questions.append({
            'number': current_number,
            'text': '\n'.join(current_question)
        })

    return questions


# ==============================
# 提取选项 — 兼容"单选或填空"中带填空的情况
# ==============================
def extract_options_from_questions(questions):
    results = {}

    for q in questions:
        number = q['number']
        text = q['text']

        # 匹配括号内的选项，支持中文括号（）和英文括号()
        # 也处理 "（  C  ）" 这种带空格的情况
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
# 主程序
# ==============================
def main():
    # ====== 修改这里的路径 ======
    folder_path = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\2026秋问卷回收（88份）"

    output_folder = os.path.join(folder_path, "processed")
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    files = os.listdir(folder_path)
    word_files = [f for f in files
                  if f.endswith((".doc", ".docx"))
                  and not f.startswith("~$")]  # 排除临时文件

    print(f"发现Word文件数量: {len(word_files)}")

    # 只对所有docx文件不需Word COM
    has_doc = any(f.endswith(".doc") for f in word_files)

    word = None
    if has_doc and HAS_WIN32:
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = 0
        print("检测到.doc文件，已启动Word转换引擎")
    elif has_doc and not HAS_WIN32:
        print("警告: 有.doc文件但未安装pywin32，将跳过.doc文件")

    all_results = []

    for file in word_files:
        print(f"正在处理: {file}")

        file_path = os.path.join(folder_path, file)

        try:
            if file.endswith(".doc") and word:
                print("  转换 doc → docx")
                file_path = convert_doc_to_docx(word, file_path, output_folder)

            questions = split_questions(file_path)
            results = extract_options_from_questions(questions)

            file_id = os.path.splitext(file)[0]
            results["ID"] = file_id
            all_results.append(results)

        except Exception as e:
            print(f"  处理失败: {file}")
            print(f"  错误: {e}")

    if word:
        word.Quit()

    # 构建DataFrame
    df = pd.DataFrame(all_results)

    # ID放第一列
    cols = df.columns.tolist()
    if "ID" in cols:
        cols.remove("ID")
        cols = ["ID"] + cols
        df = df[cols]

    output_excel = os.path.join(output_folder, "问卷汇总结果.xlsx")
    df.to_excel(output_excel, index=False)

    print(f"\n全部完成! 共处理 {len(all_results)} 份问卷")
    print(f"输出位置: {output_excel}")
    print(f"\n列名: {[c for c in df.columns if c != 'ID']}")


if __name__ == "__main__":
    main()