# ==============================
# 该脚本用于将以word形式提交的问卷结果转换为 id-题号 输出格式的excel，方便后续进行问卷结果统计
# 注  如不需要修改问卷切割形式，只需修改第144行      folder_path = r%(问卷存放文件夹路径)%
# 如需要修改问卷切割形式，请修改44-53行
# 由于回收结果中pdf格式数量很少，本脚本未加入识别pdf格式代码，请在使用后手动统计pdf格式问卷的结果
# ==============================

import re
import pandas as pd
from docx import Document
import os
import win32com.client as win32


# ==============================
# doc 转 docx
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
# 分割题目
# ==============================
def split_questions_simple(docx_path):

    doc = Document(docx_path)

    paragraphs = [para.text.strip() for para in doc.paragraphs if para.text.strip()]

    split_keywords = [
        "（单选或填空）",
        "（多选或填空）",
        "8-1","8-2","8-3","8-4",
        "9-1","9-2",
        "13-1","13-2","13-3",
        "15-1","15-2","15-3","15-4"
    ]

    main_numbers_to_filter = ['8','9','13','15']

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
# 提取选项
# ==============================
def extract_options_from_questions(questions):

    results = {}

    for q in questions:

        number = q['number']
        text = q['text']

        bracket_pattern = r'[（(]([^）)]+)[）)]'
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

    folder_path = r"C:\Users\Administrator\Desktop\问卷回收1-0317"

    # 新建输出文件夹
    output_folder = os.path.join(folder_path, "processed")

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    files = os.listdir(folder_path)

    word_files = [f for f in files if f.endswith(".doc") or f.endswith(".docx")]

    print("发现Word文件数量:", len(word_files))

    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0

    all_results = []

    for file in word_files:

        print("正在处理:", file)

        file_path = os.path.join(folder_path, file)

        try:

            # doc 转 docx
            if file.endswith(".doc"):

                print("转换 doc → docx")

                file_path = convert_doc_to_docx(word, file_path, output_folder)

            questions = split_questions_simple(file_path)

            results = extract_options_from_questions(questions)

            file_id = os.path.splitext(file)[0]

            results["ID"] = file_id

            all_results.append(results)

        except Exception as e:

            print("处理失败:", file)
            print(e)

    word.Quit()

    df = pd.DataFrame(all_results)

    cols = df.columns.tolist()

    if "ID" in cols:
        cols.remove("ID")
        cols = ["ID"] + cols
        df = df[cols]

    output_excel = os.path.join(output_folder, "问卷汇总结果.xlsx")

    df.to_excel(output_excel, index=False)

    print("\n全部完成")
    print("Excel位置:", output_excel)


# ==============================
# 运行
# ==============================
if __name__ == "__main__":
    main()