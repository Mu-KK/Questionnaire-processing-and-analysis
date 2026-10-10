# ==============================
# 该脚本用于将queations_statistic.py文件输出的excel（经人工核对后）进行统计，将输出一个excel,统计每题每个选项有多少人选择
# 注  只需修改10-16行文件路径即可
# ==============================


import pandas as pd
import os

# ===============================
# 输入文件路径
# ===============================
input_file = r"C:\Users\Administrator\Desktop\问卷回收1-0317\processed\问卷汇总结果.xlsx"

# 输出文件
output_file = r"C:\Users\Administrator\Desktop\问卷回收1-0317\processed\问卷选项统计结果.xlsx"


# ===============================
# 读取数据
# ===============================
df = pd.read_excel(input_file)

# 去掉ID列
question_cols = [col for col in df.columns if col != "ID"]

all_results = []

# ===============================
# 逐题统计
# ===============================
for q in question_cols:

    answers = df[q].dropna().astype(str)

    option_count = {}

    for ans in answers:

        # 多选题拆分
        parts = ans.split(",")

        for p in parts:

            p = p.strip()

            if p == "":
                continue

            if p not in option_count:
                option_count[p] = 0

            option_count[p] += 1

    for option, count in option_count.items():

        all_results.append({
            "题号": q,
            "选项": option,
            "人数": count
        })


# ===============================
# 生成统计表
# ===============================
result_df = pd.DataFrame(all_results)

# 按题号排序
result_df = result_df.sort_values(["题号", "选项"])

# 保存Excel
result_df.to_excel(output_file, index=False)

print("统计完成")
print("输出文件：")
print(output_file)