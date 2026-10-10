# ==============================
# count_v2.py - 统计问卷结果
# 适配2026秋季问卷(总第27次，88份)
# ==============================

import pandas as pd
import os, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# ===============================
# 输入输出路径
# ===============================
input_file = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\2026秋问卷回收（88份）\processed\问卷汇总结果_已校对.xlsx"
output_excel = r"C:\Users\LingchenKang\Desktop\ycl问卷处理\问卷处理脚本及结果\2026秋问卷回收（88份）\processed\问卷选项统计结果.xlsx"

# ===============================
df = pd.read_excel(input_file)
question_cols = [col for col in df.columns if col != "ID"]

print(f"共 {len(df)} 份问卷, {len(question_cols)} 道题")
print(f"题目列表: {question_cols}")

all_results = []

for q in question_cols:
    answers = df[q].dropna().astype(str)
    answers = answers[answers != '']  # 排除空字符串

    option_count = {}

    for ans in answers:
        # 多选题拆分（逗号分隔）
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

result_df = pd.DataFrame(all_results)
result_df = result_df.sort_values(["题号", "选项"])

# 保存
result_df.to_excel(output_excel, index=False)

print(f"\n统计完成!")
print(f"输出文件: {output_excel}")

# ===============================
# 打印摘要
# ===============================
print(f"\n=== 每题选项人数摘要 ===")
for q in question_cols:
    sub = result_df[result_df['题号'] == q]
    if len(sub) > 0:
        summary = ', '.join([f"{r['选项']}:{int(r['人数'])}" for _, r in sub.iterrows()])
        total = int(sub['人数'].sum())
        print(f"  Q{q} ({total}人): {summary}")
    else:
        print(f"  Q{q}: 无数据")