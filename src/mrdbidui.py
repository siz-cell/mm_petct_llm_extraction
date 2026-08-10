
import json

import numpy as np
import pandas as pd
from sklearn.metrics import confusion_matrix


df_excel = pd.read_excel(r'C:\Users\User\Desktop\pythondata\预后模型评估\10112\bidui\jiegouh\多模型评测\mrd表格\病灶明细表与结构化表格初诊基线MRD金标准_cleaned_fixed copy.xlsx')
data=[]
with open(r'C:\Users\User\Desktop\pythondata\预后模型评估\10112\bidui\jiegouh\多模型评测\mrd表格\o4minimrd输出.jsonl', 'r', encoding='utf-8') as f:
    for idx,line in enumerate(f):
        line = line.strip()
        if line:  # 跳过空行
            try:
                json_obj = json.loads(line)
                data.append(json_obj)
            except json.JSONDecodeError as e:
                print(f"第{idx+1}行JSON解析错误: {e}")
                print(f"错误行内容: {line[:100]}...")
df_full = pd.DataFrame(data) 
df_jsonl =  df_full[["response"]] 
print(f"成功读取 {len(df_jsonl)} 行数据")
print(f"DataFrame形状: {df_jsonl.shape}")


print("dfjosnl.shape[0]",df_jsonl.shape[0])
print("dfjosnl.shape[0]",df_jsonl.head(10))
print("dfjosnl.shape[0]",df_excel.shape[0])
print("dfjosnl.shape[0]",df_excel['MRD'].tail(242).head(10))

excel_last_242 = df_excel['MRD'].tail(242)
y_pred = df_jsonl["response"]
cm = confusion_matrix(y_true=excel_last_242, y_pred=y_pred, labels=['是', '否'])


# 创建可读性更好的DataFrame
cm_df = pd.DataFrame(cm,
                     index=['实际是', '实际否'],
                     columns=['预测是', '预测否'])

print("混淆矩阵:")
print(cm_df)
print()

# 提取各个值
TP = cm[0, 0]  # 第一行第一列
FN = cm[0, 1]  # 第一行第二列
FP = cm[1, 0]  # 第二行第一列
TN = cm[1, 1]  # 第二行第二列

print(f"TP (真正例): {TP} - 实际是，预测是")
print(f"FN (假负例): {FN} - 实际是，预测否")
print(f"FP (假正例): {FP} - 实际否，预测是")
print(f"TN (真负例): {TN} - 实际否，预测否")