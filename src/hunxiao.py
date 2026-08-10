
from sklearn.metrics import confusion_matrix

import json
import numpy as np
import pandas as pd 
df_excel = pd.read_excel(r'C:\Users\User\Desktop\pythondata\预后模型评估\10112\bidui\jiegouh\多模型评测\随访评估表\suifang结果1.xlsx')

# 示例数据
y_true =df_excel["金标准结果"]  # 实际标签
 # 预测标签

data=[]
with open(r'C:\Users\User\Desktop\pythondata\预后模型评估\10112\bidui\jiegouh\多模型评测\随访评估表\qwen34b随访输出.jsonl', 'r', encoding='utf-8') as f:
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
y_pred = df_jsonl["response"]

# 创建混淆矩阵
cm = confusion_matrix(y_true, y_pred, labels=['缓解', '进展', '稳定'])

# 转换为DataFrame显示
cm_df = pd.DataFrame(cm,
                     index=['实际-缓解', '实际-进展', '实际-稳定'],
                     columns=['预测-缓解', '预测-进展', '预测-稳定'])

print("三分类混淆矩阵:")
print(cm_df)
print()

# 提取各个值（现在是3×3矩阵）
# 定义别名方便理解
classes = ['缓解', '进展', '稳定']
class_idx = {cls: i for i, cls in enumerate(classes)}

# 提取每个类别的值
print("=== 各类别统计 ===")
for i, true_class in enumerate(classes):
    TP = cm[i, i]  # 该类别的真正例
    # 假负例：实际是这类，但预测为其他类别的总数
    FN = np.sum(cm[i, :]) - TP
    # 假正例：预测是这类，但实际是其他类别的总数
    FP = np.sum(cm[:, i]) - TP
    # 真负例：实际不是这类，预测也不是这类
    TN = np.sum(cm) - (TP + FP + FN)
    
    print(f"\n{true_class}类:")
    print(f"  TP (真正例): {TP} ")
    print(f"  FN (假负例): {FN} ")
    print(f"  FP (假正例): {FP} ")
    print(f"  TN (真负例): {TN} ")
    
    # 计算指标
    precision = TP / (TP + FP) if (TP + FP) > 0 else 0
    recall = TP / (TP + FN) if (TP + FN) > 0 else 0
    f1 = 2 * precision * recall / (precision + recall) if (precision + recall) > 0 else 0
    
    print(f"  精确率 (Precision): {precision:.3f}")
    print(f"  召回率 (Recall): {recall:.3f}")
    print(f"  F1分数: {f1:.3f}")