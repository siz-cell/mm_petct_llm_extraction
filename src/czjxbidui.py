import pandas as pd
import json



from datetime import datetime
import re
from typing import Dict, Tuple



def save_comparisons_to_json(comparisons, filename=None):
    """将比较结果保存为JSON文件"""
    if filename is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"comparisons_{timestamp}.json"
    
    # 准备保存的数据
    save_data = {
        'comparisons': comparisons
    }
    
    with open(filename, 'a', encoding='utf-8') as f:
        json.dump(save_data, f, ensure_ascii=False, indent=2)
    
    print(f"比较结果已保存到: {filename}")
    return filename

def extract_table_data(text: str) -> pd.DataFrame:
    """从文本中提取表格数据，遇到非表格行就停止"""
    lines = text.strip().split('\n')
    data = []
    table_ended = False
    
    for line in lines:
        line = line.strip()
        
        # 如果已经遇到非表格行，直接跳过后续所有行
        if table_ended:
            continue
            
        # 跳过分隔线
        if (re.search(r'^\|[-|\s]+\|$', line) or 
            re.search(r'^-+$', line) or
            re.search(r'^\|[\s|-]+\|$', line)):
            continue
        
        # 如果是表格行
        if line.startswith('|') and line.endswith('|'):
            cells = [cell.strip() for cell in line[1:-1].split('|')]
            data.append(cells)
        else:
            # 遇到非表格行，标记表格结束
            table_ended = True
    
    return pd.DataFrame(data)

def normalize_value(value: str) -> str:
    """标准化数值，处理NA、空格等"""
    if pd.isna(value) or value == 'NA' or value == '':
        return 'NA'
    return str(value).strip().lower()

def compare_tables(excel_row: str, jsonl_md: str) -> Dict:
    """
    比较两个表格的内容并计算评估指标
    
    Args:
        excel_row: Excel表格文本
        jsonl_md: JSONL标记表格文本
        
    Returns:
        Dict: 包含比较结果和评估指标的字典
    """
    # 提取表格数据
    df_excel = extract_table_data(excel_row)
    df_jsonl = extract_table_data(jsonl_md)
    
    # 确保两个DataFrame有相同的形状
    max_rows = max(len(df_excel), len(df_jsonl))
    max_cols = max(len(df_excel.columns) if len(df_excel.columns) > 0 else 0, 
                   len(df_jsonl.columns) if len(df_jsonl.columns) > 0 else 0)
    
    # 重新索引以确保形状一致
    df_excel = df_excel.reindex(range(max_rows), fill_value='')
    df_jsonl = df_jsonl.reindex(range(max_rows), fill_value='')
    
    if max_cols > 0:
        df_excel = df_excel.reindex(columns=range(max_cols), fill_value='')
        df_jsonl = df_jsonl.reindex(columns=range(max_cols), fill_value='')
    

    
        # 比较每个单元格
    comparison_cells = [
    (1, 0),  (1, 1),  (1, 2),  (1, 3),  (1, 4),  (1, 5),  (1, 6)
    ]
        # 单元格位置描述映射
    position_descriptions = {
        (1, 0): "骨髓浸润与骨质破坏模式",
        (1, 1): "局灶性病灶绝对数目", 
        (1, 2): "骨折情况",
        (1, 3): "髓外病变",
        (1, 4): "髓旁病变",
        (1, 5): "长骨浸润",
        (1, 6): "骨骼手术证据", 
        
    }
    position_descriptionsbidui = {
        (1, 0): "骨髓浸润与骨质破坏模式",
        (1, 1): "局灶性病灶绝对数目", 
        (1, 2): "骨折情况",
        (1, 3): "髓外病变",
        (1, 4): "髓旁病变",
        (1, 5): "长骨浸润",
        (1, 6): "骨骼手术证据", 
        
    }
    # 比较特定单元格（带描述）
    comparisons = []
    true_positives = 0
    false_positives = 0
    false_negatives = 0
    true_negatives = 0

    for (i, j) in comparison_cells:
        row_idx = i 
        col_idx = j 
        
        excel_val = ''
        jsonl_val = ''
        
        # 安全地获取Excel值
        if (row_idx < len(df_excel) and col_idx < len(df_excel.columns) and 
            pd.notna(df_excel.iloc[row_idx, col_idx])):
            excel_val = normalize_value(df_excel.iloc[row_idx, col_idx])
        
        # 安全地获取JSONL值
        if (row_idx < len(df_jsonl) and col_idx < len(df_jsonl.columns) and 
            pd.notna(df_jsonl.iloc[row_idx, col_idx])):
            jsonl_val = normalize_value(df_jsonl.iloc[row_idx, col_idx])
        
        is_match = excel_val == jsonl_val
        
        # 计算分类指标
        
        if excel_val != 'NA' and excel_val != '' and jsonl_val != 'NA' and jsonl_val != '':
            if is_match:
                if excel_val not in ['', 'na']:
                    true_positives += 1
                else:
                    true_negatives += 1
            else:
                if excel_val not in ['', 'na'] and jsonl_val in ['', 'na']:
                    false_negatives += 1
                elif excel_val in ['', 'na'] and jsonl_val not in ['', 'na']:
                    false_positives += 1
                else:
                    false_positives += 0.5
                    false_negatives += 0.5
        
        description = position_descriptions.get((i, j), f"位置({i},{j})")
        comparisons.append({
            'position': (i, j),
            'description': description,
            'excel': excel_val,
            'jsonl': jsonl_val,
            'match': is_match
        })
        # 在循环结束后统计
    match_count = sum(1 for item in comparisons if (item['position'] in position_descriptionsbidui and item['match']))
    total_count = len(position_descriptionsbidui)
    match_ratio = match_count / total_count

    print(f"匹配的单元格数量: {match_count}/{total_count}")
    print(f"匹配率: {match_ratio:.2%}")
    return comparisons


def read_multiple_jsonl_files(file_paths):
    """读取多个JSONL文件并合并内容"""
    all_lines = []
    for file_path in file_paths:
        with open(file_path, 'r', encoding='utf-8') as f:
            all_lines.extend(f.readlines())
    return all_lines


file_paths = [
    './10112/bidui/jiegouh/o4minihigh_results0.jsonl',
    './10112/bidui/jiegouh/o4minihigh_results0 copy.jsonl'
]

#all_lines = read_multiple_jsonl_files(file_paths)
# 逐行同步处理
df_excel = pd.read_excel(r"C:\Users\User\Desktop\pythondata\预后模型评估\10112\bidui\jiegouh\多模型评测\结构化表格\外部验证集评测\外部验证金标准.xlsx",sheet_name="Sheet2")
with open(r'C:\Users\User\Desktop\pythondata\预后模型评估\10112\bidui\jiegouh\多模型评测\结构化表格\外部验证集评测\qwen3235b8shotwucot163tiao初诊基线表格输出.jsonl', 'r', encoding='utf-8') as f:
    for idx, line in enumerate(f):
            # 读取 Excel 的对应行
            excel_row = df_excel.iloc[idx]
            excel_content = excel_row['初诊基线表格表格'].strip()
            jsonl_row = json.loads(line.strip())
            #  jsonl_md = jsonl_row['response']['body']['choices'][0]['message']['content']
            jsonl_md = jsonl_row['response']
            print(f"行号: {idx}")
            print(f"Excel: {excel_row['初诊基线表格表格']}")  # 替换为实际列名
            print(f"JSONL: {jsonl_md}")
            result = compare_tables(excel_content,jsonl_md)
            save_comparisons_to_json(result,r"C:\Users\User\Desktop\pythondata\预后模型评估\10112\bidui\jiegouh\多模型评测\结构化表格\外部验证集评测\qwen3235b8shotwucot163tiao初诊基线表格比对.jsonl")

            #new_text = re.sub(r'"response":\s*".*?</think>', '"response": "', text)