import pandas as pd
import json
pd.set_option('display.max_colwidth', None)  # 或设置一个较大的值
pd.set_option('display.max_rows', None)      # 显示所有行

df_excel_og = pd.read_excel(r'C:\Users\User\Desktop\pythondata\预后模型评估\10112\bidui\jiegouh\多模型评测\mrd表格\病灶明细表与结构化表格初诊基线MRD金标准_cleaned_fixed copy.xlsx')


data = []
with open(r'C:\Users\User\Desktop\pythondata\预后模型评估\10112\bidui\jiegouh\多模型评测\结构化表格\gemini25flash结构化表格输出.jsonl', 'r', encoding='utf-8') as f:
    for idx,line in enumerate(f):
        line = line.strip()
        if line:  # 跳过空行
            try:
                json_obj = json.loads(line)
                data.append(json_obj)
            except json.JSONDecodeError as e:
                print(f"第{idx+1}行JSON解析错误: {e}")
                print(f"错误行内容: {line[:100]}...")
df_jghbg = pd.DataFrame(data) 

data=[]
with open(r'C:\Users\User\Desktop\pythondata\预后模型评估\10112\bidui\jiegouh\多模型评测\mrd表格\gemini25flashmrd输出.jsonl', 'r', encoding='utf-8') as f:
    for idx,line in enumerate(f):
        line = line.strip()
        if line:  # 跳过空行
            try:
                json_obj = json.loads(line)
                data.append(json_obj)
            except json.JSONDecodeError as e:
                print(f"第{idx+1}行JSON解析错误: {e}")
                print(f"错误行内容: {line[:100]}...")
df_mrd = pd.DataFrame(data) 

data=[]
with open(r'C:\Users\User\Desktop\pythondata\预后模型评估\10112\bidui\jiegouh\多模型评测\初诊基线表\gemini25flash初诊基线表输出.jsonl', 'r', encoding='utf-8') as f:
    for idx,line in enumerate(f):
        line = line.strip()
        if line:  # 跳过空行
            try:
                json_obj = json.loads(line)
                data.append(json_obj)
            except json.JSONDecodeError as e:
                print(f"第{idx+1}行JSON解析错误: {e}")
                print(f"错误行内容: {line[:100]}...")
df_czjx = pd.DataFrame(data) 

markdown_content =  """
下面是随访疗效评估的术语定义用来学习参考。

随访疗效评估：
Response（缓解）指的是必须同时满足下列所有影像学要求
• Normalisation of bone marrow signal in previously affected areas
• Decrease in the number and size of focal lesions
• Resolution of severely infiltrated bone marrow infiltrate into focal lesions
• Decrease in the of number and size of soft tissue tumors (paramedullary and extramedullary)
Progression（进展）指的是只要出现以下任一项即可判定
• Worsening of diffuse bone marrow signal or new appearance of infiltration in previously unaffected areas
• Increase in the number and size of focal lesions
• Merging of focal lesions into severely infiltrated bone marrow
• Increase in the size or number of soft tissue tumours(paramedullary and extramedullary)
No change（稳定）指的是上述各项均无明显增减


"""

json_obj = []
     
df_excel_suifang = pd.read_excel(r'C:\Users\User\Desktop\pythondata\预后模型评估\10112\bidui\jiegouh\多模型评测\随访评估表\suifang结果.xlsx', sheet_name='Sheet1')
for idx, row in df_excel_suifang.iterrows():
    lineidx = row["对应治疗前索引"]
    lineidx0 = row["PD_索引(从0开始)"]
    qianjghbg = df_jghbg.loc[df_jghbg['custom_id'] == str(lineidx+1) ]["response"]
    benjghbg = df_jghbg.loc[df_jghbg['custom_id'] == str(lineidx0+1) ]["response"]
    benmrd = df_mrd.loc[df_jghbg['custom_id'] == str(lineidx0+1)]["response"]
    qianczjx = df_czjx.loc[df_czjx['custom_id'] == str(lineidx+1)]["response"]
    qianyxbx = df_excel_og.loc[lineidx]['影像表现']
    qianzdjl = df_excel_og.loc[lineidx]['诊断结论']
    benyxbx = row['影像表现']
    
    user_content = f"对应治疗前PET-CT报告如下\n" + f"对应治疗前影像表现：{qianyxbx}\n\n  对应治疗前诊断结论：{qianzdjl}\n\n" + f"AI 辅助生成的治疗前信息如下"+f"对应治疗前初诊基线表{qianczjx}\n 对应治疗前结构化表格{qianjghbg} \n"+ f"治疗后PET-CT报告如下\n" + f"治疗后影像表现：{row['影像表现']}\n\n  治疗后诊断结论：{row['诊断结论']}\n\n" + f"AI 辅助生成的治疗后信息如下" + f"治疗后mrd{benmrd}阴性\n 治疗后结构化表格{benjghbg} \n"+ f"两份多发性骨髓瘤PET-CT自由报告及相关AI辅助信息（仅参考）如上所述" + f"\n我的请求是随访疗效评估最终仅输出 稳定 缓解 进展 三个选项的其中一个"
        # 构建 JSON 对象
    json_obj.append({
        "custom_id": str(idx + 1),
        "method": "POST",
        "url": "/v1/chat/completions",
        "body": {
            "model": "qwen3",
            "messages":[
                {"role": "user", "content": markdown_content + user_content}
            ],"temperature": 0.7, "top_p": 0.9}
    })
        
        # 写入文件（JSON Lines 格式）
with open(r"C:\Users\User\Desktop\pythondata\预后模型评估\10112\bidui\jiegouh\多模型评测\随访评估表\gemini25flash随访输入prompt.jsonl", "w", encoding="utf-8") as f:    
    for item in json_obj:
        f.write(json.dumps(item, ensure_ascii=False) + "\n")

    
    
    