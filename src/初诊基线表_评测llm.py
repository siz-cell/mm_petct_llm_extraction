import asyncio
import json
import openai
from openai import AsyncOpenAI
from typing import List, Dict, Any
from tqdm import tqdm

class OpenAIBatchProcessor:
    def __init__(self, api_key,base_url, max_concurrent=12):
        self.client = AsyncOpenAI(api_key=api_key,base_url=base_url)
        self.semaphore = asyncio.Semaphore(max_concurrent)
    
    async def process_single(self, request_data: Dict[str, Any], pbar: tqdm) -> Dict[str, Any]:
        async with self.semaphore:
            try:
                # 提取消息内容
                messages = request_data["body"]["messages"]
                
                response = await self.client.chat.completions.create(
                    model="qwen3-235b-a22b-thinking-2507",
                    messages=messages,
                    temperature=0,
                    top_p=0.9
                )
                
 # 如果有思考过程，可以从response中提取
                thinking = ""  # 这里根据实际API响应结构调整
                if hasattr(response, 'thinking') and response.thinking:
                    thinking = response.thinking
                elif hasattr( response.choices[0].message, 'reasoning_content'):
                    thinking =  response.choices[0].message.reasoning_content
                # 更新进度条
                pbar.update(1)
                pbar.set_description(f"处理完成 {request_data['custom_id']}")
                return {
                    "custom_id": request_data["custom_id"],
                    "response": response.choices[0].message.content,
                    "thinking": thinking,  # 保存思考过程
                    "success": True
                }
            
            except Exception as e:
                pbar.update(1)
                print(f"Error processing {request_data.get('custom_id', 'unknown')}: {e}")
                return {
                    "custom_id": request_data.get("custom_id", "unknown"),
                    "response": None,
                    "thinking": "",
                    "success": False,
                    "error": str(e)
                }
    
    async def process_batch_from_jsonl(self, jsonl_file_path: str) -> List[Dict[str, Any]]:
        # 读取JSONL文件
        requests_data = []
        with open(jsonl_file_path, 'r', encoding='utf-8') as f:
            for line_num, line in enumerate(f, 1):
                line = line.strip()
                if line:
                    try:
                        data = json.loads(line)
                        requests_data.append(data)
                    except json.JSONDecodeError as e:
                        print(f"JSON解析错误 第{line_num}行: {e}")
        print(f"总共读取到 {len(requests_data)} 个请求")
        # 创建进度条
        with tqdm(total=len(requests_data), desc="批量处理进度", unit="req") as pbar:
            # 创建处理任务，传入进度条对象
            tasks = [self.process_single(req_data, pbar) for req_data in requests_data]
            
            # 执行所有任务并保持顺序
            results = await asyncio.gather(*tasks)
        
        # 按custom_id排序以确保顺序对应
        results.sort(key=lambda x: int(x["custom_id"]) if x["custom_id"].isdigit() else x["custom_id"])
        
        return results

# 使用示例  qwen3-235b-a22b-thinking-2507
async def main():
    processor = OpenAIBatchProcessor(
        api_key="sk-62bd45813821482cb501ad6890103dd1",
        base_url="https://dashscope.aliyuncs.com/compatible-mode/v1"
        )
#

# async def main():
#     processor = OpenAIBatchProcessor(
#         api_key="sk-x8J6walf5QYrHccaqdC0k0ZA3dHsRyQDAbTtc8r9FIMNaFiP",
#         base_url="https://api.bianxie.ai/v1"
#         )
# 从JSONL文件处理批量请求
    results = await processor.process_batch_from_jsonl(r"C:\Users\User\Desktop\pythondata\预后模型评估\10112\bidui\jiegouh\多模型评测\结构化表格\模型改进评测\初诊基线表\qwen3235b8shotwucot360tiao初诊基线表输入prompt.jsonl")

# 输出顺序一致的JSONL文件，包含思考过程
    output_file = r"C:\Users\User\Desktop\pythondata\预后模型评估\10112\bidui\jiegouh\多模型评测\结构化表格\模型改进评测\初诊基线表\qwen3235b8shotwucot360tiao初诊基线表输出.jsonl"
    with open(output_file, 'w', encoding='utf-8') as f:
        for result in results:
            output_data = {
                "custom_id": result['custom_id'],
                "success": result['success'],
                "thinking": result.get('thinking', ''),  # 保存思考过程
                "response": result['response'] if result['success'] else None,
                "error": result.get('error', '') if not result['success'] else ''
            }
            f.write(json.dumps(output_data, ensure_ascii=False) + "\n")

print(r"结果已保存到: {output_file}")
# 运行程序
if __name__ == "__main__":
    asyncio.run(main())