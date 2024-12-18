import json
import pandas as pd
import re

# 输入和输出文件路径
input_file = "collection.json"  # 导出的 JSON 文件路径
output_file = "api_details.xlsx"        # 输出的 Excel 文件路径

def clean_placeholder(text):
    if isinstance(text, str):
        # 移除占位符
        text = re.sub(r"\{\{.*?\}\}", "", text)
        # 尝试将 JSON 字符串转换为单行
        try:
            json_obj = json.loads(text)
            text = json.dumps(json_obj, ensure_ascii=False)
        except json.JSONDecodeError:

            text = re.sub(r'\s+', ' ', text).strip()
    return text

def parse_postman_collection(json_file):
    # 读取 Postman 导出的 JSON 文件
    with open(json_file, 'r', encoding='utf-8') as f:
        data = json.load(f)

    api_data = []

    # 遍历所有 API 请求项
    def extract_items(items):
        for item in items:
            if 'item' in item:  # 如果有子项，递归提取
                extract_items(item['item'])
            else:
                api_name = item.get('name', '')  # API 名称
                request = item.get('request', {})

                # 获取请求方法
                method = request.get('method', '')

                # 获取 URL 并清理占位符
                url_data = request.get('url', {})
                url = clean_placeholder(url_data.get('raw', ''))

                # 获取 Header 信息
                headers = request.get('header', [])
                header_str = "\n".join([
                    f"{h['key']}: {clean_placeholder(h['value'])}" 
                    for h in headers if 'key' in h and 'value' in h
                ])

                # 获取 Body 数据并清理占位符
                body = request.get('body', {})
                body_content = ''
                if body.get('mode') == 'raw':
                    body_content = clean_placeholder(body.get('raw', ''))

                # 存入数据列表
                api_data.append({
                    "API 名称": api_name,
                    "请求方法": method,
                    "请求 URL": url,
                    "请求头 (Headers)": header_str,
                    "请求体 (Body)": body_content
                })

    # 开始提取数据
    extract_items(data.get('item', []))
    return api_data

def save_to_excel(api_data, output_file):

    df = pd.DataFrame(api_data)
    # 保存到 Excel 文件
    df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"数据已成功保存到 {output_file}")

if __name__ == "__main__":
    api_data = parse_postman_collection(input_file)
    save_to_excel(api_data, output_file)
