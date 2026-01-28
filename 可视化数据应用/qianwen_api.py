# -*- coding: utf-8 -*-

from openai import OpenAI

class QwenClient:
    def __init__(self, api_key, base_url, model):
        self.api_key = api_key
        self.base_url = base_url
        self.model = model
        self.client = OpenAI(
            api_key=self.api_key,
            base_url=self.base_url
        )
    
    def generate(self, messages, temperature=0.7, max_tokens=2000):
        """生成文本"""
        try:
            completion = self.client.chat.completions.create(
                model=self.model,
                messages=messages,
                temperature=temperature,
                max_tokens=max_tokens
            )
            return completion.choices[0].message.content
        except Exception as e:
            return f"生成失败: {str(e)}"


import json
import numpy as np

def prepare_person_data(current_name, sheet_names, sheet_dfs):
    """
    准备人员数据用于评价生成
    
    Args:
        current_name: 当前选中的人员姓名
        sheet_names: 子表名称列表
        sheet_dfs: 子表数据字典
    
    Returns:
        格式化的人员数据字符串
    """
    person_data = {}
    
    for sheet_name in sheet_names:
        df = sheet_dfs[sheet_name]
        person_row = df[df.iloc[:, 0] == current_name]
        
        if not person_row.empty:
            numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
            if numeric_cols:
                person_data[sheet_name] = {}
                for col in numeric_cols:
                    person_data[sheet_name][col] = float(person_row.iloc[0][col])
    
    return json.dumps(person_data, ensure_ascii=False, indent=2)


def generate_evaluation(person_data, person_name, api_key, base_url, model):
    """
    生成人员评价
    
    Args:
        person_data: 人员数据JSON字符串
        person_name: 人员姓名
        api_key: API密钥
        base_url: API基础URL
        model: 模型名称
    
    Returns:
        生成的评价文本
    """
    try:
        # 初始化客户端
        client = QwenClient(api_key=api_key, base_url=base_url, model=model)
        
        # 构建提示词
        prompt = f"""请根据以下数据，为{person_name}生成一份专业的评价报告。

数据如下：
{person_data}

请从以下几个方面进行评价：
1. 整体表现概述
2. 各项指标分析
3. 优势与亮点
4. 改进建议

要求：
- 评价要客观、准确，简洁
- 语言要专业但不晦涩
- 数据中数值代表问题数量，数值越小问题越少，数值越大问题越多
- 字数控制在150字左右

"""
        
        # 调用通义千问模型
        messages = [
            {
                'role': 'system',
                'content': '你是一位专业的数据分析师，擅长根据数据生成客观、准确的人员评价。'
            },
            {
                'role': 'user',
                'content': prompt
            }
        ]
    #  调用客户端的生成方法，生成文本响应     
    #  参数说明：- messages: 输入的消息列表，用于生成上下文   
    #   - temperature: 控制生成文本的随机性，值越大随机性越高，当前设为0.7      
    #   - max_tokens: 生成文本的最大长度，当前设为200个token       
        return client.generate(messages, temperature=0.7, max_tokens=200) 

        
    except Exception as e:
        return f"生成评价时出错: {str(e)}"
