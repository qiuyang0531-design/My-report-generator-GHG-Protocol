# ai_service.py
import os
import re
from openai import OpenAI
from dotenv import load_dotenv

# 加载 .env 文件中的秘密
load_dotenv()

class AIService:
    def __init__(self):
        """
        初始化 AI 服务。
        它会从 .env 文件读取配置并准备好 AI 客户端。
        功能严格限制在"文本润色"，确保不产生数据幻觉。
        """
        try:
            self.client = OpenAI(
                api_key=os.getenv("OPENAI_API_KEY"),
                base_url=os.getenv("OPENAI_BASE_URL"),
                timeout=20.0 # 设置20秒超时
            )
            print("AI 文本润色服务初始化成功。")
        except Exception as e:
            print(f"AI 服务初始化失败: {e}")
            self.client = None

    def _get_fallback_summary(self, data):
        """
        这是"安全网"。当 AI 失败时，调用这个函数。
        它返回一个基于真实数据的简单文本，绝不产生幻觉。
        """
        print("警告：AI 文本润色失败，启动 Fallback 安全网。")
        company = data.get('company_name', '该公司')
        year = data.get('report_year', '本年度')
        total = data.get('total_emission_location', '0')
        scope1 = data.get('scope_1', '0')
        scope2 = data.get('scope_2_location', '0')
        scope3 = data.get('scope_3', '0')

        # 计算比例（如果数据有效）
        try:
            total_val = float(str(total).replace('t', '').replace(',', '').strip())
            scope1_val = float(str(scope1).replace('t', '').replace(',', '').strip())
            scope2_val = float(str(scope2).replace('t', '').replace(',', '').strip())

            if total_val > 0:
                scope1_pct = (scope1_val / total_val) * 100
                scope2_pct = (scope2_val / total_val) * 100

                return f"{company}在{year}完成温室气体盘查，总排放量为{total}tCO2e。其中范围一排放{scope1}tCO2e（占比{scope1_pct:.1f}%），范围二排放{scope2}tCO2e（占比{scope2_pct:.1f}%）。企业已识别主要排放源，并将基于此数据制定下一步减排计划。"
            else:
                return f"{company}在{year}完成温室气体盘查，总排放量为{total}tCO2e。企业已识别主要排放源，并将基于此数据制定下一步减排计划。"
        except:
            return f"{company}在{year}完成温室气体盘查，总排放量为{total}tCO2e。企业已识别主要排放源，并将基于此数据制定下一步减排计划。"

    def _assemble_data_context(self, data):
        """
        组装数据上下文，将提取到的关键数据拼接成字符串。
        确保AI只能基于这些真实数据进行文本润色。
        """
        try:
            company = data.get('company_name', '企业')
            year = data.get('report_year', '本年度')
            total_location = data.get('total_emission_location', '0')
            total_market = data.get('total_emission_market', '0')
            scope1 = data.get('scope_1', '0')
            scope2_location = data.get('scope_2_location', '0')
            scope2_market = data.get('scope_2_market', '0')
            scope3 = data.get('scope_3', '0')

            # 清理和标准化数据
            def clean_number(value):
                if value is None:
                    return '0'
                value_str = str(value).replace('t', '').replace(',', '').strip()
                try:
                    return float(value_str)
                except:
                    return '0'

            total_loc_val = clean_number(total_location)
            scope1_val = clean_number(scope1)
            scope2_loc_val = clean_number(scope2_location)
            scope3_val = clean_number(scope3)

            # 计算排放结构比例
            proportions = []
            if total_loc_val > 0:
                if scope1_val > 0:
                    scope1_pct = (scope1_val / total_loc_val) * 100
                    proportions.append(f"范围一占比{scope1_pct:.1f}%")
                if scope2_loc_val > 0:
                    scope2_pct = (scope2_loc_val / total_loc_val) * 100
                    proportions.append(f"范围二占比{scope2_pct:.1f}%")
                if scope3_val > 0:
                    scope3_pct = (scope3_val / total_loc_val) * 100
                    proportions.append(f"范围三占比{scope3_pct:.1f}%")

            # 严格的数据上下文字符串
            data_summary = f"""
企业：{company}
年份：{year}
总排放：{total_location} tCO2e
范围一排放：{scope1} tCO2e
范围二排放（基于位置）：{scope2_location} tCO2e
范围三排放：{scope3} tCO2e
排放结构：{', '.join(proportions) if proportions else '数据待分析'}
"""

            # 验证数据完整性
            if any(val is None or val == '' for val in [company, year, total_location, scope1, scope2_location]):
                print("警告：数据不完整，可能影响摘要质量")

            return data_summary.strip()

        except Exception as e:
            print(f"数据上下文组装失败: {e}")
            return f"企业：{data.get('company_name', '企业')}，年份：{data.get('report_year', '本年度')}，数据组装失败，请检查原始数据。"

    def _validate_ai_response(self, content, original_data):
        """
        验证AI响应，确保没有产生数据幻觉。
        严格检查：只允许文本润色，不允许编造数据。
        """
        try:
            # 提取原始数据中的关键数字
            original_numbers = []
            data_fields = ['total_emission_location', 'scope_1', 'scope_2_location', 'scope_3']
            for field in data_fields:
                value = original_data.get(field)
                if value:
                    # 提取数字部分
                    numbers = re.findall(r'\d+\.?\d*', str(value))
                    original_numbers.extend(numbers)

            # 检查AI响应中的数字是否都与原始数据一致
            ai_numbers = re.findall(r'\d+\.?\d*', content)

            for ai_num in ai_numbers:
                ai_num = ai_num.lstrip('0') if ai_num != '0' else ai_num
                found = False

                # 检查是否与原始数据中的数字匹配
                for orig_num in original_numbers:
                    orig_num = orig_num.lstrip('0') if orig_num != '0' else orig_num
                    if abs(float(ai_num) - float(orig_num)) < 0.01:  # 允许小的浮点误差
                        found = True
                        break

                # 如果不是原始数据中的数字，检查是否是合理的衍生数据
                if not found:
                    try:
                        # 检查是否是年份（合理的年份范围：1900-2100）
                        if 1900 <= float(ai_num) <= 2100:
                            continue

                        # 检查是否是百分比（0-100）
                        if 0 <= float(ai_num) <= 100:
                            # 检查是否以百分比形式出现
                            ai_num_index = content.find(ai_num)
                            context_start = max(0, ai_num_index - 5)
                            context_end = min(len(content), ai_num_index + len(ai_num) + 5)
                            context = content[context_start:context_end]

                            if '%' in context or '占比' in context or '比例' in context:
                                continue

                        # 检查是否是小的整数（可能是序号等）
                        if len(ai_num) <= 2 and float(ai_num) < 50:
                            continue

                        print(f"警告：AI响应中包含未经验证的数字 {ai_num}")
                        return False
                    except (ValueError, TypeError):
                        # 如果转换失败，认为是无效数字
                        print(f"警告：AI响应中包含无效数字格式 {ai_num}")
                        return False

            # 检查是否包含禁止的格式
            forbidden_patterns = ['##', '**', '```', '*', '-', '1.', '2.']
            for pattern in forbidden_patterns:
                if pattern in content:
                    print(f"警告：AI响应包含禁用格式 {pattern}")
                    return False

            # 检查是否包含可能的幻觉内容
            hallucination_keywords = ['预计', '预测', '可能', '大约', '大概', '估计', '推测']
            for keyword in hallucination_keywords:
                if keyword in content:
                    print(f"警告：AI响应包含可能的幻觉关键词 {keyword}")
                    return False

            return True

        except Exception as e:
            print(f"AI响应验证失败: {e}")
            return False

    def generate_executive_summary(self, data):
        """
        这是"总管"调用的唯一方法。
        功能严格限制在"文本润色"，确保不产生数据幻觉。
        """
        if not self.client:
            # 如果初始化都失败了，直接使用安全网
            return self._get_fallback_summary(data)

        # 1. 组装数据上下文 - 严格基于真实数据
        data_context = self._assemble_data_context(data)

        # 2. 非常严格的系统提示词 - 限制AI只能进行文本润色
        SYSTEM_PROMPT = """
你是一个专业的碳核算报告助手。你的唯一任务是对提供的排放数据进行文本润色，严禁编造任何数据。

严格要求：
1. 你只能使用我提供的数据进行文本润色
2. 严禁编造、预测、估算任何数据
3. 严禁添加任何未在数据中出现的信息
4. 语气必须专业、客观
5. 重点描述排放结构（范围一与范围二的比例关系）
6. 篇幅控制在300字以内
7. 只输出纯文本内容，不要包含任何Markdown格式
8. 不要使用"预计"、"可能"、"大约"等推测性词汇
9. 如果数据不完整，如实说明"数据待补充"

你的角色是文本润色，不是数据分析师。只做语言的优化和重组。
"""

        # 3. 用户指令 - 明确要求基于提供的数据
        USER_PROMPT = f"""
请根据以下企业提供的关键排放数据，撰写一段专业的"执行摘要"。你必须严格基于以下数据进行文本润色：

{data_context}

要求：
- 仅使用上述数据进行文本润色
- 描述排放结构和比例关系
- 保持专业客观的语气
- 控制在300字以内
- 输出纯文本格式
"""

        try:
            print("正在调用AI进行文本润色...")
            print(f"数据上下文: {data_context[:100]}...")

            response = self.client.chat.completions.create(
                model="gpt-3.5-turbo",  # 使用稳定可靠的模型
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": USER_PROMPT}
                ],
                temperature=0,  # 温度设为0，消除任何创意性，确保纯文本润色
                max_tokens=300,  # 限制最大输出长度
                top_p=1,        # 使用确定性采样
                frequency_penalty=0,  # 不改变词频
                presence_penalty=0    # 不引入新话题
            )

            # 提取AI的回复
            content = response.choices[0].message.content.strip()

            # 严格的响应验证
            if self._validate_ai_response(content, data):
                print("AI文本润色成功，响应验证通过")
                return content
            else:
                print("AI响应验证失败，启动安全网")
                return self._get_fallback_summary(data)

        except Exception as e:
            print(f"AI文本润色调用失败: {e}")
            # AI失败了，但程序不能失败。我们启动安全网。
            return self._get_fallback_summary(data)