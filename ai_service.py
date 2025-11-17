# ai_service.py 
import os 
from openai import OpenAI 
from dotenv import load_dotenv 

# 加载 .env 文件中的秘密 
load_dotenv() 

class AIService: 
    def __init__(self): 
        """
        初始化 AI 服务。
        它会从 .env 文件读取配置并准备好 AI 客户端。
        """
        try: 
            self.client = OpenAI( 
                api_key=os.getenv("OPENAI_API_KEY"), 
                base_url=os.getenv("OPENAI_BASE_URL"), 
                timeout=20.0 # 设置20秒超时 
            ) 
            print("AI 服务初始化成功。") 
        except Exception as e: 
            print(f"AI 服务初始化失败: {e}") 
            self.client = None 

    def _get_fallback_summary(self, data): 
        """
        这就是"安全网"。当 AI 失败时，调用这个函数。
        它返回一个简单但绝对安全的文本。
        """
        print("警告：AI 生成失败，启动 Fallback 安全网。") 
        company = data.get('company_name', '该公司') 
        year = data.get('report_year', '本年度') 
        total = data.get('total_emission_location', '一个数值') 
        
        # 返回一个格式化的、简单的字符串 
        return f"{company}在{year}完成了碳盘查工作。报告期内，总排放量为 {total} tCO2e。企业已识别主要排放源，并将基于此数据制定下一步减排计划。"

    def generate_executive_summary(self, data): 
        """
        这是"总管"调用的唯一方法。
        它负责：1. 尝试调用 AI。 2. 如果失败，调用安全网。
        """
        if not self.client: 
            # 如果初始化都失败了，直接使用安全网 
            return self._get_fallback_summary(data) 

        # ---------------------------------------------
        # 关键挑战：Prompt 工程 
        # ---------------------------------------------
        
        # 1. 准备给 AI 的"数据" 
        # 我们把作业一的字典数据转换成简单文本 
        data_text = f"""
        公司名称: {data.get('company_name')} 
        报告年份: {data.get('report_year')} 
        总排放(基于位置): {data.get('total_emission_location')} tCO2e 
        范围一排放: {data.get('scope_1')} tCO2e 
        范围二排放(基于位置): {data.get('scope_2_location')} tCO2e 
        范围三排放: {data.get('scope_3')} tCO2e 
        """
        
        # 2. 定义"系统角色" (约束 AI 的行为) 
        SYSTEM_PROMPT = """
        你是一个专业的碳排放报告撰写专家。
        你的任务是根据用户提供的关键数据，撰写一段专业的"执行摘要"。
        你的写作风格必须是：专业、简洁、客观、正式。
        
        你的回复必须严格遵守以下规则： 
        1. 必须是纯文本 (Plain Text)。 
        2. 绝对不允许使用任何 Markdown 格式，例如 ##、**、* 或 - 。 
        3. 绝对不允许编造数据之外的任何信息。 
        4. 内容必须控制在 200 字以内。 
        """
        
        # 3. 定义"用户指令" (给 AI 的具体任务) 
        USER_PROMPT = f"""
        请根据以下数据，为碳盘查报告撰写执行摘要： 
        {data_text} 
        
        请严格按照你被设定的规则输出。 
        """

        try: 
            print("正在尝试调用 AI 生成摘要...") 
            response = self.client.chat.completions.create( 
                model="gpt-3.5-turbo", # 或者你有的其他模型 
                messages=[ 
                    {"role": "system", "content": SYSTEM_PROMPT}, 
                    {"role": "user", "content": USER_PROMPT} 
                ], 
                temperature=0.1, # 温度调低，让它减少"创意" 
            ) 
            
            # 提取 AI 的回复 
            content = response.choices[0].message.content.strip() 
            
            # 最后的"安全检查"，防止 AI 还是不听话 
            if "##" in content or "**" in content: 
                print("警告：AI 仍然返回了 Markdown，已强制清理。") 
                content = content.replace("##", "").replace("**", "") 
            
            print("AI 摘要生成成功。") 
            return content 

        except Exception as e: 
            print(f"AI 调用失败: {e}") 
            # AI 失败了，但程序不能失败。我们启动"B 计划"。 
            return self._get_fallback_summary(data)