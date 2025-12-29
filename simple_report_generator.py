#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
极简报告生成主程序
串联整个流程：数据读取 -> AI处理 -> 报告生成
"""

import os
from datetime import datetime
from docxtpl import DocxTemplate
from data_reader import ExcelDataReader
from ai_service import AIService

def generate_report(excel_file="test_data.xlsx", csv_file="减排行动统计.csv", template_file="模板1.docx"):
    """
    极简的7步串联流程：
    1. 初始化 DataReader，加载 Excel 文件
    2. 调用提取函数，获得 context 字典
    3. 初始化 AIService，获取 ai_summary
    4. 将 ai_summary 塞回 context 字典
    5. 初始化 DocxTemplate，加载模板
    6. 执行 template.render(context)
    7. 执行 template.save("最终报告.docx")
    """

    # 步骤1: 初始化 DataReader，加载 Excel 文件
    print("=== 步骤1: 初始化 DataReader ===")
    reader = ExcelDataReader(excel_file)

    # 步骤2: 调用提取函数，获得 context 字典
    print("=== 步骤2: 提取数据上下文 ===")
    excel_data = reader.extract_data()
    csv_reader = ExcelDataReader(csv_file)
    emission_actions = csv_reader.read_to_list_of_dicts()

    # 组装完整的 context 字典
    context = {
        # 基本企业信息（来自Excel）
        "company_name": excel_data.get('company_name', '企业名称'),
        "report_year": excel_data.get('report_year', '2024'),

        # 温室气体排放数据（适配模板变量名）
        "scope_1_emissions": excel_data.get('scope_1', 0),
        "scope_2_location-based_emissions": excel_data.get('scope_2_location', 0),
        "scope_2_market-based_emissions": excel_data.get('scope_2_market', 0),
        "scope_3_emissions": excel_data.get('scope_3', 0),

        # 范围3各类别排放量（使用默认值或从scope_3拆分）
        "scope_3_category_1_emissions": excel_data.get('scope_3', 0) * 0.1,  # 示例拆分
        "scope_3_category_2_emissions": excel_data.get('scope_3', 0) * 0.1,
        "scope_3_category_3_emissions": excel_data.get('scope_3', 0) * 0.1,
        "scope_3_category_4_emissions": excel_data.get('scope_3', 0) * 0.1,
        "scope_3_category_5_emissions": excel_data.get('scope_3', 0) * 0.1,
        "scope_3_category_6_emissions": excel_data.get('scope_3', 0) * 0.1,
        "scope_3_category_7_emissions": excel_data.get('scope_3', 0) * 0.1,
        "scope_3_category_9_emissions": excel_data.get('scope_3', 0) * 0.1,
        "scope_3_category_10_emissions": excel_data.get('scope_3', 0) * 0.05,
        "scope_3_category_12_emissions": excel_data.get('scope_3', 0) * 0.05,

        # 企业基本信息（使用默认值）
        "Unified_Social_Credit_Identifier": "91420100MA4L0XX123",
        "leagal_person": "企业法人",
        "registered_capital": "注册资本",
        "date_of_establishment": "成立日期",
        "registered_address": "注册地址",
        "production_address": excel_data.get('company_name', '企业地址'),
        "scope_of_business": "经营范围",
        "company_profile": "企业简介",

        # 报告相关变量
        "reporting_period": f"{excel_data.get('report_year', '2024')}年1月1日至{excel_data.get('report_year', '2024')}年12月31日",
        "document_number": "文档编号",
        "posted_time": datetime.now().strftime("%Y年%m月%d日"),
        "deadline": "核查截止日期",
        "rule_file": "相关规则文件",
        "GWP_Value_Reference_Document": "IPCC第六次评估报告",
        "evaluation_score": "评价等级",

        # 列表变量（来自CSV）
        "emission_actions": emission_actions,

        # 其他有用信息
        "generation_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "emission_actions_count": len(emission_actions)
    }

    print(f"[SUCCESS] 数据提取完成 - 企业: {context['company_name']}, 减排行动: {context['emission_actions_count']}项")

    # 步骤3: 初始化 AIService，获取 ai_summary
    print("=== 步骤3: 生成AI摘要 ===")
    ai_service = AIService()
    ai_summary = ai_service.generate_executive_summary(excel_data)

    # 步骤4: 将 ai_summary 塞回 context 字典
    print("=== 步骤4: 整合AI摘要到上下文 ===")
    context["executive_summary"] = ai_summary

    # 步骤5: 初始化 DocxTemplate，加载模板
    print("=== 步骤5: 加载报告模板 ===")
    template = DocxTemplate(template_file)

    # 步骤6: 执行 template.render(context)
    print("=== 步骤6: 渲染报告 ===")
    template.render(context)

    # 步骤7: 执行 template.save("最终报告.docx")
    print("=== 步骤7: 保存最终报告 ===")
    output_filename = f"碳盘查报告_{context['company_name']}_{context['report_year']}.docx"
    template.save(output_filename)

    print(f"[SUCCESS] 报告生成成功！输出文件: {output_filename}")
    return output_filename, context

if __name__ == "__main__":
    try:
        output_file, final_context = generate_report()
        print(f"\n=== 最终报告信息 ===")
        print(f"企业名称: {final_context['company_name']}")
        print(f"报告年份: {final_context['report_year']}")
        print(f"减排行动数量: {final_context['emission_actions_count']}")
        print(f"AI摘要长度: {len(final_context['executive_summary'])} 字符")
        print(f"生成时间: {final_context['generation_time']}")
    except Exception as e:
        print(f"[ERROR] 程序执行出错: {e}")
        import traceback
        traceback.print_exc()