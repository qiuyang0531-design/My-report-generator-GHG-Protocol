#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
碳盘查报告生成器 - 极简主程序
"""

from data_reader import ExcelDataReader
from ai_service import AIService
from docxtpl import DocxTemplate


def generate_report(csv_path="减排行动统计.csv", output_path="carbon_report_v1.docx"):
    """
    一键生成碳盘查报告

    Args:
        csv_path: CSV数据文件路径（包含所有32个模板变量）
        output_path: 输出报告路径
    """
    # 1. 初始化 DataReader，加载 CSV 文件
    reader = ExcelDataReader(csv_path)

    # 2. 提取 context 字典（包含所有32个模板变量）
    context = reader.extract_data()

    # 3. 初始化 AIService，获取 ai_summary
    ai = AIService()
    ai_summary = ai.generate_executive_summary(context)

    # 4. 将 ai_summary 塞回 context
    context["executive_summary"] = ai_summary

    # 5. 初始化 DocxTemplate，加载模板
    template = DocxTemplate("template.docx")

    # 6. 渲染模板
    template.render(context)

    # 7. 保存报告
    template.save(output_path)
    print(f"报告已生成: {output_path}")

    return output_path


if __name__ == "__main__":
    generate_report()