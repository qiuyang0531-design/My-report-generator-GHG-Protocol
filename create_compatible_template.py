#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
创建一个与现有数据兼容的模板
基于模板1.docx，但使用template.docx的样式和结构
"""

import docx
from docxtpl import DocxTemplate
import re

def create_compatible_template():
    """创建与数据兼容的模板"""
    # 加载模板1.docx作为基础（这个是工作的）
    base_template = DocxTemplate('模板1.docx')

    # 创建新文档
    new_doc = docx.Document()

    # 复制模板1.docx的所有内容到新文档
    for element in base_template.doc.element.body:
        new_doc.element.body.append(element)

    # 添加一些额外的变量，使其更接近template.docx的内容
    # 在文档末尾添加一些额外的段落
    additional_info = [
        "统一社会信用代码：{{ Unified_Social_Credit_Identifier }}",
        "法人代表：{{ leagal_person }}",
        "注册资本：{{ registered_capital }}",
        "成立日期：{{ date_of_establishment }}",
        "注册地址：{{ registered_address }}",
        "生产地址：{{ production_address }}",
        "经营范围：{{ scope_of_business }}",
        "企业简介：{{ company_profile }}",
        "文档编号：{{ document_number }}",
        "提交时间：{{ posted_time }}",
        "核查截止日期：{{ deadline }}",
        "规则文件：{{ rule_file }}",
        "GWP参考文档：{{ GWP_Value_Reference_Document }}",
        "评价等级：{{ evaluation_score }}"
    ]

    for info in additional_info:
        new_doc.add_paragraph(info)

    # 保存为新的模板文件
    new_doc.save('template_compatible.docx')
    print("创建了兼容的模板文件：template_compatible.docx")

if __name__ == "__main__":
    create_compatible_template()