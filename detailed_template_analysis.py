#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
详细分析template.docx中的所有问题
"""

import docx
import re

def analyze_all_issues():
    """分析所有需要修改的问题"""
    doc = docx.Document('template.docx')

    print("=== template.docx 详细问题分析 ===\n")

    # 1. 语法错误
    syntax_errors = []
    all_tags = []
    undefined_variables = []
    potential_issues = []

    # 收集所有标签
    for i, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text
        tags = re.findall(r'(\{\{[^}]+\}\}|\{\%[^%]+\%\})', text)

        for tag in tags:
            all_tags.append((i+1, tag.strip()))
            clean_tag = tag.strip()

            # 语法错误检查
            if '{{' in clean_tag:
                var_content = clean_tag.replace('{{', '').replace('}}', '').strip()

                # 检查空格
                if ' ' in var_content and not any(op in var_content for op in ['|', 'filter', 'if', 'for']):
                    syntax_errors.append((i+1, f"变量名包含空格: {clean_tag}"))

                # 检查无效字符
                if any(char in var_content for char in ['\t', '\n', '\r']):
                    syntax_errors.append((i+1, f"变量名包含无效字符: {clean_tag}"))

                # 检查Python关键字
                python_keywords = ['and', 'or', 'not', 'in', 'is', 'None', 'True', 'False']
                words = var_content.split()
                if len(words) > 1 and words[0] in python_keywords:
                    syntax_errors.append((i+1, f"可能使用了Python关键字: {clean_tag}"))

    print("1. 语法错误:")
    if syntax_errors:
        for line_num, error in syntax_errors:
            print(f"   行 {line_num}: {error}")
    else:
        print("   未发现语法错误")

    print(f"\n总共发现 {len(syntax_errors)} 个语法错误\n")

    # 2. 变量名问题分析
    print("2. 变量名分析:")
    variables = set()
    for line_num, tag in all_tags:
        if '{{' in tag:
            var_name = tag.replace('{{', '').replace('}}', '').strip().split()[0]
            variables.add(var_name)

    # 检查变量名规范
    naming_issues = []
    for var in sorted(variables):
        # 检查连字符（Python不支持）
        if '-' in var:
            naming_issues.append(f"包含连字符: {var}")

        # 检查特殊字符
        if any(char in var for char in ['@', '#', '$', '%', '^', '&', '*', '(', ')', '+', '=', '[', ']', '{', '}', '|', '\\', ';', ':', '"', "'", '<', '>', ',', '.', '?', '/']):
            naming_issues.append(f"包含特殊字符: {var}")

        # 检查中文字符
        if re.search(r'[\u4e00-\u9fff]', var):
            naming_issues.append(f"包含中文字符: {var}")

    if naming_issues:
        print("   变量命名问题:")
        for issue in naming_issues[:10]:
            print(f"   - {issue}")
        if len(naming_issues) > 10:
            print(f"   - ... 还有 {len(naming_issues) - 10} 个问题")
    else:
        print("   变量命名规范")

    print(f"\n3. 需要修复的具体问题:")

    # 3. 生成修复建议
    fixes_needed = set()

    # 从之前分析得知的问题
    fixes_needed.add('{{ scope_of_ business }} → {{ scope_of_business }}')

    # 检查所有变量，找出可能的问题
    for var in variables:
        if '-' in var:
            # 将连字符替换为下划线
            fixed_var = var.replace('-', '_')
            fixes_needed.add(f'{{{{{var}}}}} → {{{{{fixed_var}}}}}')

    print(f"   总共需要修复 {len(fixes_needed)} 个变量:")
    for fix in sorted(fixes_needed):
        print(f"   - {fix}")

    print(f"\n4. 其他建议:")
    print("   - 确保所有变量在数据准备代码中都有对应的值")
    print("   - 考虑简化模板，移除不必要的变量")
    print("   - 建立变量命名规范（使用下划线分隔，避免特殊字符）")
    print("   - 添加模板验证机制")

    return {
        'syntax_errors': syntax_errors,
        'naming_issues': naming_issues,
        'fixes_needed': fixes_needed,
        'total_variables': len(variables)
    }

if __name__ == "__main__":
    issues = analyze_all_issues()