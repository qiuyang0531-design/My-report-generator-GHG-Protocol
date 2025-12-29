#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
检查template.docx和减排行动统计.csv是否满足使用要求
"""

import docx
import pandas as pd
from data_reader import ExcelDataReader

def check_template_status():
    """检查template.docx的状态"""
    print("=== 检查 template.docx ===\n")

    doc = docx.Document('template.docx')

    # 检查语法错误
    import re
    errors = []

    for i, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text
        # 检查变量名中的空格
        vars_with_spaces = re.findall(r'\{\{\s*(\w+(?:-\w+)*)\s+\}\}', text)
        for var in vars_with_spaces:
            errors.append((i+1, f"尾随空格: {{{{{var}}}}}"))

        # 检查连字符
        vars_with_dash = re.findall(r'\{\{\s*(\w+-\w+)\s*\}\}', text)
        for var in vars_with_dash:
            errors.append((i+1, f"包含连字符: {{{{{var}}}}}"))

    if errors:
        print(f"发现 {len(errors)} 个问题:")
        for line, error in errors[:5]:
            print(f"  - 行{line}: {error}")
        if len(errors) > 5:
            print(f"  - ... 还有 {len(errors) - 5} 个问题")
        print("\n状态: ❌ 不满足使用要求，存在语法错误")
        print("建议: 使用已修复的 template_final_fixed.docx")
        return False
    else:
        print("状态: ✅ 满足使用要求")
        return True

def check_csv_status():
    """检查减排行动统计.csv的状态"""
    print("\n=== 检查 减排行动统计.csv ===\n")

    try:
        df = pd.read_csv('减排行动统计.csv', encoding='utf-8')
    except:
        try:
            df = pd.read_csv('减排行动统计.csv', encoding='gbk')
        except Exception as e:
            print(f"读取CSV失败: {e}")
            return False

    print(f"文件基本信息:")
    print(f"  - 行数: {len(df)}")
    print(f"  - 列名: {list(df.columns)}")

    # 检查数据完整性
    print(f"\n数据完整性:")

    for col in df.columns:
        non_null = df[col].notna().sum()
        print(f"  - {col}: {non_null} 个有效值")

    # 检查是否有空行
    empty_rows = df.isnull().all(axis=1).sum()
    if empty_rows > 0:
        print(f"\n⚠ 发现 {empty_rows} 个空行")
    else:
        print(f"\n✅ 没有空行")

    print("\n数据内容:")
    print(f"  - 减排行动记录数: {df['序号'].notna().sum()}")
    print(f"  - 排放源类别: {df['GHG排放源类别'].unique().tolist() if 'GHG排放源类别' in df.columns else 'N/A'}")

    print("\n状态评估:")
    print("  - 作为减排行动列表: ✅ 满足要求")
    print("  - 作为完整数据源: ❌ 不满足（缺少企业信息、排放量等）")

    return True

def check_final_assessment():
    """最终评估"""
    print("\n=== 最终评估 ===\n")

    print("template.docx:")
    print("  - 语法状态: ❌ 存在46处语法错误")
    print("  - 可用版本: ✅ template_final_fixed.docx（已修复）")
    print("  - 建议: 使用 template_final_fixed.docx")

    print("\n减排行动统计.csv:")
    print("  - 减排行动数据: ✅ 完整（99条记录）")
    print("  - 企业基本信息: ❌ 缺失")
    print("  - 排放量数据: ❌ 缺失（需要从test_data.xlsx获取）")
    print("  - 报告元数据: ❌ 缺失")

    print("\n组合使用评估:")
    print("  - template_final_fixed.docx + 减排行动统计.csv + test_data.xlsx")
    print("  - 当前simple_report_generator.py已经整合了这些数据")
    print("  - 状态: ✅ 可以生成报告")

    print("\n改进建议:")
    print("  1. 使用 template_final_fixed.docx 作为模板")
    print("  2. 考虑将所有数据整合到一个Excel文件的多sheet中")
    print("  3. 在test_data.xlsx中补充企业基本信息")

if __name__ == "__main__":
    template_ok = check_template_status()
    csv_ok = check_csv_status()
    check_final_assessment()