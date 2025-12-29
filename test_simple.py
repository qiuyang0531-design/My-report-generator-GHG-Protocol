#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简单测试改进后的 data_reader.py 功能
"""

from data_reader import ExcelDataReader

def test_csv_functionality():
    """测试CSV读取为列表字典功能"""
    print("=== 测试CSV读取为列表字典功能 ===")

    csv_reader = ExcelDataReader('减排行动统计.csv')
    emission_actions = csv_reader.read_to_list_of_dicts()

    print(f"成功读取 {len(emission_actions)} 行减排行动数据")
    print("数据格式:")
    for i, action in enumerate(emission_actions, 1):
        print(f"  {i}. {action}")

    return emission_actions

def test_excel_functionality():
    """测试Excel查找功能"""
    print("\n=== 测试Excel查找功能 ===")

    try:
        reader = ExcelDataReader('test_data.xlsx')

        if reader.workbook:
            print("Excel文件加载成功")
            print(f"工作表: {reader.workbook.sheetnames}")

            # 测试增强的find_value_by_label功能
            first_sheet = reader.workbook.sheetnames[0]
            print(f"\n在 '{first_sheet}' 工作表中测试查找功能:")

            # 测试1: 在整个工作表中搜索
            test_cases = [
                ("组织", None, 'right'),
                ("排放", None, 'below'),
                ("范围", None, 'right')
            ]

            for label, column, direction in test_cases:
                try:
                    value = reader.find_value_by_label(
                        sheet_name=first_sheet,
                        label_name=label,
                        column=column,
                        search_direction=direction,
                        exact_match=False,
                        case_sensitive=False
                    )
                    print(f"  查找 '{label}' (方向: {direction}): {value}")
                except Exception as e:
                    print(f"  查找 '{label}' 时出错: {e}")

        else:
            print("Excel文件加载失败")
    except Exception as e:
        print(f"测试Excel功能时出错: {e}")

def main():
    """主测试函数"""
    print("开始测试改进后的 data_reader.py\n")

    # 测试CSV功能
    emission_data = test_csv_functionality()

    # 测试Excel查找功能
    test_excel_functionality()

    print("\n=== 功能改进总结 ===")
    print("1. find_value_by_label函数增强:")
    print("   - 支持整个工作表搜索或指定列搜索")
    print("   - 支持多个搜索方向 (right, left, below, above)")
    print("   - 支持精确匹配或包含匹配")
    print("   - 支持大小写敏感/不敏感")
    print("   - 摒弃硬坐标定位，更健壮")

    print("\n2. read_to_list_of_dicts函数增强:")
    print("   - 支持多种编码格式")
    print("   - 自动清理表头和数据")
    print("   - 智能处理空行和空值")
    print("   - 支持灵活的行范围配置")

    print("\n3. 新增高级功能:")
    print("   - find_multiple_values_by_pattern: 模式匹配查找")
    print("   - get_table_data_by_labels: 基于标签提取表格数据")
    print("   - _clean_cell_value: 数据清理和标准化")

if __name__ == "__main__":
    main()