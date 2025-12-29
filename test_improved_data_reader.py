#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试改进后的 data_reader.py 功能
"""

from data_reader import ExcelDataReader
import os

def test_improved_functions():
    print("=" * 60)
    print("测试改进后的 data_reader.py 功能")
    print("=" * 60)

    # 测试1: CSV 读取功能
    print("\n1. 测试 CSV 读取为列表字典")
    print("-" * 30)
    csv_file = '减排行动统计.csv'
    if os.path.exists(csv_file):
        csv_reader = ExcelDataReader(csv_file)
        emission_actions = csv_reader.read_to_list_of_dicts()

        print(f"[SUCCESS] 读取 {len(emission_actions)} 行减排行动数据")
        print("数据格式:")
        for i, action in enumerate(emission_actions[:3], 1):  # 显示前3行
            print(f"  {i}. {action}")

        # 验证数据结构是否符合要求
        expected_structure = True
        required_keys = ['序号', '减排行动名称', '减排量', '状态']

        for action in emission_actions:
            if not all(key in action for key in required_keys):
                expected_structure = False
                break

        if expected_structure:
            print("[成功] 数据结构正确，包含所有必要的字段")
        else:
            print("[失败] 数据结构不完整，缺少必要字段")
    else:
        print(f"✗ 找不到CSV文件: {csv_file}")

    # 测试2: 增强的 find_value_by_label 功能
    print("\n2. 测试增强的 find_value_by_label 功能")
    print("-" * 30)

    # 查找一个测试Excel文件（如果存在）
    test_files = ['test_data.xlsx', 'carbon_report_v1.docx']  # docx会被跳过

    for filename in test_files:
        if filename.endswith('.xlsx') and os.path.exists(filename):
            print(f"\n测试文件: {filename}")
            reader = ExcelDataReader(filename)

            if reader.workbook:
                print("✓ Excel文件加载成功")

                # 获取所有工作表名称
                sheet_names = reader.workbook.sheetnames
                print(f"  工作表: {sheet_names}")

                # 测试不同的查找参数
                test_cases = [
                    ("组织名称", None, 'right'),  # 在整个工作表中搜索
                    ("组织名称", 'A', 'right'),   # 在A列中搜索
                    ("排放", None, 'below'),      # 搜索包含"排放"的单元格并返回下方值
                ]

                for label, column, direction in test_cases:
                    for sheet_name in sheet_names[:2]:  # 测试前两个工作表
                        try:
                            value = reader.find_value_by_label(
                                sheet_name=sheet_name,
                                label_name=label,
                                column=column,
                                search_direction=direction,
                                exact_match=False,
                                case_sensitive=False
                            )
                            if value:
                                print(f"  ✓ 在 {sheet_name} 中找到 '{label}' (方向: {direction}): {value}")
                        except Exception as e:
                            print(f"  ✗ 查找 '{label}' 时出错: {e}")
            else:
                print("✗ Excel文件加载失败")

    # 测试3: 新增的高级功能
    print("\n3. 测试新增的高级功能")
    print("-" * 30)

    # 查找测试Excel文件
    excel_file = None
    for filename in os.listdir('.'):
        if filename.endswith('.xlsx') and filename != '减排行动统计.csv':
            excel_file = filename
            break

    if excel_file:
        print(f"\n使用测试Excel文件: {excel_file}")
        reader = ExcelDataReader(excel_file)

        if reader.workbook:
            # 测试模式匹配查找
            try:
                patterns = ['总排放量', '范围', '排放量']
                values = reader.find_multiple_values_by_pattern(
                    sheet_name=reader.workbook.sheetnames[0],
                    patterns=patterns,
                    search_direction='right',
                    max_distance=2,
                    require_numeric=True
                )
                print(f"✓ 模式匹配找到 {len(values)} 个数值: {values}")
            except Exception as e:
                print(f"✗ 模式匹配测试失败: {e}")

            # 测试表格数据提取
            try:
                table_data = reader.get_table_data_by_labels(
                    sheet_name=reader.workbook.sheetnames[0],
                    row_labels=['范围一', '范围二', '范围三'],
                    column_labels=['排放量', '总量']
                )
                print(f"✓ 表格数据提取结果: {table_data}")
            except Exception as e:
                print(f"✗ 表格数据提取测试失败: {e}")

    # 测试4: 错误处理能力
    print("\n4. 测试错误处理能力")
    print("-" * 30)

    # 测试不存在的文件
    try:
        non_existent_reader = ExcelDataReader('不存在的文件.xlsx')
        print("✓ 正确处理了不存在的文件")
    except Exception as e:
        print(f"✗ 处理不存在文件时出错: {e}")

    # 测试不支持文件类型
    try:
        unsupported_reader = ExcelDataReader('test.txt')
        print("✓ 正确处理了不支持的文件类型")
    except Exception as e:
        print(f"✗ 处理不支持文件类型时出错: {e}")

    print("\n" + "=" * 60)
    print("测试完成")
    print("=" * 60)

if __name__ == "__main__":
    test_improved_functions()