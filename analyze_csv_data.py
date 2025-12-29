#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析减排行动统计.csv的内容，并对比模板需求
"""

import pandas as pd
from data_reader import ExcelDataReader

def analyze_csv_content():
    """分析CSV文件内容"""
    print("=== 减排行动统计.csv 内容分析 ===\n")

    try:
        # 读取CSV文件
        df = pd.read_csv('减排行动统计.csv', encoding='utf-8')
    except:
        try:
            # 尝试GBK编码
            df = pd.read_csv('减排行动统计.csv', encoding='gbk')
        except:
            # 使用ExcelDataReader
            reader = ExcelDataReader('减排行动统计.csv')
            data = reader.read_to_list_of_dicts()
            print(f"使用ExcelDataReader读取到 {len(data)} 条记录")

            if data:
                print(f"CSV文件的列结构: {list(data[0].keys())}")
                print("\n前5条记录:")
                for i, record in enumerate(data[:5], 1):
                    print(f"{i}. {record}")
            return data

    print(f"CSV文件形状: {df.shape}")
    print(f"列名: {list(df.columns)}")
    print("\n前5行数据:")
    print(df.head())

    # 统计信息
    print("\n数据统计:")
    for col in df.columns:
        print(f"- {col}: {df[col].count()} 个非空值")

    # 检查是否有排放量数据
    if '排放量' in df.columns or 'GHG排放量' in df.columns:
        print("\n发现排放量数据列!")

    return df

def compare_with_template_requirements():
    """对比CSV数据与模板需求"""
    print("\n=== CSV数据与模板需求对比 ===\n")

    # 模板需要的变量（根据之前的分析）
    template_variables = [
        # 企业基本信息
        "company_name", "report_year", "Unified_Social_Credit_Identifier",
        "leagal_person", "registered_capital", "date_of_establishment",
        "registered_address", "production_address", "scope_of_business",
        "company_profile",

        # 排放数据
        "scope_1_emissions", "scope_2_location_based_emissions",
        "scope_2_market_based_emissions", "scope_3_emissions",
        "scope_3_category_1_emissions", "scope_3_category_2_emissions",
        "scope_3_category_3_emissions", "scope_3_category_4_emissions",
        "scope_3_category_5_emissions", "scope_3_category_6_emissions",
        "scope_3_category_7_emissions", "scope_3_category_9_emissions",
        "scope_3_category_10_emissions", "scope_3_category_12_emissions",

        # 报告信息
        "reporting_period", "document_number", "posted_time",
        "deadline", "rule_file", "GWP_Value_Reference_Document",
        "evaluation_score"
    ]

    # CSV文件能提供的数据
    csv_provides = [
        "减排行动列表（行动编号、GHG排放源类别、排放源、措施）",
        "总共99条减排行动记录",
        "覆盖各类排放源：固定燃烧源、移动燃烧源、过程排放、逸散排放、移除排放"
    ]

    print("模板需要的变量（29个）:")
    for var in template_variables:
        print(f"- {var}")

    print(f"\nCSV文件可以提供的数据:")
    for data in csv_provides:
        print(f"- {data}")

    print(f"\n结论:")
    print(f"- CSV文件主要提供减排行动的列表数据")
    print(f"- 缺少企业基本信息、排放量数值、报告元数据等")
    print(f"- 需要从其他来源（如test_data.xlsx）获取数值型数据")

def identify_missing_data():
    """识别缺失的数据"""
    print("\n=== 缺失数据识别 ===\n")

    # 检查test_data.xlsx
    try:
        reader = ExcelDataReader('test_data.xlsx')
        excel_data = reader.extract_data()

        print("test_data.xlsx 提供的数据:")
        for key, value in excel_data.items():
            print(f"- {key}: {value}")

        print("\n可以映射到模板的变量:")
        mapping = {
            'company_name': excel_data.get('company_name', '竞进矿产品有限公司'),
            'report_year': excel_data.get('report_year', '2024'),
            'scope_1_emissions': excel_data.get('scope_1', 0),
            'scope_2_location_based_emissions': excel_data.get('scope_2_location', 0),
            'scope_2_market_based_emissions': excel_data.get('scope_2_market', 0),
            'scope_3_emissions': excel_data.get('scope_3', 0),
        }

        for key, value in mapping.items():
            print(f"- {key}: {value}")

    except Exception as e:
        print(f"读取test_data.xlsx失败: {e}")

    print("\n仍需补充的企业基本信息:")
    missing_company_info = [
        "Unified_Social_Credit_Identifier: 统一社会信用代码",
        "leagal_person: 法人代表",
        "registered_capital: 注册资本",
        "date_of_establishment: 成立日期",
        "registered_address: 注册地址",
        "production_address: 生产地址",
        "scope_of_business: 经营范围",
        "company_profile: 企业简介",
    ]

    for info in missing_company_info:
        print(f"- {info}")

    print("\n仍需补充的报告信息:")
    missing_report_info = [
        "document_number: 文档编号",
        "deadline: 核查截止日期",
        "rule_file: 依据的规则文件",
        "GWP_Value_Reference_Document: GWP值参考文档",
        "evaluation_score: 评价等级",
    ]

    for info in missing_report_info:
        print(f"- {info}")

if __name__ == "__main__":
    analyze_csv_content()
    compare_with_template_requirements()
    identify_missing_data()