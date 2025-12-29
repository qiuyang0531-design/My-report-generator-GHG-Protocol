#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
数据补充方案和建议
"""

def create_data_supplementation_plan():
    """创建数据补充计划"""
    print("=== 使用现有文件生成报告的完整方案 ===\n")

    print("1. 现有数据源分析:")
    print("   a) test_data.xlsx:")
    print("      - ✓ company_name: 竞进矿产品有限公司")
    print("      - ✓ report_year: 2024")
    print("      - ✓ scope_1: 7122248.83 (范围一排放)")
    print("      - ✓ scope_3: 4538211.34 (范围三排放)")
    print("      - ⚠ scope_2_location/market: '环保设备' (文本，需要数值)")

    print("\n   b) 减排行动统计.csv:")
    print("      - ✓ 99条减排行动记录")
    print("      - ✓ 包含行动编号、GHG排放源类别、排放源、措施")
    print("      - 覆盖5大类排放源：")
    print("        - 固定燃烧源（21项）")
    print("        - 移动燃烧源（11项）")
    print("        - 过程排放（16项）")
    print("        - 逸散排放（20项）")
    print("        - 移除排放（31项）")

    print("\n2. 需要补充的数据:")

    print("\n   a) 企业基本信息（8项）:")
    company_info = [
        ("Unified_Social_Credit_Identifier", "统一社会信用代码", "建议格式：91420100MA4L0XX123"),
        ("leagal_person", "法人代表", "示例：张三"),
        ("registered_capital", "注册资本", "示例：5000万元"),
        ("date_of_establishment", "成立日期", "示例：2010-05-20"),
        ("registered_address", "注册地址", "示例：湖北省大冶市XX路XX号"),
        ("production_address", "生产地址", "示例：湖北省大冶市XX工业园区"),
        ("scope_of_business", "经营范围", "示例：矿产品加工、销售等"),
        ("company_profile", "企业简介", "建议50-100字描述")
    ]
    for var, desc, example in company_info:
        print(f"      - {var}: {desc} [{example}]")

    print("\n   b) 排放数据补充:")
    print("      - scope_2_location_based_emissions: 需要实际数值（当前是'环保设备'）")
    print("      - scope_2_market_based_emissions: 需要实际数值（当前是'环保设备'）")
    print("      - scope_3_category_X_emissions: 可以从scope_3按比例分配")

    print("\n   c) 报告元数据（5项）:")
    report_meta = [
        ("document_number", "文档编号", "格式：TC-2024-001"),
        ("deadline", "核查截止日期", "建议：2024年12月31日"),
        ("rule_file", "规则文件", "建议：GB/T 32150-2015"),
        ("GWP_Value_Reference_Document", "GWP参考文档", "建议：IPCC第六次评估报告"),
        ("evaluation_score", "评价等级", "建议：A级或优")
    ]
    for var, desc, example in report_meta:
        print(f"      - {var}: {desc} [{example}]")

    print("\n3. 数据补充建议方案:")
    print("\n   方案一：创建补充数据文件（推荐）")
    print("   - 创建 company_info.csv 文件存储企业基本信息")
    print("   - 修改 test_data.xlsx 修复scope_2数据")
    print("   - 在 .env 文件中设置默认值用于报告元数据")

    print("\n   方案二：代码中设置默认值（快速）")
    print("   - 在 simple_report_generator.py 中直接设置默认值")
    print("   - 适合快速测试和演示")

    print("\n   方案三：手动输入（灵活）")
    print("   - 程序运行时提示用户输入")
    print("   - 适合多企业使用场景")

    print("\n4. 范围3排放量分配建议:")
    scope3_split = {
        "scope_3_category_1_emissions": 0.20,  # 原材料采购
        "scope_3_category_2_emissions": 0.10,  # 资本货物
        "scope_3_category_3_emissions": 0.05,  # 其他燃料
        "scope_3_category_4_emissions": 0.15,  # 上游运输
        "scope_3_category_5_emissions": 0.10,  # 经营废料
        "scope_3_category_6_emissions": 0.05,  # 商务旅行
        "scope_3_category_7_emissions": 0.05,  # 员工通勤
        "scope_3_category_9_emissions": 0.15,  # 下游运输
        "scope_3_category_10_emissions": 0.05,  # 加工服务
        "scope_3_category_12_emissions": 0.10   # 废弃物处理
    }

    print("   基于 scope_3: 4,538,211.34 tCO2e，建议分配：")
    for category, ratio in scope3_split.items():
        value = 4538211.34 * ratio
        print(f"   - {category}: {value:,.2f} tCO2e ({ratio*100:.0f}%)")

    return company_info, report_meta, scope3_split

def generate_implementation_guide():
    """生成实施指南"""
    print("\n=== 实施指南 ===\n")

    print("步骤1：修复现有数据")
    print("   a) 在 test_data.xlsx 中，将 scope_2_location/market 的值改为数值")
    print("      - 例如：scope_2_location = 1234567.89")
    print("      - 例如：scope_2_market = 1234567.89")

    print("\n步骤2：创建企业信息文件")
    print("   创建 company_info.csv 文件，内容如下：")
    print("""
   变量名,值,说明
   Unified_Social_Credit_Identifier,91420100MA4L0XX123,统一社会信用代码
   leagal_person,张三,法人代表
   registered_capital,5000万元,注册资本
   date_of_establishment,2010-05-20,成立日期
   registered_address,湖北省大冶市XX路XX号,注册地址
   production_address,湖北省大冶市XX工业园区,生产地址
   scope_of_business,矿产品加工、销售,经营范围
   company_profile,专业从事矿产品加工的高新技术企业...,企业简介
   """)

    print("\n步骤3：更新报告生成代码")
    print("   修改 simple_report_generator.py，在context字典中添加所有缺失变量")

    print("\n步骤4：测试和验证")
    print("   a) 运行 simple_report_generator.py")
    print("   b) 检查生成的报告是否包含所有数据")
    print("   c) 验证数据格式和计算结果")

if __name__ == "__main__":
    company_info, report_meta, scope3_split = create_data_supplementation_plan()
    generate_implementation_guide()