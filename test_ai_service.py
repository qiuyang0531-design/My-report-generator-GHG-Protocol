#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试改进后的 ai_service.py 功能，确保AI只进行文本润色，不产生数据幻觉
"""

from ai_service import AIService
from data_reader import ExcelDataReader

def test_ai_service_with_real_data():
    """使用真实数据测试AI服务"""
    print("=== 测试AI文本润色服务（使用真实数据） ===\n")

    # 1. 使用真实的数据
    test_data = {
        'company_name': '右侧安全环保有限公司',
        'report_year': '2024',
        'total_emission_location': '7122248.83',
        'scope_1': '1234567.89',
        'scope_2_location': '2345678.90',
        'scope_3': '3542002.04'
    }

    print("输入数据:")
    for key, value in test_data.items():
        print(f"  {key}: {value}")
    print()

    # 2. 测试AI服务
    ai_service = AIService()
    summary = ai_service.generate_executive_summary(test_data)

    print("AI生成的执行摘要:")
    print(f"  {summary}")
    print()

    # 3. 验证摘要内容
    print("=== 摘要内容验证 ===")
    print(f"摘要长度: {len(summary)} 字符")

    # 检查是否包含原始数据中的关键信息
    company_mentioned = test_data['company_name'] in summary
    year_mentioned = test_data['report_year'] in summary
    total_mentioned = '7122248.83' in summary or '7122248' in summary

    print(f"包含公司名称: {company_mentioned}")
    print(f"包含年份: {year_mentioned}")
    print(f"包含总排放量: {total_mentioned}")

    # 检查是否包含禁止内容
    forbidden_patterns = ['##', '**', '```', '预计', '可能', '大约']
    has_forbidden = any(pattern in summary for pattern in forbidden_patterns)
    print(f"包含禁止内容: {has_forbidden}")

def test_data_assembly():
    """测试数据上下文组装功能"""
    print("\n=== 测试数据上下文组装功能 ===\n")

    ai_service = AIService()

    # 测试数据
    test_data = {
        'company_name': '测试企业',
        'report_year': '2024',
        'total_emission_location': '10000',
        'scope_1': '3000',
        'scope_2_location': '4000',
        'scope_3': '3000'
    }

    data_context = ai_service._assemble_data_context(test_data)
    print("组装的数据上下文:")
    print(data_context)
    print()

def test_response_validation():
    """测试AI响应验证功能"""
    print("\n=== 测试AI响应验证功能 ===\n")

    ai_service = AIService()

    # 测试数据
    test_data = {
        'company_name': '测试企业',
        'total_emission_location': '10000',
        'scope_1': '3000',
        'scope_2_location': '4000',
        'scope_3': '3000'
    }

    # 测试正常响应
    valid_response = "测试企业在2024年完成温室气体盘查，总排放量为10000 tCO2e。其中范围一排放3000 tCO2e，范围二排放4000 tCO2e，范围三排放3000 tCO2e。"
    is_valid = ai_service._validate_ai_response(valid_response, test_data)
    print(f"正常响应验证结果: {is_valid}")

    # 测试包含异常数字的响应
    invalid_response = "测试企业在2024年完成温室气体盘查，总排放量为10000 tCO2e。预计明年排放量为15000 tCO2e。"
    is_valid = ai_service._validate_ai_response(invalid_response, test_data)
    print(f"异常数字响应验证结果: {is_valid}")

    # 测试包含Markdown格式的响应
    markdown_response = "## 执行摘要\n测试企业在2024年完成温室气体盘查，总排放量为10000 tCO2e。"
    is_valid = ai_service._validate_ai_response(markdown_response, test_data)
    print(f"Markdown格式响应验证结果: {is_valid}")

def test_fallback_mechanism():
    """测试安全网机制"""
    print("\n=== 测试安全网机制 ===\n")

    # 创建一个没有AI客户端的服务实例
    ai_service = AIService()
    original_client = ai_service.client
    ai_service.client = None  # 模拟AI初始化失败

    test_data = {
        'company_name': '测试企业',
        'report_year': '2024',
        'total_emission_location': '10000',
        'scope_1': '3000',
        'scope_2_location': '4000',
        'scope_3': '3000'
    }

    fallback_summary = ai_service.generate_executive_summary(test_data)
    print("安全网生成的摘要:")
    print(f"  {fallback_summary}")

    # 恢复原始客户端
    ai_service.client = original_client

def main():
    """主测试函数"""
    print("开始测试改进后的AI服务功能\n")

    try:
        # 测试数据组装功能
        test_data_assembly()

        # 测试响应验证功能
        test_response_validation()

        # 测试安全网机制
        test_fallback_mechanism()

        # 测试真实数据（如果有AI配置）
        test_ai_service_with_real_data()

        print("\n=== 功能改进总结 ===")
        print("1. System Prompt 严格限制:")
        print("   - 明确告知AI只能进行文本润色")
        print("   - 严禁编造、预测、估算任何数据")
        print("   - 只能使用提供的数据进行语言优化")

        print("\n2. 上下文组装机制:")
        print("   - _assemble_data_context() 方法严格基于真实数据")
        print("   - 自动计算排放结构比例")
        print("   - 数据完整性验证")

        print("\n3. 防幻觉机制:")
        print("   - _validate_ai_response() 严格验证AI响应")
        print("   - 检查数字一致性")
        print("   - 禁止Markdown格式和推测性词汇")

        print("\n4. AI参数优化:")
        print("   - temperature=0，消除创意性")
        print("   - max_tokens=300，限制输出长度")
        print("   - top_p=1，使用确定性采样")

        print("\n5. 安全网增强:")
        print("   - Fallback机制基于真实数据计算")
        print("   - 自动生成比例分析")

    except Exception as e:
        print(f"测试过程中出现错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()