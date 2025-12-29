#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
测试减排行动统计.csv与系统的集成
验证CSV数据能够被正确读取并用于生成报告
"""

import os
import sys
from datetime import datetime

# 添加当前目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from data_reader import ExcelDataReader
from report_writer import WordReportWriter


def test_csv_reading():
    """测试CSV文件的读取功能"""
    print("=== 测试CSV文件读取功能 ===")
    
    # 测试直接读取CSV文件
    csv_file_path = "D:\\my_report_generator\\减排行动统计.csv"
    
    if not os.path.exists(csv_file_path):
        print(f"错误：CSV文件不存在 - {csv_file_path}")
        return False
    
    try:
        reader = ExcelDataReader(csv_file_path)
        csv_data = reader.extract_data()
        
        print(f"读取CSV文件成功")
        print(f"文件类型：{csv_data.get('file_type', '未知')}")
        
        if 'emission_reductions' in csv_data:
            reductions = csv_data['emission_reductions']
            print(f"读取到 {len(reductions)} 条减排行动记录")
            
            if reductions:
                # 打印前3条记录的字段名和值
                print("\n前3条记录示例：")
                for i, record in enumerate(reductions[:3]):
                    print(f"\n记录 {i+1}：")
                    for key, value in record.items():
                        print(f"  {key}: {value}")
        else:
            print("警告：未找到减排行动数据")
        
        return True
        
    except Exception as e:
        print(f"读取CSV文件时出错：{e}")
        import traceback
        traceback.print_exc()
        return False


def test_integrated_data_extraction():
    """测试从Excel文件中提取所有数据，包括减排行动数据"""
    print("\n=== 测试整合数据提取功能 ===")
    
    # 假设存在一个测试用的Excel文件
    excel_file_path = "D:\\my_report_generator\\温室气体盘查表格.xlsx"
    
    if not os.path.exists(excel_file_path):
        print(f"错误：Excel文件不存在 - {excel_file_path}")
        return False
    
    try:
        reader = ExcelDataReader(excel_file_path)
        all_data = reader.extract_all_data()
        
        print(f"提取所有数据成功")
        
        # 检查温室气体数据
        if 'greenhouse_gas_data' in all_data:
            print("\n温室气体数据：")
            for key, value in all_data['greenhouse_gas_data'].items():
                print(f"  {key}: {value}")
        
        # 检查减排行动数据
        if 'emission_reductions' in all_data:
            reductions = all_data['emission_reductions']
            print(f"\n减排行动数据：共 {len(reductions)} 条记录")
            
            if reductions:
                # 打印前2条记录的字段名和值
                print("\n前2条记录示例：")
                for i, record in enumerate(reductions[:2]):
                    print(f"\n记录 {i+1}：")
                    for key, value in record.items():
                        print(f"  {key}: {value}")
        
        return True
        
    except Exception as e:
        print(f"提取整合数据时出错：{e}")
        import traceback
        traceback.print_exc()
        return False


def test_report_generation():
    """测试报告生成功能，验证减排行动数据能够被正确使用"""
    print("\n=== 测试报告生成功能 ===")
    
    try:
        # 创建测试数据
        test_data = {
            'greenhouse_gas_data': {
                'company_name': '测试公司',
                'report_year': 2024,
                'scope_1': 15000,
                'scope_2_location': 8000,
                'scope_2_market': 7500,
                'scope_3': 3000,
                'total_emission_location': 26000,
                'total_emission_market': 25500
            },
            'emission_reductions': [
                {
                    '序号': 1,
                    'GHG排放类别': '范围一',
                    '排放源': '固定燃烧',
                    '措施': '更换高效燃烧设备'
                },
                {
                    '序号': 2,
                    'GHG排放类别': '范围一',
                    '排放源': '移动燃烧',
                    '措施': '更换为电动车'
                },
                {
                    '序号': 3,
                    'GHG排放类别': '范围二',
                    '排放源': '外购电力',
                    '措施': '采购绿电'
                }
            ],
            'executive_summary': '这是一份测试执行摘要，用于验证减排行动数据的集成功能。'
        }
        
        # 创建报告生成器
        writer = WordReportWriter(
            template_path='D:\\my_report_generator\\模板1.docx',
            cover_image_path='D:\\my_report_generator\\封面.png'
        )
        
        # 生成报告
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = f"D:\\my_report_generator\\test_report_{timestamp}.docx"
        
        success = writer.write_report(test_data, output_path)
        
        if success:
            print(f"报告生成成功：{output_path}")
            return True
        else:
            print("报告生成失败")
            return False
            
    except Exception as e:
        print(f"生成报告时出错：{e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    print("开始测试减排行动统计.csv与系统的集成...\n")
    
    # 运行测试
    csv_test_passed = test_csv_reading()
    data_test_passed = test_integrated_data_extraction()
    report_test_passed = test_report_generation()
    
    # 输出总结
    print("\n=== 测试总结 ===")
    print(f"CSV读取测试：{'通过' if csv_test_passed else '失败'}")
    print(f"整合数据提取测试：{'通过' if data_test_passed else '失败'}")
    print(f"报告生成测试：{'通过' if report_test_passed else '失败'}")
    
    if csv_test_passed and data_test_passed and report_test_passed:
        print("\n🎉 所有测试都已通过！减排行动数据已成功集成到系统中。")
    else:
        print("\n⚠️  部分测试失败，请检查错误信息并修复问题。")
        sys.exit(1)