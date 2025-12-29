import csv

# 尝试使用不同的编码读取文件
encodings_to_try = ['utf-8', 'utf-8-sig', 'gbk', 'gb2312', 'cp936']

for encoding in encodings_to_try:
    print(f"\n尝试使用 {encoding} 编码读取文件:")
    try:
        with open('减排行动统计.csv', 'r', encoding=encoding) as f:
            reader = csv.reader(f)
            rows = list(reader)
            if rows:
                print(f"成功! 总行数: {len(rows)}")
                print(f"总列数: {len(rows[0])}")
                print("\n前10行数据:")
                for i, row in enumerate(rows[:10]):
                    print(f"行 {i+1}: {row}")
                
                # 分析列名
                print("\n列名:")
                for i, column_name in enumerate(rows[0]):
                    print(f"列 {i+1}: '{column_name}'")
                
                # 分析数据行
                print("\n数据行示例:")
                if len(rows) > 1:
                    for i, row in enumerate(rows[1:10]):
                        print(f"行 {i+2}: {row}")
                break
    except Exception as e:
        print(f"失败: {e}")