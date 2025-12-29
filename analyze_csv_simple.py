import csv
import chardet

# 尝试检测文件编码
with open('减排行动统计.csv', 'rb') as f:
    result = chardet.detect(f.read())
    encoding = result['encoding']
print(f"检测到文件编码: {encoding}")

# 使用检测到的编码读取文件
print("\n=== 文件内容预览 ===")
try:
    with open('减排行动统计.csv', 'r', encoding=encoding) as f:
        reader = csv.reader(f)
        rows = list(reader)
        print(f"总行数: {len(rows)}")
        print(f"总列数: {len(rows[0]) if rows else 0}")
        print("\n前20行数据:")
        for i, row in enumerate(rows[:20]):
            print(f"行 {i+1}: {row}")
            
    # 分析列名
    if rows:
        print("\n=== 列名分析 ===")
        for i, column_name in enumerate(rows[0]):
            print(f"列 {i+1}: '{column_name}'")
            
except Exception as e:
    print(f"读取文件时出错: {e}")
    print("尝试使用utf-8-sig编码读取...")
    
    # 尝试使用utf-8-sig编码
    try:
        with open('减排行动统计.csv', 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            rows = list(reader)
            print(f"总行数: {len(rows)}")
            print(f"总列数: {len(rows[0]) if rows else 0}")
            print("\n前20行数据:")
            for i, row in enumerate(rows[:20]):
                print(f"行 {i+1}: {row}")
                
    except Exception as e:
        print(f"使用utf-8-sig编码读取也失败了: {e}")
        print("尝试使用gbk编码读取...")
        
        # 尝试使用gbk编码
        try:
            with open('减排行动统计.csv', 'r', encoding='gbk') as f:
                reader = csv.reader(f)
                rows = list(reader)
                print(f"总行数: {len(rows)}")
                print(f"总列数: {len(rows[0]) if rows else 0}")
                print("\n前20行数据:")
                for i, row in enumerate(rows[:20]):
                    print(f"行 {i+1}: {row}")
                    
        except Exception as e:
            print(f"使用gbk编码读取也失败了: {e}")