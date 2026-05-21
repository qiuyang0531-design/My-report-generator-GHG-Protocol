"""
工具函数模块
=====================

通用工具函数，可在整个包中使用。
"""

import re
from datetime import datetime
from typing import Any, Dict


def excel_date_to_string(date_value):
    """将Excel日期序列号转换为 'YYYY年MM月DD日' 格式"""
    if date_value is None:
        return None
    if isinstance(date_value, str):
        return date_value
    try:
        # Excel日期基准是1899-12-30
        delta = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(date_value) - 2)
        return f'{delta.year}年{delta.month}月{delta.day}日'
    except (ValueError, TypeError):
        return str(date_value)


def safe_str(value) -> str:
    """安全地转换为字符串"""
    if value is None:
        return ''
    return str(value).strip()


def safe_float(value) -> float:
    """安全地转换为浮点数"""
    try:
        if value is None:
            return 0.0
        return float(value)
    except (ValueError, TypeError):
        return 0.0


def safe_get_cell(row, col_idx):
    """安全获取单元格值"""
    try:
        if col_idx < len(row):
            cell = row[col_idx]
            return cell.value if cell.value is not None else ''
        return ''
    except:
        return ''


def normalize_chinese_dates(text: str) -> str:
    """去除中文日期格式中月/日前导零：2007年03月09日 → 2007年3月9日"""
    if not text:
        return text
    text = re.sub(r'(\d{4})年0(\d)月', r'\1年\2月', text)
    text = re.sub(r'月0(\d)日', r'月\1日', text)
    return text


def clean_multiline_text(value: str) -> str:
    """清理多行文本，保留段落分隔（\\n\\n）"""
    if value is None:
        return ''
    if isinstance(value, str):
        # 规范化换行符
        value = value.replace('\r\n', '\n').replace('\r', '\n')
        # 临时保护段落分隔
        value = value.replace('\n\n', '\x00')
        # 段落内换行变空格
        value = re.sub(r'\n+', ' ', value)
        # 恢复段落分隔
        value = value.replace('\x00', '\n\n')
        # 清理多余空格和制表符（保留段落分隔符）
        value = re.sub(r'[ \t]+', ' ', value)
        # 清理段落分隔符前后空格
        value = re.sub(r' ?\n\n ?', '\n\n', value).strip()
        # 统一中英文括号
        value = value.replace('(', '（').replace(')', '）')
        # 去除中文日期前导零
        value = normalize_chinese_dates(value)
    return value


__all__ = [
    'excel_date_to_string',
    'safe_str',
    'safe_float',
    'safe_get_cell',
    'clean_multiline_text',
]
