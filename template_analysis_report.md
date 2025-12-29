# Template.docx 分析报告

## 检查结果概览

### 1. 静态内容标签

#### ✅ 已找到的标签
- `{{ reporting_period }}` - 报告期间
- `{{ GWP_Value_Reference_Document }}` - GWP值参考文档
- `{{ posted_time }}` - 提交时间
- `{{ evaluation_score }}` - 评价分数

#### ❌ 缺少的核心标签
- `{{ company_name }}` - 公司名称
- `{{ report_year }}` - 报告年份
- `{{ total_emission }}` - 总排放量

### 2. 动态表格标签

#### ❌ 完全缺失的标签
- `{% tr for item in items %}` - 表格行循环开始标签
- `{% endtr %}` - 表格行循环结束标签
- `{{ item.name }}`、`{{ item.emission }}`、`{{ item.note }}` - 表格项目标签

## 详细检查结果

### 静态标签分布
| 段落位置 | 标签内容 |
|---------|---------|
| 第258段 | `{{ reporting_period }}` |
| 第264段 | `{{ reporting_period }}` |
| 第268段 | `{{ GWP_Value_Reference_Document }}` |
| 第269段 | `{{ GWP_Value_Reference_Document }}` |
| 第288段 | `{{ posted_time }}`、`{{ evaluation_score }}` |
| 第291段 | `{{ reporting_period }}` |
| 第297段 | `{{ reporting_period }}` |

### 动态表格检查
未在任何表格中发现循环标签。

## 修改建议

### 1. 添加缺失的静态标签

在模板中找到以下位置并替换：
- 公司名称位置 → 替换为 `{{ company_name }}`
- 报告年份位置 → 替换为 `{{ report_year }}`
- 总排放量位置 → 替换为 `{{ total_emission }}`

### 2. 实现动态表格

找到报告中的"减排行动列表"或"主要排放源列表"表格，进行以下修改：

1. **保留表头** - 不要修改表头行
2. **保留一行数据行** - 在表头下方只保留一行数据行
3. **添加循环开始标签** - 在该行的第一个单元格输入 `{% tr for item in items %}`
4. **替换单元格内容** - 将该行各单元格的具体数据替换为：
   - 第一列 → `{{ item.name }}`
   - 第二列 → `{{ item.emission }}`
   - 第三列 → `{{ item.note }}`
5. **添加循环结束标签** - 在该行的最后一个单元格或行尾输入 `{% endtr %}`

### 3. 注意事项
- 所有标签必须直接在Word中输入，不要复制粘贴可能包含隐藏格式的文本
- 可以为标签设置字体样式（加粗、颜色等），生成的报告会继承这些样式
- 确保循环标签 `{% tr for %}` 和 `{% endtr %}` 配对使用
- 表格中只保留一行数据行用于循环，不要保留多余的空行

## 后续步骤

1. 根据上述建议修改template.docx
2. 再次运行 `python check_template.py` 验证修改结果
3. 确保所有要求的标签都已正确添加

完成这些修改后，模板将满足"所见即所得"的样式管理要求，并能支持动态表格生成。