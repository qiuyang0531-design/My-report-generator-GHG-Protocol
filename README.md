# 温室气体盘查报告生成系统

从 Excel 表单自动生成企业温室气体（GHG）盘查报告，输出符合 GHG Protocol 标准的中文 Word 文档。

## 快速开始

### 环境要求

- Python 3.7+
- Windows（macOS/Linux 可运行但无 COM 目录更新）

### 安装

```bash
pip install -r requirements.txt
```

### 运行

```bash
python main.py --generate
```

默认读取 `blank_form.xlsx`，使用 `blank_template.docx` 模板，输出 `carbon_report.docx`。

指定输入输出：

```bash
python main.py --generate 数据文件.xlsx 输出报告.docx
```

## 工作原理

1. **数据读取** — `data_reader/` 包从 Excel 中自动识别协议类型并提取数据（排放源清单、活动数据、排放因子、GWP 值等）
2. **模板渲染** — 使用 docxtpl（Jinja2）将数据填充到 `blank_template.docx` 模板
3. **后处理** — 表格合并、化学式下标、括号统一、空表清理、目录插入等
4. **COM 目录更新**（Windows + pywin32）— 自动更新 TOC 页码

## 项目结构

```
.
├── main.py                  # 主入口，报告生成流程
├── post_processor.py        # 后处理：TOC、表格合并、化学式、格式清理
├── report_config.py         # 量化方法与类别名称配置
├── jinja2_filters.py        # Jinja2 自定义过滤器
├── inventory_summary_generator.py  # 基准年清单汇总生成
├── data_reader/             # 数据读取包
│   ├── __init__.py          # 公开 API
│   ├── main.py              # ExcelDataReader 主类
│   ├── config.py            # Excel 模板骨架配置
│   ├── extractor.py         # 通用数据提取器
│   ├── fingerprint.py       # Sheet 指纹识别
│   ├── protocols.py         # 协议定义
│   ├── post_processors.py   # 提取后处理
│   └── utils.py             # 工具函数
├── tools/                   # 调试与辅助脚本
├── blank_template.docx      # Word 模板文件
├── blank_form.xlsx          # 空白表单（数据源）
└── requirements.txt
```

## 数据源

使用 Excel（`.xlsx`）格式。`blank_form.xlsx` 为空模板，包含以下 sheet：

- **基本信息** — 公司名称、地址、经营范围等
- **温室气体盘查清册** — 排放源清单（Scope 1/2/3 排放数据）
- **活动数据汇总表** — 燃料、电力等活动数据
- **排放因子表** — 各排放源对应的排放因子
- **GWP 值表** — 全球暖化潜势值
- **减排行动统计** — 减排措施与绩效

## 依赖

| 包 | 用途 |
|---|---|
| `pandas>=2.0.0` | 数据处理 |
| `openpyxl>=3.1.0` | Excel 读取 |
| `python-docx>=0.8.11` | Word 文档生成 |
| `docxtpl` | Jinja2 模板渲染 |
| `pywin32>=305`（Windows 可选） | Word COM 目录页码更新 |

## 许可

仅供学习和研究使用。
