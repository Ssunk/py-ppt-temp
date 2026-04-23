# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Response Language

用中文回答，包括思考过程。

## Project Overview

CSV-to-PPTX 报告生成工具。读取多个 CSV 文件，基于 PowerPoint 模板自动生成带封面、分隔页和数据表格的 PPT 报告。

## Commands

```bash
# 虚拟环境: D:\workspace\venv
D:\workspace\venv\Scripts\python.exe <script.py>

# 生成报告 (主入口)
D:\workspace\venv\Scripts\python.exe generate_ppt.py

# 生成模板
D:\workspace\venv\Scripts\python.exe create_template.py
```

## Architecture

两个核心脚本:

1. **create_template.py** -- 构建 `report_template.pptx` 模板文件（含3张幻灯片: 封面/分隔页/数据页）
2. **generate_ppt.py** -- 主程序，加载模板 + CSV 数据，克隆模板幻灯片填入数据后删除原始模板页，输出 `report.pptx`

### 模板结构 (report_template.pptx)

| Slide Index | Name | 用途 | 关键 Shape |
|---|---|---|---|
| 0 | COVER | 封面 | Title, Subtitle, ContentList |
| 1 | SECTION | 分隔页 | SectionTitle, DimensionInfo |
| 2 | DATA_TABLE | 数据页 | SectionTitle, PageInfo, SampleTable |

DATA_TABLE 中的 SampleTable 是示例表格（默认1行表头+10行数据），行数决定了最终报告每页展示的数据行数。`generate_ppt.py` 通过 `_read_table_config()` 从该表格动态读取行数、行高以及单元格样式（字体/字号/颜色/填充）。

### 模板样式同步机制

用户在 PowerPoint 中编辑 `report_template.pptx` 的 SampleTable 样式后，生成报告时会自动同步：

| 读取来源 | 读取内容 | 用途 |
|---|---|---|
| `cell(0,0)` 第0行表头 | 字体/字号/颜色/粗体/对齐 + 填充色 | 报告表头样式 |
| `cell(1,0)` 第1行数据 | 字体/字号/颜色/粗体/对齐 | 报告数据行样式 |
| `cell(2,0)` 第2行数据 | 填充色 | 报告交替行背景色 |

模板未显式设置的属性会回退到代码中的硬编码默认值。幻灯片上新增的 shape（文本框/图片等）通过 `_clone_slide` 深拷贝自动同步；表格列数由 CSV 决定，不支持模板同步。

### generate_ppt.py 核心流程

- `_clone_slide()`: 深拷贝模板幻灯片（可选排除指定 shape，如排除 SampleTable 再创建新表格）
- 对每个 CSV: 生成一张 section 分隔页 + N 张数据页（按 rows_per_page 分页）
- 最后倒序删除3张原始模板幻灯片

### 样式常量

PRISM_BLUE `#0070C0`, DARK_GRAY `#333333`, MEDIUM_GRAY `#666666`, TABLE_HEADER_BG `#0070C0`, TABLE_HEADER_FG `#FFFFFF`, TABLE_ROW_ALT `#E8F0FE`

## Dependencies

python-pptx (依赖 lxml)

## Data Files

employees.csv, products.csv, orders.csv -- 输入数据，首行为表头
