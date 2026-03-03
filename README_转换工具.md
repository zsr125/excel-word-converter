# Excel / Numbers 转 Word 产品技术手册工具（分级大纲版）

将 **Excel (.xlsx)** 或 **Apple Numbers (.numbers)** 中的多 sheet（设计、智能、底盘、电池、性能等）转化为一份 **纯文字 + 分级大纲** 的 Word 产品技术手册。  
**不是单纯的文件格式转换**，而是把表格内容重组为「一级/二级大纲 + 正文」的阅读形态，便于在 Word 中编辑与排版。

## 转换规则（三级大纲 + 大纲符 + 格式）

- **一个表格文件** → **一份 Word 文档**（无表格，仅三级大纲 + 纯文字）
- **一级大纲**：每个 sheet 名称，前加大纲符 **一、二、三、…**（标题 1，加粗，大号字，无缩进）
- **二级大纲**：每个 sheet 内每一行（第一列有内容时），前加大纲符 **（一）（二）（三）…**（标题 2，加粗，中号字，左缩进 0.5cm）
- **三级内容**：同一行其余列，每列前加大纲符 **①②③…**，格式为「列名：内容」或纯内容（正文，左缩进 1cm，段后间距统一）
- **第一列为空、其他列有内容**：仅输出为正文段落（视为上一条目的补充），不新增二级标题
- 全文 **宋体**、**缩进与段间距统一**，便于阅读与打印

## 环境要求

- Python 3.10+（使用 numbers-parser 时建议 3.10+）
- 输入格式：`.xlsx`（Excel 2007+）或 `.numbers`（Apple Numbers）

## 安装依赖

```bash
pip install -r requirements_converter.txt
```

或单独安装：

```bash
pip install openpyxl python-docx
```

## 使用方法

### 命令行

```bash
# Excel：指定文件，输出 Word 与输入同目录、同名 .docx
python excel_to_word_converter.py "你的产品点技术信息.xlsx"

# Numbers（Mac 上的 .numbers 文件）
python excel_to_word_converter.py "你的产品点技术信息.numbers"

# 指定输出 Word 路径
python excel_to_word_converter.py "你的产品点技术信息.xlsx" -o "产品技术手册.docx"
python excel_to_word_converter.py "你的产品点技术信息.numbers" -o "产品技术手册.docx"
```

### 在 Python 中调用

```python
from excel_to_word_converter import convert_to_word

# 自动根据扩展名识别 Excel 或 Numbers，输出默认与输入同目录、同名 .docx
out_path = convert_to_word("你的产品点技术信息.xlsx")
out_path = convert_to_word("你的产品点技术信息.numbers")

# 指定输出路径
out_path = convert_to_word("你的产品点技术信息.numbers", "产品技术手册.docx")
```

## 说明

- 输出文档中 **不含任何表格**，全部为 **一级/二级大纲 + 正文段落** 的纯文字结构。
- 首行若为简短文字且无长数字，会识别为表头，正文中以「列名：内容」形式呈现；否则直接按「第一列 = 二级标题、其余列 = 段落」输出。
- **Excel**：合并单元格按合并区域左上角取值；公式按计算后的值读取。
- **Numbers**：每个 sheet 内若有多个表格，会按顺序合并到同一一级大纲下；合并单元格会从合并区域左上角取值。
- 空 sheet：会生成一级标题并标注「（本页无数据）」。
- 同一份清单在 **Windows 用 .xlsx、Mac 用 .numbers** 均可直接转换。

## 文件说明

| 文件 | 说明 |
|------|------|
| `excel_to_word_converter.py` | 转换脚本入口 |
| `requirements_converter.txt` | 依赖列表 |
| `README_转换工具.md` | 本说明 |
