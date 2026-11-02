# SheetPorter · 表格搬运工

[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](#许可证)

把 Excel 源表按**列名**搬进 Excel 模板：自动换算单位、自动加行、保护表头表尾。

适合每月重复的「从系统导出明细 → 填进固定格式表格」工作：采购计划单、出入库单、对账单、工资表、报表等。

> A small desktop tool that copies Excel rows into a fixed-format template by column name, with automatic unit conversion, row insertion, and header/footer protection. Pure local, no network.


---

## 特性

- **按列名匹配**：源表和目标表都按表头名称找列，列顺序变了也不怕
- **自动加行**：源数据比模板数据区多几行就插几行，复制上一行的边框、字体、对齐、行高
- **保护表头表尾**：标题、表头、备注、签名行不动，只在数据区操作
- **单位转换**：`1.0000袋/盒` → `盒`（取斜杠后），`1.0000支/支` → `支`
- **多套规则**：一套规则 = 一张模板 + 若干任务，可新建、复制、导入、删除
- **中文图形界面**：现代卡片式布局，状态胶囊 + 日志着色，不写代码也能配规则
- **纯本地**：只读写本地文件，不联网、不上传任何数据
- **单文件 exe**：Windows 双击即用，不需要装 Python

---

## 快速开始

### 方式一：直接用 exe

1. 从 [Releases](https://github.com/suyongle98/SheetPorter/releases) 下载并解压 `SheetPorter_v1.1.1.zip`
2. 双击 `SheetPorter.exe`
3. 顶部规则栏默认有「示例：月度采购计划单」，直接点 **开始生成**
4. 打开 `输出/` 文件夹看结果

示例会演示：18 行源数据 → 12 行模板自动插 6 行、单位从「粒/盒」变成「盒」、备注和签名行整体下移。

### 方式二：从源码运行

```bash
pip install openpyxl
python sheetporter.py
```

需要 Python 3.9+。

---

## 它怎么工作

```
源 Excel ──按列名取数──▶ 值转换（取斜杠后等）──▶ 模板数据区
                                                      │
                        数据比模板行数多 ──自动插行、复制格式──▶ 输出 Excel
                                                      │
                                              表头/表尾保持不动
```

一套 **规则**（rule）包含：

| 字段 | 说明 |
|---|---|
| `name` | 规则名称 |
| `template` | 模板文件路径 |
| `output_dir` | 输出文件夹 |
| `output_name` | 输出文件名，支持 `{年}{月}{日}` 等占位符 |
| `tasks` | 任务列表 |

一个 **任务**（task）= 一张源表填一个 sheet：

| 字段 | 说明 |
|---|---|
| `source.dir` / `source.pattern` | 源文件所在文件夹 + 文件名通配符（如 `医院*.xlsx`） |
| `source.sheet` | 源表第几个 sheet（0 = 第一个），也可写 sheet 名 |
| `source.header_row` | 源表表头在第几行 |
| `source.data_start` | 源表数据从第几行开始 |
| `target.sheet` | 目标 sheet 名 |
| `target.header_row` | 目标表头在第几行 |
| `target.data_start` | 目标数据从第几行开始 |
| `cells` | 可选，固定单元格内容，如 `{"A1": "XX 单位 {年} 年 {月} 月计划表"}` |
| `columns` | 列映射：`{"target": "目标列名", "source": "源列名", "transform": "转换"}` |

### 路径与占位符

- 相对路径相对于程序所在目录；也可以写绝对路径。
- 文件名、`cells` 里可以用：`{年} {月} {日} {yy} {mm} {dd} {yyyy} {MM} {DD}`
  例：`采购计划_{yy}年{mm}月{dd}.xlsx` → `采购计划_26年11月02.xlsx`

### 支持的转换

| transform | 说明 |
|---|---|
| `direct` | 原样写入 |
| `after_slash` | 取 `/` 后面：`1.0000袋/盒` → `盒` |
| `before_slash` | 取 `/` 前面 |
| `strip` | 去掉首尾空格 |
| `fixed` | 固定值（配合 `arg` 字段） |

---

## 规则示例

```json
{
  "name": "示例：月度采购计划单",
  "template": "examples/template.xlsx",
  "output_dir": "输出",
  "output_name": "示例输出_{yy}年{mm}月{dd}.xlsx",
  "tasks": [
    {
      "name": "示例表",
      "source": {
        "dir": "examples",
        "pattern": "source*.xlsx",
        "sheet": 0,
        "header_row": 1,
        "data_start": 2
      },
      "target": { "sheet": "示例表", "header_row": 2, "data_start": 3 },
      "cells": { "A1": "示例医院 {年} 年 {月} 月药品采购计划表" },
      "columns": [
        { "target": "药品名称", "source": "药品名称", "transform": "direct" },
        { "target": "药品编码", "source": "医保代码", "transform": "direct" },
        { "target": "单位", "source": "包装信息", "transform": "after_slash" }
      ]
    }
  ]
}
```

完整示例见 `examples/demo_rule.json`。

---

## 目录结构

```
SheetPorter/
├─ SheetPorter.exe      # 打包版（可选）
├─ sheetporter.py       # 图形界面入口
├─ fillcore.py          # 填表引擎
├─ examples/            # 示例模板、源表、规则
├─ rules/               # 你的规则（json，一文件一套）
├─ 输出/                # 默认输出目录
├─ README.md
├─ 使用说明.md
├─ LICENSE
├─ requirements.txt
├─ build.bat            # 自己打包 exe
└─ .gitignore
```

---

## 自己打包 exe

Windows 下双击 `build.bat`，或手动：

```bash
python -m venv .venv
.venv\Scripts\activate
pip install pyinstaller openpyxl
pyinstaller --noconfirm --clean --onefile --windowed --name SheetPorter --add-data "examples;examples" sheetporter.py
```

产物在 `dist/SheetPorter.exe`。

---

## 常见问题

**源表列名改了怎么办？**
工具按列名找列，找不到会提示「源表里找不到这些列」并列出实际表头。改规则里的 `source` 即可。

**模板数据行不够怎么办？**
自动插入，不用手工拖。

**生成的文件在哪？**
规则里的 `output_dir`。

**能处理多个 sheet 吗？**
一个任务对应一个 sheet，加多个任务即可（示例规则只有一个任务）。

**会不会改坏模板？**
不会。工具先把模板复制一份，再在副本上填。

**数据安全？**
纯本地处理，没有任何网络请求。

---

## 许可证

MIT License，见 [LICENSE](LICENSE)。

## 适用范围与验证

支持 `.xlsx`。数据区按第一列连续边框判断；请用空白模板，表尾前保留一行无边框空行。复杂公式、图表、宏、跨行合并和打印区域在插行后可能需要人工调整。源表公式读取 Excel 已保存的计算结果，程序不计算公式。多个文件匹配时按文件名排序取最后一个，忽略 `~$` 临时文件。

`direct` 保留源值类型（数字保持数字）；源表中的文字始终按文字写入。同名旧输出只在新结果生成成功后备份，生成失败不替换旧输出。示例中的机构、企业、地址、编码和业务记录均为演示数据。

验证命令：`python -m unittest discover -s tests -v`。
