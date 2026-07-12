# 五分钟快速开始

这一节使用仓库中已经准备好的教学文件，让你先看到结果，再学习每一项配置。

如果尚未安装，请先阅读[安装 ExcelFlow](installation.md)。

## 1. 准备教学文件

克隆 ExcelFlow 仓库后，在项目根目录运行：

```bash
uv sync
uv run python examples/generate.py
```

你将使用两个文件：

```text
examples/tutorial/01_single_sheet.xlsx  计划文件
examples/tutorial/source.xlsx           源数据文件
```

计划文件声明“从订单 Sheet 取出订单号、状态和金额”；源数据文件保存真实订单。

## 2. 校验计划

```bash
uv run excelflow validate examples/tutorial/01_single_sheet.xlsx
```

成功时显示：

```text
校验通过
```

校验的意义是先检查计划内部的填写规则，例如是否有且只有一个主表、字段引用的对象别名是否已声明。它不会读取或修改源数据。源 Excel 中是否真的存在相应 Sheet 和列，要到运行任务时才能确认。

## 3. 预览任务

```bash
uv run excelflow preview examples/tutorial/01_single_sheet.xlsx lesson_01
```

`lesson_01` 是“抽取计划”中的任务ID。预览让你在执行前确认要读取哪个 Sheet、输出多少列。

本例会显示：

```text
Pandas 执行计划: lesson_01
读取: 订单 -> o
过滤条件: 0 条
输出字段: 3 个
```

## 4. 执行抽取

```bash
uv run excelflow run examples/tutorial/01_single_sheet.xlsx lesson_01 \
  examples/tutorial/source.xlsx csv examples/tutorial/output/01_orders.csv
```

成功时显示：

```text
抽取完成: 3 行 -> examples/tutorial/output/01_orders.csv
```

输出内容是：

```csv
order_id,status,amount
1001,paid,727.0
1002,pending,499.0
1003,paid,80.0
```

## 5. 看懂 run 命令的五个参数

```text
excelflow run <计划文件> <任务ID> <源数据文件> <输出格式> <输出路径>
```

| 参数 | 本例 | 为什么需要 |
|---|---|---|
| 计划文件 | `01_single_sheet.xlsx` | 告诉工具怎样加工数据 |
| 任务ID | `lesson_01` | 一个计划里可能有多个任务，需要选一个 |
| 源数据文件 | `source.xlsx` | 真正要处理的数据 |
| 输出格式 | `csv` | 可以选择 `csv`、`jsonl` 或 `xlsx` |
| 输出路径 | `01_orders.csv` | 告诉工具把结果放在哪里 |

输出格式由命令参数决定，不是由文件后缀猜测。为了避免误解，建议让后缀与格式一致。

## 6. 创建自己的模板

```bash
excelflow template extraction_plan.xlsx
```

打开新文件，将示例内容替换成自己的任务。第一次修改时，建议只做单 Sheet 抽取，验证成功后再增加过滤和关联。

下一步阅读[模板填写完整教程](../tutorials/template-tutorial.md)，或按顺序学习[单 Sheet](../tutorials/single-sheet.md)、[过滤条件](../tutorials/filters.md)、[Sheet 关联](../tutorials/joins.md)和[衍生列](../tutorials/derived-columns.md)。
