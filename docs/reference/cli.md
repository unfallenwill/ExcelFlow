# CLI 参考

ExcelFlow 安装后提供 `excelflow` 命令，也可以使用：

```bash
python -m excelflow
```

## 命令概览

```text
excelflow [-V]
excelflow template [-o PATH]
excelflow validate -p PATH
excelflow preview -p PATH -t TASK_ID
excelflow run -p PATH -s SOURCE.xlsx [-t TASK_ID] [-f {csv,jsonl,xlsx}] [-o PATH]
```

上面的命令概览使用短选项；每个短选项都可以替换为下文参数表中的对应长选项，例如 `-p` 等价于 `--plan`。

不提供子命令或使用未知参数时，参数解析器显示用法并以非零状态退出。

## template

创建计划模板。

```bash
excelflow template --output extraction_plan.xlsx
```

| 参数 | 必填 | 默认值 | 说明 |
|---|---:|---|---|
| `-o, --output` | 否 | `extraction_plan.xlsx` | 模板输出路径；父目录会自动创建。 |

成功输出：

```text
已生成: extraction_plan.xlsx
```

如果目标文件已存在，当前实现会直接覆盖，不会交互确认。

## validate

读取并校验整个计划文件。

```bash
excelflow validate --plan extraction_plan.xlsx
```

| 参数 | 必填 | 说明 |
|---|---:|---|
| `-p, --plan` | 是 | 计划 Excel 路径。 |

成功输出 `校验通过`，退出码为 `0`。失败时每个问题输出为：

```text
错误: <错误内容>
```

并返回退出码 `1`。该命令不读取源数据 Excel，因此不能发现源 Sheet 缺失、源列名错误或表达式内部字段不存在等运行时问题。

## preview

显示某个任务的 Pandas 执行计划摘要。

```bash
excelflow preview --plan extraction_plan.xlsx --task order_report
```

| 参数 | 必填 | 说明 |
|---|---:|---|
| `-p, --plan` | 是 | 计划 Excel 路径。 |
| `-t, --task` | 是 | “抽取计划”中的任务ID。 |

输出包括读取对象、按顺序排列的关联、过滤条件数量和输出字段数量，例如：

```text
Pandas 执行计划: order_report
读取: 订单 -> o, 订单明细 -> i
关联1: LEFT JOIN i ON o.order_id=i.order_id
过滤条件: 1 条
输出字段: 2 个
```

`preview` 读取计划并查找任务，但当前不会执行完整计划校验，也不会读取源数据或生成输出文件。

## run

执行抽取并写出结果。

```bash
excelflow run --plan PATH --source SOURCE.xlsx [--task TASK_ID] [--format {csv,jsonl,xlsx}] [--output PATH]
```

| 参数 | 必填 | 默认值 | 说明 |
|---|---:|---|---|
| `-p, --plan` | 是 | | 计划 Excel 路径。 |
| `-t, --task` | 否 | | 要执行的任务ID。省略时执行计划中所有“启用”为“是”的任务（多任务模式）。 |
| `-s, --source` | 是 | | 包含源数据 Sheet 的 Excel 文件。 |
| `-f, --format` | 否 | `xlsx` | `csv`、`jsonl` 或 `xlsx`。 |
| `-o, --output` | 否 | 见说明 | **单任务**：输出文件路径（父目录自动创建）；省略时用 `<任务ID>.<格式>` 写入当前目录。**多任务**：输出目录，每个任务生成 `<目录>/<任务ID>.<格式>`；省略时写入当前目录；若该路径指向已存在文件则报错。 |

示例：

```bash
excelflow run -p extraction_plan.xlsx -t order_report -s source.xlsx -f csv -o output/report.csv
excelflow run -p extraction_plan.xlsx -t order_report -s source.xlsx -f jsonl -o output/report.jsonl
excelflow run -p extraction_plan.xlsx -t order_report -s source.xlsx -f xlsx -o output/report.xlsx
# 省略 --format（默认 xlsx）与 --output（默认 order_report.xlsx，写入当前目录）
excelflow run -p extraction_plan.xlsx -t order_report -s source.xlsx
# 多任务：省略 --task，每个启用任务输出 results/<任务ID>.<格式>
excelflow run -p extraction_plan.xlsx -s source.xlsx -f csv --output results
# 多任务，默认写入当前目录
excelflow run -p extraction_plan.xlsx -s source.xlsx
```

输出格式由 `--format` 参数决定，不根据 `--output` 的后缀推断。CSV 使用 UTF-8 BOM 且不写索引；JSONL 每行一个对象、保留非 ASCII 字符并使用 ISO 日期；XLSX 的结果 Sheet 名为 `data`。

执行前会校验整个计划。**单任务模式**（指定 `--task`）确认所选任务的“启用”为“是”，否则报错；**多任务模式**（省略 `--task`）只执行“启用”为“是”的任务、跳过其余任务，任一任务失败即停止（退出码 `1`，已写出的文件保留）。成功输出：

```text
# 单任务
抽取完成: 120 行 -> output/report.csv
# 多任务，每个任务一行
抽取完成 [order_report]: 120 行 -> results/order_report.csv
```

## 版本信息

使用全局选项查看当前安装的 ExcelFlow 版本：

```bash
excelflow --version
excelflow -V
```

显示版本后命令立即退出，不需要指定子命令。

## 退出码与错误输出

| 退出码 | 含义 |
|---:|---|
| `0` | 命令成功。 |
| `1` | 计划校验失败或命令执行过程中发生异常。 |
| `2` | 缺少子命令、缺少必填参数、输出格式不在允许值中或其他 `argparse` 参数解析错误。 |

运行异常写入标准错误，格式为：

```text
失败: <错误内容>
```

计划内容校验失败时，`validate` 将逐条 `错误:` 信息写到标准错误；文件无法读取等异常也会由 `validate` 包装成一条校验错误并写到标准错误。stdout 只用于成功结果和正常预览内容。`preview`、`run` 和 `template` 的异常使用上述 `失败:` 格式写到标准错误。
