# CLI 参考

ExcelFlow 安装后提供 `excelflow` 命令，也可以使用：

```bash
python -m excelflow
```

## 命令概览

```text
excelflow template [path]
excelflow validate path
excelflow preview path task_id
excelflow run path task_id source_excel {csv,jsonl,xlsx} output_path
```

不提供子命令或使用未知参数时，参数解析器显示用法并以非零状态退出。

## template

创建计划模板。

```bash
excelflow template [path]
```

| 参数 | 必填 | 默认值 | 说明 |
|---|---:|---|---|
| `path` | 否 | `extraction_plan.xlsx` | 模板输出路径；父目录会自动创建。 |

成功输出：

```text
已生成: extraction_plan.xlsx
```

如果目标文件已存在，当前实现会直接覆盖，不会交互确认。

## validate

读取并校验整个计划文件。

```bash
excelflow validate extraction_plan.xlsx
```

| 参数 | 必填 | 说明 |
|---|---:|---|
| `path` | 是 | 计划 Excel 路径。 |

成功输出 `校验通过`，退出码为 `0`。失败时每个问题输出为：

```text
错误: <错误内容>
```

并返回退出码 `1`。该命令不读取源数据 Excel，因此不能发现源 Sheet 缺失、源列名错误或表达式内部字段不存在等运行时问题。

## preview

显示某个任务的 Pandas 执行计划摘要。

```bash
excelflow preview extraction_plan.xlsx order_report
```

| 参数 | 必填 | 说明 |
|---|---:|---|
| `path` | 是 | 计划 Excel 路径。 |
| `task_id` | 是 | “抽取计划”中的任务ID。 |

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
excelflow run path task_id source_excel output_format output_path
```

| 参数 | 必填 | 说明 |
|---|---:|---|
| `path` | 是 | 计划 Excel 路径。 |
| `task_id` | 是 | 要执行的任务ID。 |
| `source_excel` | 是 | 包含源数据 Sheet 的 Excel 文件。 |
| `output_format` | 是 | `csv`、`jsonl` 或 `xlsx`。 |
| `output_path` | 是 | 输出文件路径；父目录会自动创建。 |

示例：

```bash
excelflow run extraction_plan.xlsx order_report source.xlsx csv output/report.csv
excelflow run extraction_plan.xlsx order_report source.xlsx jsonl output/report.jsonl
excelflow run extraction_plan.xlsx order_report source.xlsx xlsx output/report.xlsx
```

输出格式由 `output_format` 参数决定，不根据 `output_path` 后缀推断。CSV 使用 UTF-8 BOM 且不写索引；JSONL 每行一个对象、保留非 ASCII 字符并使用 ISO 日期；XLSX 的结果 Sheet 名为 `data`。

执行前会校验整个计划，并确认所选任务的“启用”为“是”。成功输出：

```text
抽取完成: 120 行 -> output/report.csv
```

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

计划内容校验失败属于 `validate` 的正常结果，逐条 `错误:` 信息写到标准输出；文件无法读取等异常也会由 `validate` 包装成一条校验错误并写到标准输出。`preview`、`run` 和 `template` 的异常则使用上述 `失败:` 格式写到标准错误。
