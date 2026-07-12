# ExcelFlow

一个由 Excel 声明驱动、使用 Pandas 执行的数据抽取工具。Excel 管理抽取任务、Sheet 关联、过滤和字段映射。

## 快速开始

```bash
uv run excelflow template extraction_plan.xlsx
uv run excelflow validate extraction_plan.xlsx
uv run excelflow preview extraction_plan.xlsx demo_orders
uv run excelflow run extraction_plan.xlsx demo_orders ./data/source.xlsx csv ./output/orders.csv
```

也可以通过模块入口运行：

```bash
uv run python -m excelflow validate extraction_plan.xlsx
```

生成的工作簿包含：

- `抽取计划`：一行一个任务，仅管理任务ID、启用状态和备注。
- `字段映射`：选择、重命名和转换字段；不填写时保留全部字段。
- `过滤条件`：一行一个条件；同组内使用 AND，不同组之间使用 OR。
- `数据对象`：声明任务使用的 Sheet、对象别名和唯一主表。
- `关联关系`：配置 Sheet 之间的 `INNER JOIN` 或 `LEFT JOIN`。
- `填写说明`：字段语义和操作流程。

仅支持 Excel 数据源，输出支持 CSV、JSONL、XLSX。源 Excel 文件在执行 `run` 时传入，计划文件本身不保存数据源路径。一个任务可以声明同一源 Excel 中的多个 Sheet，并通过对象别名关联查询。

源工作表由 Pandas 加载到内存中，通过 `merge`、布尔掩码和 Series 运算完成关联、过滤和字段转换。

过滤运算符支持 `=`、`!=`、`>`、`>=`、`<`、`<=`、`IN`、`NOT IN`、`BETWEEN`、`LIKE`、`NOT LIKE`、`IS NULL` 和 `IS NOT NULL`。

字段映射和过滤条件使用 `对象别名.字段` 格式。每个任务必须且只能有一个主表；相同“关联顺序”的多条关联记录会组合为多个 `AND` 条件。

## 安全边界

转换表达式通过受限 AST 解释器执行，不使用 Python `eval()`。目前只允许字段引用、常量、四则运算、取模以及 `coalesce`、`abs`、`round`、`clip`。

## 项目结构

采用标准 `src/excelflow` 布局。`repository` 读取计划，`validator` 校验声明，`engine` 实现 Pandas 执行策略，`expression` 负责安全衍生列，`output` 提供输出策略，`service` 编排完整流程，`cli` 只处理命令行交互。

## 教学示例

[`examples/README.md`](examples/README.md) 提供从单 Sheet 抽取、条件过滤、Sheet 关联到复合关联键和衍生列的四课递进教程。示例计划和源数据已生成，可以直接运行。
