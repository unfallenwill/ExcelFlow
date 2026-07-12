# pdcheck

一个由 Excel 声明驱动的数据抽取工具。Excel 管理数据源、抽取任务和字段映射；程序负责校验、SQL 预览和执行。

## 快速开始

```bash
uv run python main.py template extraction_plan.xlsx
uv run python main.py validate extraction_plan.xlsx
uv run python main.py preview extraction_plan.xlsx demo_orders
uv run python main.py run extraction_plan.xlsx demo_orders ./data/source.xlsx
```

生成的工作簿包含：

- `抽取计划`：一行一个任务，定义全量/增量范围、过滤条件及输出。
- `字段映射`：选择、重命名和转换字段；不填写时使用 `SELECT *`。
- `过滤条件`：一行一个条件；同组内使用 AND，不同组之间使用 OR。
- `填写说明`：字段语义和操作流程。

增量区间采用左闭右开 `[开始值, 结束值)`，便于相邻批次无缝衔接。开始值和结束值可以写成 `${START_TIME}` 这样的环境变量占位符。

仅支持 Excel 数据源，输出支持 CSV、JSONL、XLSX。源 Excel 文件在执行 `run` 时传入，`源对象`填写工作表名称，`表头行`从 1 开始计数。计划文件本身不保存数据源路径。

源工作表会加载到内存中执行筛选、字段转换和增量区间判断。

过滤运算符支持 `=`、`!=`、`>`、`>=`、`<`、`<=`、`IN`、`NOT IN`、`BETWEEN`、`LIKE`、`NOT LIKE`、`IS NULL` 和 `IS NOT NULL`。条件值使用参数化查询，不作为 SQL 片段执行。

## 安全边界

`源对象`和普通字段名会进行标识符校验。`过滤条件`及`转换表达式`属于受信任配置，会作为 SQL 片段原样执行，因此 Excel 的编辑权限应只授予可信人员，并在运行前使用 `preview` 审核 SQL。
