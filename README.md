# pdcheck

一个由 Excel 声明驱动的数据抽取工具。Excel 管理数据源、抽取任务和字段映射；程序负责校验、SQL 预览和执行。

## 快速开始

```bash
uv run python main.py template extraction_plan.xlsx
uv run python main.py validate extraction_plan.xlsx
uv run python main.py preview extraction_plan.xlsx demo_orders
uv run python main.py run extraction_plan.xlsx demo_orders
```

生成的工作簿包含：

- `抽取计划`：一行一个任务，定义全量/增量范围、过滤条件及输出。
- `数据源`：连接信息；密码只保存环境变量名，不能保存明文。
- `字段映射`：选择、重命名和转换字段；不填写时使用 `SELECT *`。
- `填写说明`：字段语义和操作流程。

增量区间采用左闭右开 `[开始值, 结束值)`，便于相邻批次无缝衔接。开始值和结束值可以写成 `${START_TIME}` 这样的环境变量占位符。

支持 SQLite、Excel、JSON 三类数据源，以及 CSV、JSONL、XLSX 输出：

- SQLite：`连接地址`填写数据库文件，`源对象`填写表或视图。
- Excel：`连接地址`填写工作簿，`源对象`填写工作表；扩展参数可用 `{"header_row": 1}` 指定表头行。
- JSON：文件内容可以直接是对象数组，也可以是包含对象数组的根对象；后一种情况由`源对象`指定数组键。扩展参数可用 `{"encoding": "utf-8"}` 指定编码。

Excel 和 JSON 会加载到内存临时表，然后复用 SQLite 的筛选、字段转换和增量区间语义。JSON 中的嵌套对象或数组会以 JSON 字符串写入输出。

## 安全边界

`源对象`和普通字段名会进行标识符校验。`过滤条件`及`转换表达式`属于受信任配置，会作为 SQL 片段原样执行，因此 Excel 的编辑权限应只授予可信人员，并在运行前使用 `preview` 审核 SQL。
