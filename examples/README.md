# ExcelFlow 入门教程

这套示例按照从简单到复杂的顺序，帮助用户理解 ExcelFlow 的每一张配置表。所有命令都在项目根目录执行。

如果是第一次使用，建议先阅读[核心概念](../docs/getting-started/core-concepts.md)；想理解每一项为什么这样填写，请阅读[模板填写完整教程](../docs/tutorials/template-tutorial.md)。

先生成教学文件：

```bash
uv run python examples/generate.py
```

源数据统一保存在 `examples/tutorial/source.xlsx`，其中包含“订单”“客户”“订单明细”三个 Sheet。

## 第 1 课：抽取单个 Sheet

计划文件：`tutorial/01_single_sheet.xlsx`

这一课只需要关注：

- “抽取计划”：启用任务。
- “数据对象”：将“订单”Sheet 声明为主表，别名为 `o`。
- “字段映射”：选择字段并修改输出列名。

```bash
uv run excelflow validate --plan examples/tutorial/01_single_sheet.xlsx
uv run excelflow preview \
  --plan examples/tutorial/01_single_sheet.xlsx \
  --task lesson_01
uv run excelflow run \
  --plan examples/tutorial/01_single_sheet.xlsx \
  --task lesson_01 \
  --source examples/tutorial/source.xlsx \
  --format csv \
  --output examples/tutorial/output/01_orders.csv
```

## 第 2 课：添加过滤条件

计划文件：`tutorial/02_filters.xlsx`

在第一课基础上增加“过滤条件”：

- `o.status = paid`
- `o.amount >= 100`
- 两条记录属于同一条件组，因此使用 AND 连接。

```bash
uv run excelflow run \
  --plan examples/tutorial/02_filters.xlsx \
  --task lesson_02 \
  --source examples/tutorial/source.xlsx \
  --format jsonl \
  --output examples/tutorial/output/02_paid_orders.jsonl
```

如果条件放在不同的“条件组”，组之间会使用 OR 连接。

## 第 3 课：关联两个 Sheet

计划文件：`tutorial/03_join.xlsx`

这一课增加“客户”数据对象，并配置：

```text
订单 o LEFT JOIN 客户 c
ON o.customer_id = c.customer_id
```

字段映射可以使用 `c.customer_name` 输出客户名称。

```bash
uv run excelflow run \
  --plan examples/tutorial/03_join.xlsx \
  --task lesson_03 \
  --source examples/tutorial/source.xlsx \
  --format xlsx \
  --output examples/tutorial/output/03_orders_with_customer.xlsx
```

## 第 4 课：多 Sheet、复合关联键和衍生列

计划文件：`tutorial/04_derived_columns.xlsx`

最后一课加入“订单明细”Sheet，并使用两个字段关联：

```text
o.order_id = i.order_id
AND o.tenant = i.tenant
```

相同“关联顺序”的两条记录会组成复合关联键。字段映射中的衍生列表达式：

```text
coalesce(i.quantity, 0) * coalesce(i.unit_price, 0)
```

用于计算每一行的 `line_amount`。

```bash
uv run excelflow run \
  --plan examples/tutorial/04_derived_columns.xlsx \
  --task lesson_04 \
  --source examples/tutorial/source.xlsx \
  --format csv \
  --output examples/tutorial/output/04_order_lines.csv
```

预期结果：

```csv
order_id,customer,product,quantity,unit_price,line_amount
1001,张三,键盘,2,299.0,598.0
1001,张三,鼠标,1,129.0,129.0
```

## 第 5 课：分组聚合

计划文件：`tutorial/05_aggregation.xlsx`

这一课按客户分组，统计订单数、订单总金额，并用 `|` 合并状态：

```bash
uv run excelflow run \
  --plan examples/tutorial/05_aggregation.xlsx \
  --task lesson_05 \
  --source examples/tutorial/source.xlsx \
  --format csv \
  --output examples/tutorial/output/05_customer_summary.csv
```

预期结果：

```csv
customer,order_count,total_amount,statuses
张三,1,727.0,paid
李四,2,579.0,pending|paid
```

## 推荐学习方式

依次打开五个计划文件，对比各工作表中新增的配置。每修改一次计划，先运行 `validate`，再运行 `preview` 检查执行计划，最后用 `run` 执行。

`order_report/` 保留了一个独立的综合案例，可用于复制后改造成自己的任务。
