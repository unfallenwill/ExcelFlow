# 教程 1：从一个 Sheet 取出需要的列

本课解决最常见的问题：源 Excel 有很多列，但最终只想要订单号、状态和金额。

可运行计划是 [`examples/tutorial/01_single_sheet.xlsx`](https://github.com/unfallenwill/ExcelFlow/blob/main/examples/tutorial/01_single_sheet.xlsx)，源数据是 [`examples/tutorial/source.xlsx`](https://github.com/unfallenwill/ExcelFlow/blob/main/examples/tutorial/source.xlsx)。

## 先看源数据

“订单”Sheet 有以下内容：

| order_id | customer_id | tenant | status | amount |
|---:|---:|---|---|---:|
| 1001 | 10 | A | paid | 727 |
| 1002 | 20 | A | pending | 499 |
| 1003 | 20 | B | paid | 80 |

我们不需要 `customer_id` 和 `tenant`，所以用字段映射只选择其余三列。

## 第一步：声明任务

在“抽取计划”填写：

| 任务ID | 启用 | 备注 |
|---|---|---|
| lesson_01 | 是 | 单 Sheet 和字段映射 |

- 任务ID是命令中选择任务的名称，必须唯一。
- “启用”设为“是”才能执行。暂时不想运行的任务可以设为“否”，而不必删除配置。

## 第二步：告诉 ExcelFlow 数据在哪

在“数据对象”填写：

| 任务ID | Sheet名称 | 对象别名 | 表头行 | 是否主表 | 备注 |
|---|---|---|---:|---|---|
| lesson_01 | 订单 | o | 1 | 是 | 主表 |

为什么这样填：

- `Sheet名称` 必须和源 Excel 底部标签完全一致。
- `o` 是“订单”的短名称，后面用 `o.order_id` 表示订单 Sheet 的 `order_id` 列。
- 表头位于第一行，所以“表头行”填 1。
- 当前只有一个 Sheet，它自然是结果的出发点，所以是主表。

如果源文件前两行是标题和说明、第三行才是列名，表头行就应填 3。

## 第三步：整理输出列

在“字段映射”填写：

| 任务ID | 源字段 | 目标字段 | 目标类型 | 转换表达式 | 字段顺序 |
|---|---|---|---|---|---:|
| lesson_01 | o.order_id | order_id | integer |  | 1 |
| lesson_01 | o.status | status | string |  | 2 |
| lesson_01 | o.amount | amount | decimal |  | 3 |

`源字段` 是原料，`目标字段` 是成品列名。字段顺序决定输出列从左到右的排列。

目标类型用于稳定输出：订单号按整数、状态按文本、金额按小数处理。即使 Excel 看起来显示的是数字，单元格实际也可能是文本。类型转换发生在运行任务时；不能转换的内容会让运行失败，从而暴露需要清理的数据。

本课不需要“关联关系”和“过滤条件”，保持只有表头即可。

## 运行并检查

```bash
uv run excelflow validate --plan examples/tutorial/01_single_sheet.xlsx
uv run excelflow preview --plan examples/tutorial/01_single_sheet.xlsx --task lesson_01
uv run excelflow run \
  --plan examples/tutorial/01_single_sheet.xlsx \
  --task lesson_01 \
  --source examples/tutorial/source.xlsx \
  --format csv \
  --output examples/tutorial/output/01_orders.csv
```

预期得到 3 行、3 列：

```csv
order_id,status,amount
1001,paid,727.0
1002,pending,499.0
1003,paid,80.0
```

若运行时提示 Sheet 不存在，先检查“订单”的文字是否和源 Excel 完全一致；若提示字段不存在，检查列名和 `o.` 前缀。`validate` 只检查计划中的别名关系，不会打开源数据核对真实 Sheet 和列。

下一课：[只保留符合条件的数据](filters.md)。
