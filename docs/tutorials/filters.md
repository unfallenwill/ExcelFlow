# 教程 2：只保留符合条件的数据

本课在订单中只保留“已经支付，并且金额不少于 100”的记录。

可运行计划是 [`examples/tutorial/02_filters.xlsx`](../../examples/tutorial/02_filters.xlsx)。它沿用[单 Sheet 教程](single-sheet.md)的任务、数据对象和字段映射，只增加“过滤条件”。

## 先把业务语言写清楚

我们的要求是：

```text
状态等于 paid
并且
金额大于等于 100
```

关键词是“并且”：一条订单必须同时通过两次检查。因此两条条件放在同一个条件组。

## 填写过滤条件

| 任务ID | 条件组 | 条件序号 | 字段 | 运算符 | 值1 | 值2 | 备注 |
|---|---:|---:|---|---|---|---|---|
| lesson_02 | 1 | 1 | o.status | = | paid |  | 同组条件使用 AND |
| lesson_02 | 1 | 2 | o.amount | >= | 100 |  |  |

ExcelFlow 会把它理解为：

```text
(o.status = paid AND o.amount >= 100)
```

条件序号只负责表达同一组中的先后，建议从 1 开始连续填写，方便阅读。

源数据中：订单 1002 没有支付，订单 1003 金额不足，只有订单 1001 同时符合要求。

## 什么时候使用不同条件组

如果要求变成“已支付的订单，或者金额不少于 100”，任意一个条件成立即可，就把它们放进不同组：

| 条件组 | 条件序号 | 字段 | 运算符 | 值1 |
|---:|---:|---|---|---|
| 1 | 1 | o.status | = | paid |
| 2 | 1 | o.amount | >= | 100 |

ExcelFlow 会理解为：

```text
(o.status = paid) OR (o.amount >= 100)
```

记忆方法：同组是一个必须整体满足的要求，不同组是多套可选方案。

## 运行

```bash
uv run excelflow validate examples/tutorial/02_filters.xlsx
uv run excelflow preview examples/tutorial/02_filters.xlsx lesson_02
uv run excelflow run examples/tutorial/02_filters.xlsx lesson_02 \
  examples/tutorial/source.xlsx jsonl examples/tutorial/output/02_paid_orders.jsonl
```

预期只有订单 1001：

```json
{"order_id":1001,"status":"paid","amount":727.0}
```

## 常见填写错误

- 想表达“并且”却放在不同组，会留下比预期更多的数据。
- 想表达“或者”却放在同一组，会留下比预期更少的数据。
- `IN` 的多个值要用英文逗号分隔，例如 `paid,pending`。
- `BETWEEN` 需要同时填写值1和值2。
- `IS NULL` 和 `IS NOT NULL` 不需要填写值。
- `LIKE` 中 `%` 代表任意长度文字，`_` 代表任意一个字符。

下一课：[把两个 Sheet 的信息放到一起](joins.md)。
