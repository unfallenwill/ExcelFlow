# 教程：分组聚合

逐行衍生回答“每一行算出什么”，分组聚合回答“一组记录合起来是多少”。例如订单明细有三行，按客户汇总后可能只剩一行。

本课使用可直接运行的计划 [`examples/tutorial/05_aggregation.xlsx`](https://github.com/unfallenwill/ExcelFlow/blob/main/examples/tutorial/05_aggregation.xlsx)，任务ID为 `lesson_05`。

## 场景

源文件中的订单关联客户后，要按客户统计订单数、总金额，并把订单状态连接起来。

## 填写分组字段

| 任务ID | 源字段 | 目标字段 | 目标类型 | 分组顺序 |
|---|---|---|---|---:|
| lesson_05 | c.customer_name | customer | string | 1 |

这表示“客户相同的记录放在一组”。多个分组字段会进一步细分结果。

## 填写聚合规则

| 任务ID | 源字段 | 聚合函数 | 目标字段 | 目标类型 | 分隔符 | 聚合顺序 |
|---|---|---|---|---|---|---:|
| lesson_05 |  | count_all | order_count | integer |  | 1 |
| lesson_05 | o.amount | sum | total_amount | decimal |  | 2 |
| lesson_05 | o.status | concat_agg | statuses | string | \| | 3 |

本例在订单粒度上汇总，所以 `count_all` 就是订单数。如果先关联一对多明细，应该对订单号使用 `count_distinct`，避免重复计数。不填写分组字段时，所有记录放在一组，输出一行总计。

## 执行顺序和限制

```text
读取 → 关联 → 过滤 → 分组聚合 → 类型转换 → 输出
```

聚合任务的最终列由“分组字段”和“聚合规则”决定，不能同时填写“字段映射”。聚合源字段目前不能写表达式。

```bash
uv run excelflow validate --plan examples/tutorial/05_aggregation.xlsx
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

全部函数和空值语义见[模板参考](../reference/template-reference.md#aggregation-rules)。
