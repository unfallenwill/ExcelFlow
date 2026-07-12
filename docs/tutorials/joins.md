# 教程 3：把两个 Sheet 的信息放到一起

订单 Sheet 只有客户编号，但报表希望显示客户姓名。本课用相同的 `customer_id` 把“订单”和“客户”对应起来。

可运行计划是 [`examples/tutorial/03_join.xlsx`](https://github.com/unfallenwill/ExcelFlow/blob/main/examples/tutorial/03_join.xlsx)。

## 为什么需要关联

订单数据：

| order_id | customer_id | amount |
|---:|---:|---:|
| 1001 | 10 | 727 |

客户数据：

| customer_id | customer_name |
|---:|---|
| 10 | 张三 |

ExcelFlow 不会猜测两张表的关系。我们要明确告诉它：订单中的客户编号，与客户表中相同编号的记录对应。

应使用 `customer_id` 这类稳定编号，不建议用客户姓名。同名客户会造成错误匹配，姓名变化也会让历史数据无法对应。

## 声明两个数据对象

| 任务ID | Sheet名称 | 对象别名 | 表头行 | 是否主表 | 备注 |
|---|---|---|---:|---|---|
| lesson_03 | 订单 | o | 1 | 是 | 主表 |
| lesson_03 | 客户 | c | 1 | 否 | 客户信息 |

订单是主表，意味着关联从订单行开始，再补充客户信息。是否保留找不到客户的订单，仍由下一步选择的 `LEFT JOIN` 或 `INNER JOIN` 决定。`o` 和 `c` 让字段来源一目了然。

## 填写关联关系

| 任务ID | 关联顺序 | 关联类型 | 左侧字段 | 右侧对象 | 右侧字段 |
|---|---:|---|---|---|---|
| lesson_03 | 1 | LEFT JOIN | o.customer_id | c | c.customer_id |

用日常语言翻译：

> 从订单开始，对每条订单拿 `customer_id` 去客户 Sheet 寻找相同编号；找到后补上客户信息。

“右侧对象”只填别名 `c`，“右侧字段”则填完整的 `c.customer_id`。

## LEFT JOIN 和 INNER JOIN 怎样选择

| 类型 | 找不到客户时 | 适合场景 |
|---|---|---|
| LEFT JOIN | 订单仍保留，客户字段为空 | 订单不能丢，缺失信息需要后续检查 |
| INNER JOIN | 订单被排除 | 结果只接受客户资料完整的订单 |

不确定时，通常先用 LEFT JOIN，观察哪些补充字段为空，再决定是否应该排除。

## 输出客户姓名

字段映射加入：

| 源字段 | 目标字段 | 目标类型 | 字段顺序 |
|---|---|---|---:|
| o.order_id | order_id | integer | 1 |
| c.customer_name | customer | string | 2 |
| o.amount | amount | decimal | 3 |

`c.customer_name` 能使用，是因为关联的第一步已经把客户对象 `c` 加入结果。

## 运行

```bash
uv run excelflow validate --plan examples/tutorial/03_join.xlsx
uv run excelflow preview --plan examples/tutorial/03_join.xlsx --task lesson_03
uv run excelflow run \
  --plan examples/tutorial/03_join.xlsx \
  --task lesson_03 \
  --source examples/tutorial/source.xlsx \
  --format xlsx \
  --output examples/tutorial/output/03_orders_with_customer.xlsx
```

命令会生成包含 3 行数据的 Excel 文件：

| order_id | customer | amount |
|---:|---|---:|
| 1001 | 张三 | 727.0 |
| 1002 | 李四 | 499.0 |
| 1003 | 李四 | 80.0 |

## 为什么关联后行数可能变化

如果一个客户编号只对应一个客户，订单行数通常不变。如果一条订单关联到三条明细，订单会展开成三行。这不是重复执行，而是一对多关系的自然结果。

若行数意外暴增，应检查关联字段是否唯一、是否漏掉复合键中的第二个字段。

下一课：[计算源数据中没有的新列](derived-columns.md)。
