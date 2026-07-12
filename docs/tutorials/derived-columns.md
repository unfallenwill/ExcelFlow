# 教程 4：复合关联与衍生列

本课把订单、客户和订单明细三个 Sheet 放到一起，并计算每条明细的金额：

```text
明细金额 = 数量 × 单价
```

可运行计划是 [`examples/tutorial/04_derived_columns.xlsx`](https://github.com/unfallenwill/ExcelFlow/blob/main/examples/tutorial/04_derived_columns.xlsx)。

## 为什么需要复合关联键

订单明细中存在 `order_id = 1003`、`tenant = A`，但订单 1003 属于租户 B。如果只按 `order_id` 关联，这条别人的租户数据会被错误接入。

正确要求是两个字段同时相同：

```text
订单号相同 AND 租户相同
```

因此“关联关系”中相同关联顺序填写两行：

| 关联顺序 | 关联类型 | 左侧字段 | 右侧对象 | 右侧字段 |
|---:|---|---|---|---|
| 2 | LEFT JOIN | o.order_id | i | i.order_id |
| 2 | LEFT JOIN | o.tenant | i | i.tenant |

相同顺序表示它们共同描述一次关联，并用 AND 组合。顺序 1 已经把客户 `c` 加入结果；顺序 2 再加入明细 `i`。

## 为什么关联顺序不能随意填

可以把关联理解为搭积木：

```text
第一步：订单 o 找到客户 c
第二步：已有的订单 o 找到明细 i
```

左侧字段只能来自主表或之前已经关联的对象。如果还没有加入客户 `c`，就不能用 `c.region_id` 去关联下一个对象。

## 用表达式计算新列

在“字段映射”新增一行：

| 源字段 | 目标字段 | 目标类型 | 转换表达式 | 字段顺序 |
|---|---|---|---|---:|
|  | line_amount | decimal | coalesce(i.quantity, 0) * coalesce(i.unit_price, 0) | 6 |

这是一个衍生列，所以“源字段”留空。“转换表达式”说明怎样用已有列计算结果。

为什么不直接写 `i.quantity * i.unit_price`？因为 LEFT JOIN 找不到明细时，这两个值可能为空；空值参与乘法仍会得到空值。`coalesce(值, 默认值)` 会在值为空时改用默认值 0，使“没有明细”的金额明确为 0。

## 常用逐行函数

表达式支持字段引用、数字、括号、算术、比较与安全函数。常见函数包括：

- `+` 加法
- `-` 减法
- `*` 乘法
- `/` 除法
- `%` 取余
- `coalesce(value, default)` 空值替换
- `abs(value)` 绝对值
- `round(value, digits)` 按 Pandas 规则保留指定位数；字段计算应显式填写 `digits`
- `clip(value, lower, upper)` 把数值限制在范围内
- `trim`、`upper`、`lower`、`concat`、`concat_ws` 清洗和拼接文本
- `to_number`、`to_string`、`to_date` 转换类型
- `dateformat`、`date_add`、`date_diff` 处理日期
- `if_else(condition, true_value, false_value)` 生成条件列

例如金额保留两位小数：

```text
round(coalesce(i.quantity, 0) * coalesce(i.unit_price, 0), 2)
```

例如生成订单标签：`concat_ws("-", upper(trim(o.tenant)), to_string(o.order_id))`。

表达式只能引用源 Sheet 字段，不能引用另一个目标字段或衍生列。`sum`、`count` 等跨行计算应填写“聚合规则”，不能写在转换表达式中。完整签名、空值行为和限制见[表达式参考](../reference/expressions.md)。

## 运行综合示例

本例还保留“已支付且金额不少于 100”的订单：

```bash
uv run excelflow validate --plan examples/tutorial/04_derived_columns.xlsx
uv run excelflow preview --plan examples/tutorial/04_derived_columns.xlsx --task lesson_04
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

原始订单只有一行 1001，但它有两条明细，所以关联后变成两行。这正是“一对多关联”的预期行为。

回到[模板填写完整教程](template-tutorial.md)，了解如何把这些步骤迁移到自己的业务。
