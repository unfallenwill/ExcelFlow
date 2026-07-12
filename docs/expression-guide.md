# 字段映射表达式指南

ExcelFlow 的“字段映射”工作表支持通过转换表达式生成衍生列。表达式由受限解释器执行，不使用 Python `eval()`。

## 字段引用

表达式中的变量引用源 Excel Sheet 的列，格式为：

```text
对象别名.列名
```

例如“数据对象”配置如下：

| Sheet名称 | 对象别名 |
|---|---|
| 订单 | o |
| 订单明细 | i |

字段引用示例：

```text
o.amount
i.quantity
i.unit_price
```

注意事项：

- 使用“对象别名”，不是 Sheet 名称。
- 列名必须与源 Excel 表头一致。
- 当前列名需要使用简单标识符，例如 `unit_price`。
- 表达式只能引用源字段，不能引用字段映射产生的目标字段。
- 衍生列之间目前不能相互引用。

## 支持的运算

表达式支持以下运算符：

```text
+  加法
-  减法
*  乘法
/  除法
%  取模
```

例如：

```text
i.quantity * i.unit_price
```

## `coalesce`：替换空值

语法：

```text
coalesce(value, default)
```

当 `value` 为空时返回 `default`，否则返回原值。

例如缺失数量或单价按 0 处理：

```text
coalesce(i.quantity, 0) * coalesce(i.unit_price, 0)
```

| quantity | unit_price | 结果 |
|---:|---:|---:|
| 2 | 100 | 200 |
| 空 | 100 | 0 |
| 2 | 空 | 0 |

典型场景：

- 缺失数量按 0 处理。
- 缺失折扣使用默认值。
- `LEFT JOIN` 后处理未匹配字段。

缺失折扣按 1 计算：

```text
o.amount * coalesce(o.discount_rate, 1)
```

## `abs`：取绝对值

语法：

```text
abs(value)
```

例如计算实际金额与计划金额之间的绝对差异：

```text
abs(o.actual_amount - o.plan_amount)
```

| actual_amount | plan_amount | 结果 |
|---:|---:|---:|
| 80 | 100 | 20 |
| 120 | 100 | 20 |

典型场景：

- 计划值与实际值的差异。
- 日期序号偏差。
- 将正负偏差统一转换为距离。

## `round`：四舍五入

语法：

```text
round(value)
round(value, digits)
```

例如计算含税金额并保留两位小数：

```text
round(o.amount * 1.13, 2)
```

计算单价并保留两位小数：

```text
round(o.total_amount / o.quantity, 2)
```

典型场景：

- 金额保留两位小数。
- 比率和百分比格式化。
- 消除浮点计算产生的多余尾数。

## `clip`：限制数值范围

语法：

```text
clip(value, lower, upper)
```

规则：

- 小于下限时返回下限。
- 大于上限时返回上限。
- 位于范围内时返回原值。

例如将评分限制在 0～100：

```text
clip(s.score, 0, 100)
```

| score | 结果 |
|---:|---:|
| -10 | 0 |
| 80 | 80 |
| 120 | 100 |

典型场景：

- 评分、比例或金额的封顶和保底。
- 限制异常值范围。
- 计算正负超窗天数。

## 正负超窗天数

假设：

```text
v.actual_day   实际检查日
v.plan_day     计划检查日
v.window_days  允许窗口天数
```

字段映射表达式：

```text
v.actual_day - clip(v.actual_day, v.plan_day - v.window_days, v.plan_day + v.window_days)
```

当计划日为第 28 天、允许窗口为 ±3 天时，允许范围为第 25～31 天：

| 实际检查日 | 表达式结果 | 含义 |
|---:|---:|---|
| 22 | -3 | 提前超窗 3 天 |
| 35 | 4 | 延后超窗 4 天 |
| 30 | 0 | 未超窗 |

结果约定：

- 负数表示早于允许窗口。
- 正数表示晚于允许窗口。
- 0 表示位于允许窗口内。

对应的字段映射可以配置为：

| 源字段 | 目标字段 | 目标类型 | 转换表达式 |
|---|---|---|---|
| 留空 | overrun_days | integer | `v.actual_day - clip(v.actual_day, v.plan_day - v.window_days, v.plan_day + v.window_days)` |

## 函数组合

函数可以嵌套使用。例如计算明细金额，将空值按 0 处理，将结果限制在 0～10000，最后保留两位小数：

```text
round(clip(coalesce(i.quantity, 0) * coalesce(i.unit_price, 0), 0, 10000), 2)
```

## 目标类型

“目标类型”列可以从下拉框选择：

| 类型 | 用途 |
|---|---|
| `integer` | 整数 |
| `decimal` | 小数 |
| `string` | 字符串 |
| `datetime` | 日期时间 |

表达式计算完成后，ExcelFlow 会按照目标类型转换输出列。

## 当前限制

转换表达式目前不支持：

- `if` 或条件分支。
- 比较表达式，例如 `o.amount > 100`。
- 任意 Python 函数调用。
- 属性访问、文件操作或系统命令。

不支持的表达式会在执行时被拒绝，不会交给 Python `eval()`。
