# 转换表达式参考

“字段映射”的“转换表达式”用于逐行生成输出列。表达式由安全白名单解释器映射为受控的 Pandas、NumPy 或标量操作，不使用 `eval()`。跨行汇总请使用[分组字段与聚合规则](template-reference.md#group-fields)。

## 字段、常量和运算符

源字段写成 `对象别名.列名`，例如 `o.amount`。支持数字、字符串、`True`、`False`、`None`，以及 `+ - * / %`、一元 `+ -`、`== != > >= < <=`、`and or not` 和括号。字段名必须能按 Python 属性语法解析。

```text
if_else((o.amount >= 1000) and (o.status == "paid"), "大额", "普通")
```

空条件按 `False` 处理。函数名不区分大小写，字段名区分大小写。`if_else` 会先计算三个参数，不提供短路求值。

## 空值与数值函数

| 函数 | 作用 | 示例 |
|---|---|---|
| `coalesce(value, default)` | 用默认值替换空值 | `coalesce(o.amount, 0)` |
| `is_null(value)` | 判断是否为空 | `is_null(o.amount)` |
| `abs(value)` | 绝对值 | `abs(o.actual-o.plan)` |
| `round(value[, digits])` | 四舍五入，默认 0 位 | `round(o.amount, 2)` |
| `clip(value, lower, upper)` | 限制 Series 范围 | `clip(o.score, 0, 100)` |
| `ceil(value)` / `floor(value)` | 向上/向下取整 | `ceil(o.amount)` |
| `sqrt(value)` / `power(value, exponent)` | 平方根/幂 | `power(o.value, 2)` |
| `min_value(...)` / `max_value(...)` | 同一行多个值取最小/最大 | `max_value(o.amount, 0)` |

`min_value` 和 `max_value` 是逐行函数；聚合规则中的 `min`、`max` 是跨行函数。

## 字符串函数

字符串函数会把非字符串输入转换为 Pandas `string`。`concat`、`concat_ws`、大小写、清理、长度、替换和截取遇到空值时继续为空；`contains`、`startswith`、`endswith` 遇到空值时返回 `False`。需要把空值当空字符串拼接时，显式使用 `coalesce(value, "")`。

| 函数 | 作用 | 示例 |
|---|---|---|
| `concat(a, b, ...)` | 直接拼接，至少一个参数 | `concat(p.first, p.last)` |
| `concat_ws(separator, a, b, ...)` | 用分隔符拼接 | `concat_ws("-", p.region, p.code)` |
| `upper(value)` / `lower(value)` | 大小写转换 | `upper(p.code)` |
| `trim(value)` | 删除首尾空格 | `trim(p.name)` |
| `length(value)` | 字符长度 | `length(p.name)` |
| `replace(value, old, new)` | 按普通文本替换，不使用正则 | `replace(p.phone, " ", "")` |
| `substring(value, start[, length])` | 从 0 开始截取，支持负索引 | `substring(p.code, 0, 4)` |
| `contains(value, text)` | 是否包含普通文本 | `contains(p.name, "医院")` |
| `startswith(value, text)` | 是否以文本开头 | `startswith(p.code, "CN")` |
| `endswith(value, text)` | 是否以文本结尾 | `endswith(p.file, ".xlsx")` |

## 类型和日期函数

| 函数 | 作用 | 示例 |
|---|---|---|
| `to_number(value)` | 严格转换为数值 | `to_number(o.amount_text)` |
| `to_string(value)` | 转换为可空字符串 | `to_string(o.order_id)` |
| `to_date(value)` | 严格转换为日期时间 | `to_date(v.visit_date)` |
| `dateformat(value, format)` | 按 Python/Pandas 格式输出字符串 | `dateformat(v.date, "%Y-%m-%d")` |
| `year(value)` / `month(value)` / `day(value)` | 提取年月日 | `year(v.date)` |
| `date_add(value, amount, unit)` | 日期加减 | `date_add(v.date, 3, "day")` |
| `date_diff(left, right, unit)` | 返回 `left-right` 的差值 | `date_diff(v.actual, v.plan, "day")` |

`date_add` 和 `date_diff` 支持 `day/days`、`hour/hours`、`minute/minutes`，结果允许小数。`to_number` 使用严格数值转换；日期函数使用严格日期解析；非法输入或不支持的日期单位会让任务失败。其他数值函数遵循 Pandas/NumPy 行为，例如负数开平方可能得到空值并产生运行时警告。

## 条件函数

```text
if_else(condition, value_if_true, value_if_false)
```

条件可以是比较、`and`、`or`、`not` 或 `is_null` 的结果：

```text
if_else(is_null(o.amount), 0, o.amount)
if_else((o.amount >= 1000) and (o.status == "paid"), "大额已支付", "其他")
```

## 嵌套示例

```text
concat_ws("-", upper(trim(p.region)), to_string(p.customer_id))
round(clip(coalesce(i.quantity, 0) * coalesce(i.unit_price, 0), 0, 10000), 2)
v.actual_day - clip(v.actual_day, v.plan_day-v.window_days, v.plan_day+v.window_days)
```

## 限制和错误

- 只能引用源 Sheet 字段，不能引用目标字段或另一条衍生列。
- 不支持属性链、下标、列表、字典、lambda、方法调用、关键字参数、自定义函数或任意 Python 代码。
- `sum`、`count` 等聚合函数不能写在转换表达式中。
- `min_value`、`max_value` 会逐行跳过空值；同一行所有参数都为空时结果为空。
- `validate` 当前不解析表达式；字段、语法、函数参数及转换错误在 `run` 时报告。
- 除零和不兼容类型遵循 Pandas/Python 行为，可能返回无穷值或失败。
