# 转换表达式参考

“字段映射”的“转换表达式”用于按行生成输出列。ExcelFlow 使用受限的语法树解释器执行表达式，不调用 Python `eval()`。

## 字段与常量

源字段使用 `对象别名.列名`：

```text
i.quantity
o.amount
```

表达式按 Python 属性语法解析，因此表达式中使用的别名和列名还必须能作为 Python 名称和属性被语法解析（不要使用 Python 关键字）。虽然普通字段映射允许名称中出现 `$`，包含 `$` 的字段不能直接写入转换表达式。

支持由 Python AST 表示为单个常量的字面值，例如数字、字符串、`True`、`False` 和 `None`；常见示例为 `0`、`1.13`、`"unknown"`。常量主要用作字段运算或函数的参数；只有常量、不产生 Series 的表达式目前不能保证通过后续所有目标类型转换。字段必须在关联后的数据中存在，否则运行时报错。列表、元组和字典不是受支持的常量。

## 运算符

| 语法 | 用途 | 示例 |
|---|---|---|
| `+` | 加法或兼容类型的拼接 | `o.amount + o.tax` |
| `-` | 减法 | `o.actual - o.plan` |
| `*` | 乘法 | `i.quantity * i.unit_price` |
| `/` | 除法 | `o.total / o.quantity` |
| `%` | 取余 | `o.sequence % 2` |
| `+value` | 一元正号 | `+o.amount` |
| `-value` | 一元负号 | `-o.amount` |
| `( )` | 控制计算顺序 | `(o.amount - o.discount) * 1.13` |

不支持比较运算、布尔运算、条件表达式、变量赋值、列表推导、任意方法调用或任意 Python 函数。

## coalesce

```text
coalesce(value, default)
```

将 `value` 中的空值替换为 `default`。第一个参数可以是字段或标量；处理字段时调用 Pandas `fillna`。若第一个参数是非空标量则原样返回，是空标量则返回默认值。

```text
coalesce(i.quantity, 0) * coalesce(i.unit_price, 0)
```

参数数量必须为 2。

## abs

```text
abs(value)
```

返回绝对值，可用于字段或数值表达式：

```text
abs(o.actual_amount - o.plan_amount)
```

参数数量必须为 1。

## round

```text
round(value)
round(value, digits)
```

对数值取整，或保留指定的小数位数：

```text
round(i.quantity * i.unit_price, 2)
```

语法解释器接受 1 或 2 个参数。对标量调用时使用 Python 内置 `round` 的参数规则。**当前版本对字段（Series）调用时必须显式提供 `digits`**，例如 `round(o.amount, 0)`；省略它会在运行时报错。Series 的 `digits` 会转换为整数。

## clip

```text
clip(series, lower, upper)
```

把字段值限制在闭区间 `[lower, upper]` 内：小于下界时返回下界，大于上界时返回上界，区间内保持原值。

```text
clip(s.score, 0, 100)
```

第一个参数必须是字段或计算结果产生的 Pandas Series，且参数数量必须为 3。标量作为第一个参数不受支持。

正负超窗天数示例：

```text
v.actual_day - clip(v.actual_day, v.plan_day - v.window_days, v.plan_day + v.window_days)
```

负数表示早于允许窗口，正数表示晚于允许窗口，`0` 表示在窗口内。

## 嵌套

受支持的函数和运算可以嵌套：

```text
round(clip(coalesce(i.quantity, 0) * coalesce(i.unit_price, 0), 0, 10000), 2)
```

## 错误与限制

- 表达式只能引用源 Sheet 字段，不能引用“目标字段”。
- 各字段映射相互独立，衍生列不能引用另一衍生列。
- 不支持 `if`、`if_else`、比较和布尔条件。
- 不支持 `sum`、`min`、`max`、字符串方法、日期方法或自定义函数。
- 不支持聚合或跨行计算；表达式面向逐行的 Series 运算。
- 除零和不兼容的数据类型遵循 Pandas 或 Python 的运算行为，可能得到无穷值，也可能抛出异常；ExcelFlow 不额外替换结果。
- 函数名不区分大小写；字段别名与列名区分大小写。

不支持的语法、函数参数数量错误或非法调用统一报告为转换表达式不受支持；不存在的字段会单独指出字段名称。
