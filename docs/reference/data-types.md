# 目标数据类型参考

每条字段映射必须指定目标类型。模板下拉框及计划校验仅接受 `integer`、`decimal`、`string`、`datetime`。

| 目标类型 | Pandas 输出类型 | 转换方式 | 常见用途 |
|---|---|---|---|
| `integer` | 可空整数 `Int64` | `to_numeric(errors="raise")` 后转 `Int64` | 编号、数量、天数 |
| `decimal` | 可空浮点 `Float64` | `to_numeric(errors="raise")` 后转 `Float64` | 金额、比率、度量值 |
| `string` | Pandas `string` | `astype("string")` | 名称、状态、文本编号 |
| `datetime` | Pandas datetime | `to_datetime(errors="raise")` | 日期和时间 |

## integer

```text
integer
```

先将值解析为数字，再转换为 Pandas 可空整数。空值可保留为缺失值。包含无法解析的文本或非整数数值时，转换会失败。

适合数量和超窗天数。带前导零的业务编号应使用 `string`，否则前导零会丢失。

## decimal

```text
decimal
```

转换为 Pandas 可空 64 位浮点数。名称中的 `decimal` 不表示定点十进制；金额仍采用浮点表示，必要时可先用表达式 `round(value, 2)` 控制小数位。

## string

```text
string
```

转换为 Pandas 字符串类型，空值保留为 Pandas 缺失值。适合姓名、状态、代码和不参与数值运算的编号。

源 Excel 在读取阶段可能已把纯数字单元格解释为数值；如果编号必须保留前导零，应确保源单元格本身保存为文本。

## datetime

```text
datetime
```

使用 Pandas 日期解析并在无法解析时失败。输入可以是 Excel 日期单元格或 Pandas 可识别的日期文本。数字输入也会遵循 `pandas.to_datetime` 的默认规则（通常按 Unix 纳秒解释），未被特殊处理；因此建议源数据使用真实 Excel 日期或明确且一致的日期文本。

JSONL 输出使用 ISO 日期格式。CSV 和 XLSX 的具体展示形式由 Pandas 写出器和查看软件决定。

## 转换发生的阶段

执行顺序为：

```text
读取 Sheet → 关联 → 过滤 → 字段映射/表达式 → 目标类型转换 → 写出
```

因此过滤条件依据源字段的读取类型执行，而目标类型只影响最终映射列。转换使用严格错误模式；遇到不能解析的值时任务失败，不会静默改为空值。
