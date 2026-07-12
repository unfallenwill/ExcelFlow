# 模板填写完整教程

这份教程面向第一次配置数据抽取的用户。目标不是让你记住所有字段，而是理解每一项配置解决什么问题。

我们始终使用同一套订单数据，并逐步从“取出三列”扩展到“多 Sheet 关联、过滤和计算明细金额”。示例文件位于 [`examples/tutorial/`](https://github.com/unfallenwill/ExcelFlow/tree/main/examples/tutorial)。

## 开始前：区分两个 Excel

```text
计划文件：写“怎样处理”的规则
源数据文件：放“要处理什么”的数据
```

执行时才把二者放在一起：

```bash
excelflow run <计划文件> <任务ID> <源数据文件> <输出格式> <输出路径>
```

计划不保存源文件路径，是为了让同一套规则可以反复处理不同批次、但结构相同的数据。

## 第一步：明确想得到什么

先用一句话描述目标，例如：

> 从订单中保留已支付且金额不少于 100 的记录，补充客户和商品信息，并计算每条商品明细金额。

再把句子拆开：

| 业务问题 | 应填写的工作表 |
|---|---|
| 这项工作叫什么 | 抽取计划 |
| 数据在哪些 Sheet | 数据对象 |
| 订单、客户和明细怎样对应 | 关联关系 |
| 哪些订单要保留 | 过滤条件 |
| 最终输出哪些列、怎样计算 | 字段映射 |

先写目标能避免“为了填满模板而填写”。没有关联时，关联关系可以为空；没有过滤时，过滤条件也可以为空。

## 第二步：填写抽取计划

| 任务ID | 启用 | 备注 |
|---|---|---|
| lesson_04 | 是 | 订单明细报表 |

- `任务ID` 是唯一短名称，各工作表都用它把配置归到同一任务。
- `启用` 为“是”才能执行；设为“否”可暂时停用但保留规则。
- `备注` 写给人看，建议描述输出用途。

任务ID会出现在命令中：

```bash
excelflow preview extraction_plan.xlsx lesson_04
```

## 第三步：声明数据对象和主表

| 任务ID | Sheet名称 | 对象别名 | 表头行 | 是否主表 | 备注 |
|---|---|---|---:|---|---|
| lesson_04 | 订单 | o | 1 | 是 | 从订单开始 |
| lesson_04 | 客户 | c | 1 | 否 | 补充客户姓名 |
| lesson_04 | 订单明细 | i | 1 | 否 | 补充商品明细 |

为什么需要这些信息：

- Sheet 名称告诉工具到源 Excel 的哪个标签页读取。
- 对象别名让后续字段能写成 `o.order_id`，明确来自哪个 Sheet。
- 表头行告诉工具哪一行是列名，而不是数据。
- 主表决定结果从哪些行开始。每个任务必须且只能有一个主表。

对象别名不是随意装饰。如果订单和客户都有 `status`，`o.status` 与 `c.status` 才能消除歧义。

为了让字段既能直接映射，也能用于转换表达式，对象别名和列名请只使用英文字母、数字和下划线，并且不要以数字开头。中文列名、空格和连字符目前不受支持。

## 第四步：按业务关系连接 Sheet

先问自己：“一条订单凭什么找到对应客户和明细？”答案应是稳定编号，而不是姓名或显示文字。

| 关联顺序 | 关联类型 | 左侧字段 | 右侧对象 | 右侧字段 | 含义 |
|---:|---|---|---|---|---|
| 1 | LEFT JOIN | o.customer_id | c | c.customer_id | 用客户编号补充客户 |
| 2 | LEFT JOIN | o.order_id | i | i.order_id | 订单号相同 |
| 2 | LEFT JOIN | o.tenant | i | i.tenant | 并且租户相同 |

相同关联顺序的多行共同组成一个复合键。第二次关联的真实含义是：

```text
o.order_id = i.order_id AND o.tenant = i.tenant
```

如果只用订单号，不同租户的同号数据可能错误混在一起。

`LEFT JOIN` 表示主表订单必须保留，找不到客户或明细时补充列为空；`INNER JOIN` 表示只有两边能对应的行才保留。选择时先问业务：“匹配不到的信息是要暴露出来，还是整行排除？”

关联顺序表示搭建过程。左侧字段必须来自主表或前面已经加入的对象。

## 第五步：把业务条件翻译成过滤条件

业务要求是：

```text
状态是 paid AND 金额 >= 100
```

所以两条条件放在同一组：

| 条件组 | 条件序号 | 字段 | 运算符 | 值1 |
|---:|---:|---|---|---|
| 1 | 1 | o.status | = | paid |
| 1 | 2 | o.amount | >= | 100 |

同组内用 AND，不同组之间用 OR。可以把每个条件组想成一套完整的“入选方案”：满足任意一套方案，就能留下。

常用规则：

- `IN` / `NOT IN`：值1用英文逗号分隔，例如 `paid,pending`。
- `BETWEEN`：值1是下界，值2是上界。
- `LIKE` / `NOT LIKE`：`%` 匹配任意长度文字，`_` 匹配一个字符。
- `IS NULL` / `IS NOT NULL`：不填写值1和值2。

## 第六步：设计最终输出

| 源字段 | 目标字段 | 目标类型 | 转换表达式 | 字段顺序 |
|---|---|---|---|---:|
| o.order_id | order_id | integer |  | 1 |
| c.customer_name | customer | string |  | 2 |
| i.product | product | string |  | 3 |
| i.quantity | quantity | integer |  | 4 |
| i.unit_price | unit_price | decimal |  | 5 |
|  | line_amount | decimal | coalesce(i.quantity, 0) * coalesce(i.unit_price, 0) | 6 |

字段映射相当于“成品清单”：选择列、改名、排序、统一类型，并计算源数据没有的新列。

目标类型下拉框有：

| 类型 | 适合内容 |
|---|---|
| integer | 订单号、数量等整数 |
| decimal | 金额、比率等小数 |
| string | 姓名、状态、代码等文本 |
| datetime | 日期和时间 |

最后一行是衍生列，因此源字段留空。`coalesce` 在数量或单价为空时使用 0，避免结果不明确地变成空值。

表达式还支持 `+ - * / %`、`abs`、`round` 和 `clip`。复杂公式应先用小样本验证，并在备注中写清业务含义。

## 第七步：先校验，再预览，最后运行

以仓库中的完整示例为例：

```bash
uv run excelflow validate examples/tutorial/04_derived_columns.xlsx
uv run excelflow preview examples/tutorial/04_derived_columns.xlsx lesson_04
uv run excelflow run examples/tutorial/04_derived_columns.xlsx lesson_04 \
  examples/tutorial/source.xlsx csv examples/tutorial/output/04_order_lines.csv
```

- 校验发现计划内部矛盾，例如没有主表、别名重复或关联顺序无效。它不读取源数据，因此 Sheet 或真实列名是否存在要到运行时确认。
- 预览帮助确认读取对象、关联、过滤和输出字段数量。
- 运行才读取源数据，并把结果写到指定位置。

输出格式支持 `csv`、`jsonl` 和 `xlsx`。格式由命令参数决定，建议输出路径使用对应后缀。

## 第八步：解释结果行数

综合示例最后只有订单 1001，因为过滤排除了其他订单；订单 1001 有两条商品明细，所以关联后得到两行：

```csv
order_id,customer,product,quantity,unit_price,line_amount
1001,张三,键盘,2,299.0,598.0
1001,张三,鼠标,1,129.0,129.0
```

遇到行数变化时按执行顺序检查：

1. 过滤会减少行数。
2. INNER JOIN 可能减少行数。
3. 一对多关联会增加行数。
4. LEFT JOIN 不丢主表行，但补充字段可能为空。
5. 行数异常暴增往往意味着关联字段不唯一或复合键漏填。

## 推荐的学习顺序

四份可运行教程会逐步展示每次增加了什么配置：

1. [单 Sheet 与字段映射](single-sheet.md)
2. [过滤条件与 AND/OR](filters.md)
3. [两个 Sheet 的关联](joins.md)
4. [复合关联键与衍生列](derived-columns.md)

把自己的业务迁移到 ExcelFlow 时，也建议按这个顺序逐层增加功能。每增加一层就校验并运行小样本，比一次填完所有规则更容易定位问题。
