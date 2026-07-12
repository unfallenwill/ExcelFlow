# 先理解 ExcelFlow 在做什么

可以把 ExcelFlow 看成一张“数据加工订单”。你用一个 Excel 文件说明想怎样加工数据，ExcelFlow 再按这份说明处理另一个 Excel 文件。

```text
计划文件 extraction_plan.xlsx：写加工规则
源数据 source.xlsx：放真正要处理的数据
输出文件 result.csv：加工后的结果
```

计划文件和源数据文件不是同一个文件。这样做的好处是：规则只写一次，以后换一份同样结构的源数据，还可以重复执行。

## 一项任务经过哪些步骤

```text
选择任务
  → 从主 Sheet 开始
  → 补充其他 Sheet 的信息
  → 保留符合条件的行
  → 选择、改名、逐行计算，或者分组汇总
  → 写入 CSV、JSONL 或 Excel
```

计划模板里的工作表分别回答一个问题。

| 工作表 | 它回答的问题 | 容易理解的比喻 |
|---|---|---|
| 抽取计划 | 这次要做哪项工作？ | 加工订单 |
| 数据对象 | 原始数据在哪些 Sheet？ | 原料清单 |
| 关联关系 | 不同 Sheet 的记录怎样对应？ | 把原料对齐 |
| 字段映射 | 每一行要输出或计算哪些列？ | 单件加工 |
| 过滤条件 | 哪些记录应该保留？ | 筛选原料 |
| 分组字段 | 哪些记录算作同一组？ | 分装归类 |
| 聚合规则 | 每组如何计数、求和或合并？ | 汇总称重 |

## 为什么要有任务ID

一个计划文件可以放多个任务，例如“订单清单”和“客户清单”。`任务ID` 是任务的唯一短名称，也是执行命令中用来选择任务的名称：

```bash
excelflow run --plan extraction_plan.xlsx --task order_report \
  --source source.xlsx --format csv --output output.csv
```

这里的 `order_report` 就是任务ID。各工作表使用同一个任务ID，ExcelFlow 才知道哪些配置属于同一项工作。

## 为什么必须指定主表

主表决定“从哪些行开始”。例如以订单为主表，ExcelFlow 会先拿到每一条订单，再为订单补充客户姓名和商品信息。

主表不是“最重要的 Sheet”，而是结果的出发点。每个任务必须且只能有一个主表。

## 为什么要给 Sheet 起别名

对象别名是 Sheet 的短名称，例如把“订单”叫作 `o`，把“客户”叫作 `c`。以后使用：

```text
o.status          订单的 status 列
c.customer_name   客户的 customer_name 列
```

如果两个 Sheet 都有 `status` 列，只写 `status` 无法判断来自哪里。`别名.列名` 能消除这种歧义，也让长 Sheet 名更容易书写。

别名和列名应使用英文字母、数字和下划线，并且不能以数字开头。推荐使用简短英文别名和 `snake_case` 列名，例如 `order_detail`。这样不仅适用于直接取列，也能在转换表达式中正常引用。中文、空格、连字符等写法目前不受支持。

## 为什么需要关联关系

订单 Sheet 可能只有 `customer_id`，客户姓名却在客户 Sheet。ExcelFlow 必须知道怎样找到对应客户：

```text
订单的 customer_id = 客户的 customer_id
```

应优先使用稳定且唯一的编号关联，不要用姓名。姓名可能重复或改变，编号通常更可靠。

关联会改变结果行数：

- 一条订单对应一个客户，通常仍是一行。
- 一条订单对应多个明细，关联后会变成多行。
- `INNER JOIN` 会丢掉找不到对应记录的主表行。
- `LEFT JOIN` 会保留主表行，找不到的补充字段为空。

## 为什么条件要分组

假设只保留“已支付并且金额不少于 100”的订单，这两个要求必须同时满足，所以放在同一条件组：

```text
status = paid AND amount >= 100
```

同组内使用 AND，不同组之间使用 OR。分组是为了让 Excel 能表达“同时满足”和“满足任意一组”这两种常见业务规则。

## 为什么需要字段映射

源数据可能有很多列，结果只需要其中几列，而且列名可能需要改得更易懂。字段映射可以：

- 选择输出列；
- 修改输出列名；
- 决定列的先后顺序；
- 统一数据类型；
- 根据已有列计算新列。

例如 `数量 × 单价` 可以生成源数据中没有的 `line_amount`。

## 三个检查动作

养成以下顺序可以更早发现错误：

```bash
excelflow validate --plan extraction_plan.xlsx
excelflow preview --plan extraction_plan.xlsx --task order_report
excelflow run --plan extraction_plan.xlsx --task order_report \
  --source source.xlsx --format csv --output output.csv
```

- `validate` 检查计划本身是否完整、前后一致。
- `preview` 以文字列出要读取的 Sheet、关联步骤、过滤条件数量和输出字段数量，但不会读取或处理源数据。
- `run` 才真正读取源数据并生成结果。

接下来可以用[五分钟快速开始](quickstart.md)跑通第一个任务，或阅读[模板填写完整教程](../tutorials/template-tutorial.md)。
