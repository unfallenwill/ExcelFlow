# 在一个计划文件中管理多个任务

当同一份源 Excel 需要生成多种结果时，可以把多个任务放进同一个计划文件。例如同时生成订单清单、客户清单和超窗检查结果。

在“抽取计划”中为每项工作填写唯一的任务ID：

| 任务ID | 启用 | 备注 |
|---|---|---|
| order_report | 是 | 订单报表 |
| customer_report | 是 | 客户报表 |
| window_check | 否 | 尚未投入使用 |

其他工作表中的每一行也要填写任务ID。ExcelFlow 用它判断一行配置属于哪项任务；任务之间不会共享数据对象、关联、字段映射或过滤条件。

每个任务都必须：

- 在“数据对象”中有且只有一个主表；
- 拥有自己唯一的对象别名集合；
- 将所有非主表对象通过“关联关系”接入；
- 在需要输出指定列时填写自己的“字段映射”。

命令通过任务ID选择任务：

```bash
excelflow preview extraction_plan.xlsx order_report
excelflow run extraction_plan.xlsx order_report source.xlsx xlsx output/orders.xlsx
```

“启用”为“否”的任务可以保留配置，但执行 `run` 会失败。`preview` 只读取并展示指定任务的声明：它不检查任务是否启用，也不会读取源数据或替代 `validate`。删除任务时，也应删除其他工作表中使用该任务ID的行，否则校验会报告任务ID不存在。
