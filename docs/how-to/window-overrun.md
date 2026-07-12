# 计算正负超窗天数

假设计划在第 28 天检查，允许前后浮动 3 天，那么第 25～31 天都合规。我们需要的不是“实际日与计划日相差几天”，而是“实际日越过最近边界几天”。

先给保存检查记录的 Sheet 配置对象别名 `v`，并确保它包含 `actual_day`、`plan_day` 和 `window_days` 三列。然后在“字段映射”中新增目标字段 `overrun_days`，目标类型选择 `integer`，转换表达式填写：

```text
v.actual_day - clip(v.actual_day, v.plan_day - v.window_days, v.plan_day + v.window_days)
```

`clip` 会把实际日限制在允许范围内：太早就取最早日，太晚就取最晚日，窗口内保持原值。再用实际日减去这个值，便得到带方向的超窗天数。

| 实际日 | 允许范围 | 输出 | 含义 |
|---:|---|---:|---|
| 22 | 25～31 | -3 | 提前超窗 3 天 |
| 35 | 25～31 | 4 | 延后超窗 4 天 |
| 30 | 25～31 | 0 | 未超窗 |

不要直接计算 `actual_day - plan_day`。计划日到边界之间属于允许浮动范围，不应计入超窗。

完成配置后依次运行：

```bash
excelflow validate extraction_plan.xlsx
excelflow preview extraction_plan.xlsx window_check
excelflow run extraction_plan.xlsx window_check source.xlsx csv output/window_check.csv
```

这里的 `validate` 只检查计划结构，不会读取 `source.xlsx`。列名是否存在、列值能否参与数值计算，要到执行 `run` 时才能确认。实际日、计划日和窗口天数应使用数值列；空值会沿计算传播，最终得到空值。
