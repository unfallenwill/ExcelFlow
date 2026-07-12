# ExcelFlow 文档

ExcelFlow 让用户通过 Excel 模板描述数据抽取规则，再由 Pandas 完成读取、关联、过滤、字段转换和输出。

## 我是第一次使用

建议按顺序阅读：

1. [安装 ExcelFlow](getting-started/installation.md)
2. [五分钟快速开始](getting-started/quickstart.md)
3. [理解核心概念](getting-started/core-concepts.md)
4. [模板填写完整教程](tutorials/template-tutorial.md)

## 我想学习某项功能

- [抽取单个 Sheet](tutorials/single-sheet.md)
- [配置过滤条件](tutorials/filters.md)
- [关联多个 Sheet](tutorials/joins.md)
- [计算衍生列](tutorials/derived-columns.md)
- [计算正负超窗天数](how-to/window-overrun.md)
- [处理关联后的空值](how-to/handle-missing-values.md)
- [选择 INNER JOIN 或 LEFT JOIN](how-to/choose-join-type.md)
- [在一个计划中管理多个任务](how-to/multi-task-plan.md)

## 我需要查一个字段或命令

- [模板字段参考](reference/template-reference.md)
- [过滤运算符参考](reference/filter-operators.md)
- [表达式参考](reference/expressions.md)
- [目标数据类型参考](reference/data-types.md)
- [命令行参考](reference/cli.md)

## 我遇到了问题

查看[故障排查](troubleshooting.md)，其中按“任务无法执行”“结果行数变化”“字段为空”“表达式失败”等用户可见现象组织解决办法。

## 我想参与开发

- [贡献指南](https://github.com/unfallenwill/ExcelFlow/blob/main/CONTRIBUTING.md)
- [开发环境](development/development.md)
- [架构说明](development/architecture.md)
- [测试指南](development/testing.md)
- [发布流程](development/release.md)
- [安全策略](https://github.com/unfallenwill/ExcelFlow/blob/main/SECURITY.md)
- [行为准则](https://github.com/unfallenwill/ExcelFlow/blob/main/CODE_OF_CONDUCT.md)
- [MIT 许可证](https://github.com/unfallenwill/ExcelFlow/blob/main/LICENSE)

## 可运行示例

仓库的 [`examples/`](https://github.com/unfallenwill/ExcelFlow/tree/main/examples) 目录包含源数据、计划文件、生成脚本和预期结果。教程文档使用的案例均来自这些文件。
