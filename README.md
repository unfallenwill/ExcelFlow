# ExcelFlow

ExcelFlow 是一个由 Excel 计划驱动的数据抽取工具。业务用户在模板中声明需要读取的 Sheet、关联方式、过滤条件、输出字段和衍生列，ExcelFlow 使用 Pandas 执行这些规则并输出 CSV、JSONL 或 Excel。

[![PyPI](https://img.shields.io/pypi/v/excelflow)](https://pypi.org/project/excelflow/)
[![Python](https://img.shields.io/pypi/pyversions/excelflow)](https://pypi.org/project/excelflow/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

## 适合解决的问题

- 从一个 Excel Sheet 中选择、重命名和转换字段。
- 使用结构化条件筛选数据，而不需要编写代码。
- 在同一个工作簿中关联多个 Sheet。
- 通过安全表达式计算金额、差值和超窗天数等衍生列。
- 将同一份抽取计划保存、审核、复用和纳入版本管理。

## 安装

推荐使用 uv 安装为独立命令行工具：

```bash
uv tool install excelflow
```

也可以使用 pipx：

```bash
pipx install excelflow
```

详细说明参见[安装指南](docs/getting-started/installation.md)。

## 五分钟快速开始

生成计划模板：

```bash
excelflow template --output extraction_plan.xlsx
```

填写模板后，依次校验、预览和执行。假设任务ID为 `order_report`，源数据为 `source.xlsx`：

```bash
excelflow validate --plan extraction_plan.xlsx
excelflow preview --plan extraction_plan.xlsx --task order_report
excelflow run --plan extraction_plan.xlsx --task order_report \
  --source source.xlsx --format csv --output output/order_report.csv
```

完整操作过程参见[快速开始](docs/getting-started/quickstart.md)和[模板填写完整教程](docs/tutorials/template-tutorial.md)。

## 核心能力

- 单 Sheet 与多 Sheet 数据抽取。
- `INNER JOIN`、`LEFT JOIN`、多级关联和复合关联键。
- AND/OR 条件组及 13 种过滤运算符。
- `integer`、`decimal`、`string`、`datetime` 输出类型。
- 四则运算、`coalesce`、`abs`、`round`、`clip` 安全表达式。
- CSV、JSONL、XLSX 输出。
- 计划校验和 Pandas 执行计划预览。

## 文档

- [在线文档](https://unfallenwill.github.io/ExcelFlow/)
- [文档首页](docs/index.md)
- [安装指南](docs/getting-started/installation.md)
- [核心概念](docs/getting-started/core-concepts.md)
- [模板填写完整教程](docs/tutorials/template-tutorial.md)
- [模板字段参考](docs/reference/template-reference.md)
- [过滤运算符参考](docs/reference/filter-operators.md)
- [表达式参考](docs/reference/expressions.md)
- [命令行参考](docs/reference/cli.md)
- [故障排查](docs/troubleshooting.md)

## 示例

[`examples/`](examples/README.md) 提供四个递进案例：单 Sheet、过滤、Sheet 关联、复合关联键与衍生列。示例计划和源数据已生成，可以直接运行。

## 安全边界

字段转换由受限 AST 解释器执行，不使用 Python `eval()`。表达式只能访问已声明的源字段、常量和受支持函数，不能执行文件操作、系统命令或任意 Python 代码。

## 参与项目

- [贡献指南](CONTRIBUTING.md)
- [开发环境](docs/development/development.md)
- [架构说明](docs/development/architecture.md)
- [测试指南](docs/development/testing.md)
- [安全策略](SECURITY.md)
- [变更记录](CHANGELOG.md)

源码托管于 [GitHub](https://github.com/unfallenwill/ExcelFlow)，问题和功能建议请提交到 [GitHub Issues](https://github.com/unfallenwill/ExcelFlow/issues)。

ExcelFlow 使用 [MIT License](LICENSE) 发布。
