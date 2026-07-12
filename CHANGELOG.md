# 更新日志

本项目的重要变化记录于此，格式参考 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)，版本遵循[语义化版本](https://semver.org/lang/zh-CN/)约定。

## [Unreleased]

### Added

- 重构后的开源项目文档体系。
- MkDocs Material 文档站和 GitHub Pages 自动部署。
- MIT 许可证及贡献、安全、行为准则文档。
- 字符串、日期、类型转换、条件判断及常用数值表达式函数。
- 可选“分组字段”和“聚合规则”配置表，以及计数、求和、平均值、极值和文本合并聚合。

### Changed

- **破坏性变更：** CLI 全面改用 Unix/POSIX 风格的命名选项（如 `--plan`/`-p`、`--task`/`-t`、`--source`/`-s`、`--format`/`-f` 和 `--output`/`-o`），不再支持旧的位置参数语法。

## [0.1.0]

初始公开版本。

### Added

- Excel 声明驱动的数据抽取。
- 单 Sheet 与多 Sheet、多级和复合键关联。
- 条件组过滤以及常用比较、集合、范围、模糊和空值运算符。
- 字段映射、目标类型转换和安全衍生列表达式。
- CSV、JSONL 和 XLSX 输出。
- `coalesce`、`abs`、`round` 与 `clip` 表达式函数。
- 模板生成、校验、预览和执行 CLI。
