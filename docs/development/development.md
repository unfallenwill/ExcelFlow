# 开发指南

## 环境要求

- Python 3.12 或更高版本
- [uv](https://docs.astral.sh/uv/)（推荐）

克隆仓库后安装锁定依赖：

```bash
uv sync
```

运行 CLI：

```bash
uv run excelflow --help
uv run python -m excelflow --help
```

生成并校验模板：

```bash
uv run excelflow template /tmp/extraction_plan.xlsx
uv run excelflow validate /tmp/extraction_plan.xlsx
```

## 项目布局

```text
src/excelflow/   Python 包
tests/           单元与端到端测试
examples/        可运行示例与生成脚本
docs/            用户和维护者文档
pyproject.toml   包元数据、依赖和 CLI 入口
uv.lock          锁定依赖
```

## 修改约定

- 保持 CLI 只负责输入输出，业务逻辑放在 Service 或相应组件中。
- 使用依赖注入维持组件可替换性。
- 新功能同时补充成功路径、错误路径和边界测试。
- 模板或语义发生变化时，同步更新示例和文档。
- 不执行任意用户表达式；扩展表达式时维持白名单设计。
- 提交前运行完整测试，并确认 `git diff` 中没有生成物、临时 Excel 锁文件或敏感数据。
