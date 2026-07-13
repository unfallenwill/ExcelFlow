# 参与贡献

感谢你帮助改进 ExcelFlow。

## 提交问题

请先搜索已有 Issue。报告问题时提供：

- ExcelFlow 与 Python 版本；
- 使用的命令；
- 可公开的最小计划和源数据结构；
- 实际结果、预期结果和完整错误信息。

请勿上传包含个人信息、业务秘密、凭据或其他敏感内容的 Excel 文件。

## 开发流程

```bash
uv sync --group dev                              # 安装运行依赖与质量门禁工具（ruff / pyright / pre-commit）
uv run pre-commit install                        # 一次性：把质量门禁挂到 git commit
uv run python -m unittest discover -s tests -v
```

提交前会自动运行 ruff（lint + 格式化）、Pyright（`standard` 模式类型检查）与若干基础卫生检查；也可手动执行 `uv run pre-commit run --all-files`。CI 会在每次 push 与 PR 上运行同一套门禁加单测（见 `.github/workflows/quality.yml`）。

修改应保持范围清晰，并包含相应测试和文档。模板字段、CLI 或配置语义变化属于用户可见变化，必须同步更新示例和参考资料。

提交 Pull Request 时请说明：解决的问题、设计选择、验证方式以及兼容性影响。维护者可能要求把过大的变更拆分为更容易审查的提交。

架构与本地开发细节见 `docs/development/`。

## 安全问题

如果问题可能导致任意代码执行、敏感数据泄露或绕过表达式限制，请不要创建公开 Issue，改按 [安全政策](SECURITY.md) 私下报告。
