# 发布指南

发布需要 PyPI 项目权限和 GitHub 仓库权限。不要把 PyPI Token 写入仓库、命令历史文档或 Issue。

## 1. 发布前检查

1. 更新 `pyproject.toml` 中的版本号；
2. 更新 `CHANGELOG.md`，把待发布内容移入对应版本；
3. 运行完整测试；
4. 检查用户文档和示例与当前行为一致；
5. 提交版本号、CHANGELOG 和相关变更，确认工作区干净且发布提交已经推送。

```bash
uv run python -m unittest discover -s tests -v
git status --short
```

确认 `pyproject.toml` 的 `requires-python` 和依赖范围仍符合实际支持情况。当前项目要求 Python 3.12 或更高版本。

## 2. 构建

```bash
uv build
```

构建结果位于 `dist/`，通常包含 wheel 和源码包。发布前检查构建产物只来自本次版本，避免误传 `dist/` 中遗留的旧版本文件。建议在干净虚拟环境中安装本次 wheel，并运行：

```bash
excelflow --help
excelflow template --output /tmp/extraction_plan.xlsx
```

## 3. 发布 PyPI

在本地安全设置 Token 后发布：

```bash
export UV_PUBLISH_TOKEN='pypi-...'
uv publish
```

`uv publish` 默认会处理 `dist/` 下的分发文件，因此发布前必须确认其中的文件名和版本正确。发布后从 PyPI 新建隔离环境，明确安装刚发布的版本并再次验证 CLI。版本号不可在 PyPI 上覆盖；发现问题时应发布新的补丁版本。

## 4. GitHub 发布

在已经推送的发布提交上创建与版本一致的标签（例如 `v0.1.1`），推送标签并创建 GitHub Release。Release 说明应来自 CHANGELOG，明确新增、变化、修复和兼容性注意事项。标签、PyPI 版本和 `pyproject.toml` 版本必须一致。

项目使用 MIT License。发布前应确认源码分发包和 wheel 元数据正确包含许可证声明。
