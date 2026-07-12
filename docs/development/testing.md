# 测试指南

项目当前使用 Python 标准库 `unittest`。

运行全部测试：

```bash
uv run python -m unittest discover -s tests -v
```

现有测试分为：

- `test_cli.py`：控制台与模块入口、长短选项、输出内容、错误流和旧语法拒绝；
- `test_expression.py`、`test_expression_functions.py`：安全表达式、字符串/日期/条件/数值函数和恶意语法拦截；
- `test_aggregation.py`：聚合语义、模板兼容、校验和真实 Excel 端到端汇总；
- `test_components.py`：过滤运算符、输出 Writer、Service 错误路径；
- `test_main.py`：真实 Excel 计划和源数据的端到端抽取。

测试数量会随功能演进而变化，应以完整测试命令的实际输出为准，不在文档中维护固定数量。

新增功能至少应测试：

1. 正常配置的预期输出；
2. 非法配置是否在校验或执行时给出明确错误；
3. 空值、类型转换和无匹配关联等边界；
4. 必要时通过真实 Excel 覆盖完整链路。

测试应使用临时目录创建输入和输出，避免污染仓库。涉及模板的测试应通过 `create_template` 生成基础文件，再修改所需单元格。

覆盖率不是项目运行依赖。可用 uv 临时提供 `coverage` 并运行分支覆盖率：

```bash
uv run --with coverage coverage run --branch -m unittest discover -s tests
uv run --with coverage coverage report -m
```

第二条命令读取第一条命令生成的 `.coverage` 文件。该文件是本地测试产物，不应提交。覆盖率用于发现遗漏，不应替代对业务边界的断言；项目目前没有在 `pyproject.toml` 或 CI 中声明强制覆盖率门槛。提交前必须保证完整测试通过。
