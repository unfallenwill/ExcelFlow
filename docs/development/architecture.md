# 架构设计

ExcelFlow 使用标准 `src` 包布局，核心调用链为：

```text
CLI
  → ExtractionService
      → ExcelSpecRepository
      → SpecValidator
      → PandasExtractionEngine
          → SafeExpressionEvaluator
      → OutputWriterFactory → OutputWriter
```

## 模块职责

- `cli.py`：解析 `template/validate/preview/run` 命令；成功返回 0，校验失败或捕获到异常时返回 1。
- `repository.py`：读取五张基础配置表及两个可选聚合配置表，构造 `ExtractionSpec`。
- `schema.py`：配置模型、校验结果和字段名称规则。
- `validator.py`：执行不依赖源数据的计划结构校验。源 Sheet、源表头、源字段和运行期类型问题由 Engine 在 `run` 时发现。
- `engine.py`：读取源 Sheet、按顺序合并 DataFrame、过滤，并执行逐行字段映射或分组聚合。
- `expression.py`：用受限 AST 解释器计算衍生列，避免执行任意 Python。
- `output.py`：CSV、JSONL、XLSX 输出策略及工厂。
- `service.py`：编排读取、校验、启用检查、执行与写出。
- `template.py`：生成计划模板及下拉选项。

## 设计原则

Repository 隔离 Excel 配置读取；Strategy 让执行引擎和 Writer 可替换；Factory 集中选择 Writer；Service 保持 CLI 与业务编排分离。构造函数依赖注入使测试可以替换 Repository、Validator、Engine 和 Writer Factory。

运行时会把任务声明的所有 Sheet 读入内存并交给 Pandas，因此当前定位是中小规模 Excel 数据加工，不是流式或超大数据引擎。

表达式能力通过显式 AST 节点和函数白名单扩展，不能改用 `eval()`。逐行函数由表达式解释器映射到受控 Pandas/NumPy/标量操作；跨行函数由 Engine 的独立聚合阶段执行。普通路径为“读取→关联→过滤→字段映射”，聚合路径为“读取→关联→过滤→分组聚合”。新增配置字段时应同时更新模板、Schema 常量、Repository/Validator、教程与测试。
