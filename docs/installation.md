# ExcelFlow 安装与运行教程

ExcelFlow 安装后会提供一个名为 `excelflow` 的命令行程序。

## PyPI 页面

ExcelFlow 已发布到 PyPI：<https://pypi.org/project/excelflow/>

## 推荐方式：使用 uv tool

`uv tool` 会为 ExcelFlow 创建独立环境，同时把 `excelflow` 命令安装到用户命令目录，不会污染其他 Python 项目。

### 从 PyPI 安装

```bash
uv tool install excelflow
```

### 安装 GitHub 开发版

```bash
uv tool install --force git+https://github.com/unfallenwill/ExcelFlow.git
```

### 验证安装

```bash
excelflow --help
```

如果安装成功，会看到以下子命令：

```text
template
validate
preview
run
```

### 升级

升级到 PyPI 最新版本：

```bash
uv tool upgrade excelflow
```

### 卸载

```bash
uv tool uninstall excelflow
```

## 使用 pipx 安装

如果已经安装 pipx，可以直接从 PyPI 安装：

```bash
pipx install excelflow
```

验证：

```bash
excelflow --help
```

## 在虚拟环境中使用 pip

### Linux 和 macOS

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install excelflow
excelflow --help
```

### Windows PowerShell

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install excelflow
excelflow --help
```

退出虚拟环境：

```bash
deactivate
```

## 第一次运行

### 1. 生成计划模板

```bash
excelflow template extraction_plan.xlsx
```

当前目录会生成 `extraction_plan.xlsx`，其中包括：

- 抽取计划
- 数据对象
- 关联关系
- 字段映射
- 过滤条件
- 填写说明

### 2. 编辑计划

打开 `extraction_plan.xlsx`，至少完成以下配置：

1. 在“抽取计划”中填写任务ID并设置为启用。
2. 在“数据对象”中填写源 Excel 的 Sheet 名称和对象别名。
3. 指定唯一的主表。
4. 在“字段映射”中选择需要输出的字段。
5. 如有需要，填写关联关系、过滤条件和转换表达式。

源 Excel 文件路径不写入计划，而是在执行命令中传入。

## 填写计划后的标准流程

完成 `extraction_plan.xlsx` 后，依次执行“校验 → 预览 → 抽取”。

需要准备两个不同的 Excel 文件：

```text
extraction_plan.xlsx  抽取计划和规则
source.xlsx           实际需要处理的数据
```

`source.xlsx` 中的 Sheet 名称、表头行和字段名必须与计划文件中的声明一致。执行前还要确认“抽取计划”里的“启用”已经设置为“是”。

假设任务ID为 `order_report`，完整流程为：

```bash
excelflow validate extraction_plan.xlsx
excelflow preview extraction_plan.xlsx order_report
excelflow run extraction_plan.xlsx order_report source.xlsx csv output/order_report.csv
```

命令格式：

```text
excelflow run <计划文件> <任务ID> <源数据文件> <输出格式> <输出路径>
```

不同输出格式示例：

```bash
# CSV
excelflow run extraction_plan.xlsx order_report source.xlsx csv output/order_report.csv

# Excel
excelflow run extraction_plan.xlsx order_report source.xlsx xlsx output/order_report.xlsx

# JSONL
excelflow run extraction_plan.xlsx order_report source.xlsx jsonl output/order_report.jsonl
```

成功时会输出：

```text
抽取完成: 120 行 -> output/order_report.csv
```

### 3. 校验计划

```bash
excelflow validate extraction_plan.xlsx
```

成功时输出：

```text
校验通过
```

如果任务ID、字段别名、关联顺序或目标类型不合法，命令会输出具体错误。

### 4. 预览执行计划

假设任务ID为 `order_report`：

```bash
excelflow preview extraction_plan.xlsx order_report
```

预览会显示：

- 需要读取的 Sheet 和对象别名。
- Sheet 之间的关联顺序。
- 过滤条件数量。
- 输出字段数量。

### 5. 执行抽取

命令格式：

```text
excelflow run <计划Excel> <任务ID> <源数据Excel> <输出格式> <输出路径>
```

输出格式支持：

```text
csv
jsonl
xlsx
```

例如输出 CSV：

```bash
excelflow run \
  extraction_plan.xlsx \
  order_report \
  source.xlsx \
  csv \
  output/order_report.csv
```

输出 JSONL：

```bash
excelflow run extraction_plan.xlsx order_report source.xlsx jsonl output/order_report.jsonl
```

输出 Excel：

```bash
excelflow run extraction_plan.xlsx order_report source.xlsx xlsx output/order_report.xlsx
```

输出目录不存在时，ExcelFlow 会自动创建。

## 使用 Python 模块入口

如果 `excelflow` 命令没有出现在 PATH 中，但包已经安装，可以使用：

```bash
python -m excelflow --help
python -m excelflow validate extraction_plan.xlsx
```

在 uv 项目环境中：

```bash
uv run python -m excelflow validate extraction_plan.xlsx
```

## 运行项目自带教程

克隆仓库后，可以直接运行递进教学示例：

```bash
git clone https://github.com/unfallenwill/ExcelFlow.git
cd ExcelFlow
uv sync
uv run python examples/generate.py
```

执行第一课：

```bash
uv run excelflow run \
  examples/tutorial/01_single_sheet.xlsx \
  lesson_01 \
  examples/tutorial/source.xlsx \
  csv \
  examples/tutorial/output/01_orders.csv
```

完整教程参见 [examples/README.md](../examples/README.md)。

## 常见问题

### 找不到 `excelflow` 命令

先尝试：

```bash
python -m excelflow --help
```

如果模块入口可用，说明包已安装，但用户命令目录未加入 PATH。

使用 uv 时可以执行：

```bash
uv tool update-shell
```

然后重新打开终端。

使用 pipx 时可以执行：

```bash
pipx ensurepath
```

### 提示任务未启用

在计划文件的“抽取计划”工作表中，将对应任务的“启用”设置为“是”。

### 提示 Sheet 不存在

检查“数据对象”中的 `Sheet名称` 是否与源 Excel 底部显示的工作表名称完全一致。

### 提示字段不存在

确认：

- 字段名与源 Excel 表头一致。
- 字段使用 `对象别名.列名` 格式。
- 表头行配置正确。

### 输出格式与文件后缀不同

ExcelFlow 根据命令中的“输出格式”参数选择 Writer，不根据文件后缀推断格式。建议保持两者一致，例如：

```bash
excelflow run plan.xlsx task source.xlsx csv result.csv
```
