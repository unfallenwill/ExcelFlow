# 安装 ExcelFlow

ExcelFlow 是一个命令行工具。安装完成后，你会得到一个 `excelflow` 命令，用它生成模板、检查填写内容并执行抽取。

ExcelFlow 需要 Python 3.12 或更高版本。`uv tool` 通常会自动准备兼容的独立 Python 环境；使用 pipx 或虚拟环境时，请先确认对应环境的 Python 版本。

## 推荐：使用 uv tool

如果电脑已安装 [uv](https://docs.astral.sh/uv/)，运行：

```bash
uv tool install excelflow
```

`uv tool` 会把 ExcelFlow 放在独立环境中，不会影响电脑上的其他 Python 项目。

验证安装：

```bash
excelflow --help
```

看到 `template`、`validate`、`preview` 和 `run` 四个子命令，就表示安装成功。

升级或卸载：

```bash
uv tool upgrade excelflow
uv tool uninstall excelflow
```

## 使用 pipx

已经使用 pipx 的用户可以运行：

```bash
pipx install excelflow
excelflow --help
```

升级与卸载：

```bash
pipx upgrade excelflow
pipx uninstall excelflow
```

## 在 Python 虚拟环境中安装

Linux 或 macOS：

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install excelflow
excelflow --help
```

Windows PowerShell：

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install excelflow
excelflow --help
```

这种方式下，只有激活虚拟环境后才能直接使用 `excelflow`。

## 找不到 excelflow 命令怎么办

先关闭并重新打开终端，再运行 `excelflow --help`。如果仍然找不到：

- uv 用户运行 `uv tool update-shell`，然后重新打开终端。
- pipx 用户运行 `pipx ensurepath`，然后重新打开终端。
- 虚拟环境用户确认命令行开头显示 `(.venv)`。

安装完成后，继续阅读[五分钟快速开始](quickstart.md)。如果想先理解模板为什么分成多张工作表，请阅读[核心概念](core-concepts.md)。
