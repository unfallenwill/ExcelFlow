# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project is

ExcelFlow is a declarative Excel-data extraction tool. A business user fills in a plan workbook (which sheets to read, how to join, filter, map/derive columns, and optionally aggregate), and ExcelFlow executes it with Pandas, outputting CSV/JSONL/XLSX. There is no SQL and no `eval` — derived columns run through a restricted AST interpreter.

## Commands

Install dependencies (uses `uv`, Python ≥3.12):

```bash
uv sync
```

Run the CLI (the venv must be on PATH; `uv run` handles activation):

```bash
uv run excelflow --help
uv run python -m excelflow --help
```

Run the full test suite:

```bash
uv run python -m unittest discover -s tests -v
```

Run a single test (the `tests.` prefix is required when running from the repo root):

```bash
uv run python -m unittest tests.test_components.EngineFilterSemanticsTest.test_same_group_combines_with_AND -v
```

Branch coverage (coverage is not a declared dependency — pull it transiently):

```bash
uv run --with coverage coverage run --branch -m unittest discover -s tests
uv run --with coverage coverage report -m
```

Build/preview docs (strict mode):

```bash
uv sync --group docs
uv run --group docs mkdocs build --strict
uv run --group docs mkdocs serve
```

## Architecture

Call chain:

```
CLI (cli.py) → ExtractionService → ExcelSpecRepository   (load plan)
                              → SpecValidator            (structural checks)
                              → PandasExtractionEngine → SafeExpressionEvaluator
                              → OutputWriterFactory → Csv/JsonLines/ExcelWriter
```

- `service.py` orchestrates load → validate → enabled-check → execute → write. Its constructor injects repository/validator/engine/writer_factory, so tests substitute mocks.
- `repository.py` reads the plan workbook into an `ExtractionSpec` — a frozen dataclass of plan/object/join/field/filter/group/aggregation row dicts.
- `validator.py` checks structure that does **not** depend on the source workbook (header shape, join graph, operator/target-type enums, aggregation-vs-field-mapping exclusivity). It cannot catch missing source sheets or bad source headers — those surface in the engine at `run` time. "Validate passes" does not imply "run passes."
- `engine.py` `execute()` runs two **mutually exclusive** paths: if aggregations are configured → `_aggregate` (read → join → filter → group/aggregate); otherwise `_select` (read → join → filter → field-map/derive). `_join`, `_filter`, `_coerce` are the shared pre-steps.
- `expression.py` `SafeExpressionEvaluator` is a hand-written whitelist AST walker (`_node` + `_call`). It allows only declared `alias.field` references, constants, a fixed BinOp/UnaryOp/Compare/BoolOp set, and the named functions defined in `_call`. Never replace with `eval`. To add an expression function, add a branch in `_call`.
- `template.py` generates the plan workbook with styled headers and `DataValidation` dropdowns — these dropdowns are a primary defense against invalid input from non-programmer users.

## Critical conventions

**The plan workbook is the configuration.** Required sheets: `抽取计划` / `数据对象` / `关联关系` / `字段映射` / `过滤条件`. Optional, for aggregation: `分组字段` / `聚合规则`. The Chinese sheet names **and** column headers (`任务ID`, `条件组`, `条件序号`, `源字段`, `目标类型`, …) are exact-match dict keys read by `repository._records` and consumed throughout `validator` and `engine`. They are load-bearing identifiers, not display labels — do not translate, rename, or "fix" them. Header constants live in `schema.py` (`*_HEADERS`).

**Field references use `别名.字段` (alias.field)** everywhere, e.g. `o.order_id`, `i.qty`. The regex `QUALIFIED_FIELD` in `schema.py` defines the shape; the alias prefix must match a declared `对象别名`.

**Filter semantics:** within a `条件组` (group), conditions AND; across groups, they OR (intra-group `mask &=`, inter-group `total |=`, in `engine._filter`). The operator whitelist lives in `SpecValidator.operators`; `IN` values are comma-separated.

**Aggregation and field-mapping are mutually exclusive** per task (validator enforces; `engine.execute` branches on whether aggregations exist). `count_all` takes no source field; every other aggregate function requires one.

**CLI stays thin** — argument parsing and exit codes only (0 success, 1 on validation failure or any caught exception). Subcommands: `template`, `validate`, `preview`, `run`; `allow_abbrev=False`, so flag abbreviations are rejected. New logic goes in `ExtractionService` or a component, not `cli.py`.

**Adding a new plan field or option** touches, in order: `schema.py` header constant → `template.py` (header + dropdown) → `repository.py` (if a new sheet) → `validator.py` (enum/range rules) → `engine.py` (execution) → `docs/reference/` → tests. Skipping the template/validator steps lets invalid values flow silently into the engine.

**Commits** follow Conventional Commits prefixes (`feat:`, `test:`, `docs:`, `chore:`) — see `git log`.

**Documentation updates track contracts, not commits.** `docs/reference/` is part of the extension chain — sync it in the same `feat:` that ships the capability. Touch `CLAUDE.md` only when a constraint-layer contract actually moves (workbook structure, filter/CLI semantics, execution-path exclusivity, test infrastructure), not on routine refactors, tests, or dependency bumps. If a stated anchor no longer matches the code, fix it on sight.

## Testing notes

- Standard library `unittest` only — there is **no pytest** configured.
- `uv run` is required so the `excelflow` console entrypoint is on PATH; a CLI test asserts `shutil.which("excelflow")` and fails without it.
- Tests build inputs inside `tempfile.TemporaryDirectory`; template-based fixtures are generated via `create_template`, not checked in.
- Curated reference plan workbooks live under `examples/` (one per docs tutorial, plus a full end-to-end sample) — known-good specs, separate from the generated test fixtures.
- When locking behavior, assert actual values (read writers back, compare frames), not just "no exception" — the suite prioritizes catching silent-wrong-data regressions over raw line coverage.

## Docs

Deeper write-ups already exist — read them before large changes: `docs/development/architecture.md`, `docs/development/development.md`, `docs/development/testing.md`, and the `docs/reference/*` pages. User-facing docs deploy to GitHub Pages on push to `main` via `.github/workflows/docs.yml`.
