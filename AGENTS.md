## Docxnote AGENTS 指南

**目标**：保持实现简洁可靠，并在代码、`README.md`、`README_zh.md`、`docs/API.md`、`docs/API_zh.md`、`docs/CLI.md`、`docs/CLI_zh.md`、`skills/SKILL.md`、`skills/library-usage.md`、`skills/cli-usage.md` 之间**同步功能与用法**。

---

## 文档分工

- **`AGENTS.md`**：开发/维护本库时看的内部说明（依赖、测试、设计习惯）。
- **`README.md`** / **`README_zh.md`**：项目概览、安装、极简快速开始、CLI 一屏介绍；**不写长 API / CLI 细节**。
- **`docs/`**：仓库内权威参考。
  - **`docs/API.md`** / **`docs/API_zh.md`**：完整 Python API。
  - **`docs/CLI.md`** / **`docs/CLI_zh.md`**：完整 CLI 参考（子命令、选项、退出码、JSON 结构）。
  - **`docs/README.md`**：文档目录索引。
- **`skills/SKILL.md`**：Agent 入口，**概述 docxnote 能力与文档索引**，保持简短。
- **`skills/library-usage.md`**：面向 Agent 的 **Python 编码方案**（工作流、示例、常见误解、速查表），与 `docs/API*.md` 对齐。
- **`skills/cli-usage.md`**：面向 Agent 的 **CLI 使用方案**（locate → annotate → verify、模式、坑点、速查），与 `docs/CLI*.md` 对齐。

公共 API 或行为变更时：

- 更新实现与测试（含 `tests/test_cli.py`）；
- 同步更新 `README.md` 与 `README_zh.md` 的简要说明；
- 同步更新 **`docs/API.md`** 与 **`docs/API_zh.md`**（完整 Python API）；
- 涉及命令行时同步更新 **`docs/CLI.md`** 与 **`docs/CLI_zh.md`**；
- 同步更新 `skills/SKILL.md` 概述（若有整体能力变化）、`skills/library-usage.md` 与（CLI 变更时）`skills/cli-usage.md` 的示例与推荐用法。

---

## 开发与提交要点

- 依赖用 `uv` 管理（见 `pyproject.toml` / `uv.lock`）。
- 使用 Ruff/ruff-format（通过 pre-commit）统一风格和静态检查。
- 新功能或行为修改必须有测试覆盖。
- 提交前跑：
  - `uv run pre-commit run --all-files`
  - `uv run pytest`
- 发版：更新 `pyproject.toml` 的 `version`，打标签 `v*` 推送；PyPI 使用 Trusted Publisher，与 `.github/workflows/publish.yml` 一致。

---

## 实现习惯

- 同一 `DocxDocument` 使用 `threading.RLock` 保护共享状态；新增访问 XML / 批注列表的路径须持锁。
- 对外只暴露文本视图与高层对象：
 - `DocxDocument` / `Paragraph` / `Table` / `Cell`
 - 使用 `paragraph.text` 和 `paragraph.comment(start, end, ...)` 等简单接口。
- Word Run / XML 细节全部封装在内部模块，不放进公共 API。

## CLI 约定

- 入口：`src/docxnote/cli.py` 的 `main(argv=None)`；通过 `[project.scripts] docxnote = "docxnote.cli:main"` 暴露。
- 只使用 stdlib（`argparse` + `json`），不引入新依赖。
- 读命令 (`list` / `show` / `comments`) 必须支持 `--json`，输出 JSON schema 需稳定；新增字段可以，**不要重命名或删除既有字段**（测试里有断言，改动须同步更新）。
- 写命令 (`annotate`) 必须：1) 采用显式 `INPUT OUTPUT` 参数，不支持原地覆盖；2) 任一 op 失败则整体失败、不写出文件；3) 失败时以退出码 `2` 退出并把错误写到 stderr。
- 新加子命令或选项时：同步 `tests/test_cli.py`、`docs/CLI*.md`、`skills/cli-usage.md`、`README*.md` CLI 段落。
