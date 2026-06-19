## Docxnote AGENTS 指南

**目标**：保持实现简洁、可靠、可维护，并在代码、测试、README、`docs/` 与 `skills/docxnote/` 之间**同步功能、行为与用法说明**。

---

## 文档分工

- **`AGENTS.md`**：仓库内部维护说明，面向开发者 / coding agent，记录依赖、测试、设计约束与同步要求。
- **`README.md`**（英文入口）/**`docs/README_zh.md`**（中文入口）：项目概览、安装、最小快速开始、CLI 简介；**不要写冗长 API / CLI 细节**。
- **`docs/`**：仓库内权威参考。
  - **`docs/API.md`** / **`docs/API_zh.md`**：完整 Python API、路径语义、批注行为、表格能力与高级用法。
  - **`docs/CLI.md`** / **`docs/CLI_zh.md`**：完整 CLI 参考（子命令、参数、退出码、JSON 结构、行为约束）。
- **`skills/docxnote/SKILL.md`**：Agent 入口，只保留能力概述与文档索引，保持简短。
- **`skills/docxnote/library-usage.md`**：面向 Agent 的 Python 使用指南（工作流、示例、推荐模式、常见误区、速查表），与 `docs/API*.md` 对齐。
- **`skills/docxnote/cli-usage.md`**：面向 Agent 的 CLI 使用指南（locate → annotate → verify、批处理模式、坑点、速查表），与 `docs/CLI*.md` 对齐。
- **`tests/README.md`**：测试目录结构、覆盖范围与常用命令；当测试布局变化时一起更新。

出现以下任一变化时，必须同步更新文档：

- 公共 API 签名或返回值变化；
- `keep_comments`、批注锚点、字符偏移、路径语义等行为变化；
- CLI 子命令、参数、JSON 输出、退出码变化；
- `skills/docxnote/` 中的推荐工作流、示例或约束变化；
- 测试目录结构或关键测试入口变化。

公共 API 或行为变更时：

- 更新实现与测试（至少覆盖相关目录；涉及 CLI 时必须覆盖 `tests/cli/test_cli.py`）；
- 同步更新 `README.md` 与 `docs/README_zh.md` 的简要说明；
- 同步更新 **`docs/API.md`** 与 **`docs/API_zh.md`**（完整 Python API）；
- 涉及命令行时同步更新 **`docs/CLI.md`** 与 **`docs/CLI_zh.md`**；
- 同步更新 `skills/docxnote/SKILL.md` 概述（若整体能力变化）、`skills/docxnote/library-usage.md` 与（CLI 变更时）`skills/docxnote/cli-usage.md`；
- 测试布局或运行方式变化时，同步更新 `tests/README.md`。

---

## 开发与提交要点

- 依赖用 `uv` 管理（见 `pyproject.toml` / `uv.lock`）。
- 使用 Ruff/ruff-format（通过 pre-commit）统一风格和静态检查。
- 新功能或行为修改必须有测试覆盖。
- 优先保持补丁小而清晰；若行为变化较大，先补失败测试，再修实现。
- 提交前跑：
  - `uv run pre-commit run --all-files`
  - `uv run pytest`
- 发版：更新 `pyproject.toml` 的 `version`，打标签 `v*` 推送；PyPI 使用 Trusted Publisher，与 `.github/workflows/publish.yml` 一致。

---

## 实现习惯

- 同一 `DocxDocument` 使用 `threading.RLock` 保护共享状态；新增访问 XML / 批注列表的路径须持锁。
- 对外只暴露文本视图与高层对象：`DocxDocument` / `Paragraph` / `Table` / `Cell` / `Comment`。
- 优先围绕 `paragraph.text`、`paragraph.comment(start, end, ...)`、`paragraph.comments`、`doc.comments()` 等高层接口设计行为。
- Word run / XML 细节全部封装在内部模块，不放进公共 API。
- 对批注范围的任何修改，都要坚持 `paragraph.text` 上的字符偏移语义 `[start, end)`；不要把内部 run 边界泄漏到公共接口。
- 对 `keep_comments=True` 的修改，要特别注意：不仅要保留可见批注，还要避免破坏已有 `comments.xml` 元数据与包关系。
- 涉及超链接、嵌套容器、合并单元格、嵌套表格的改动，优先补回归测试，因为这些地方最容易出现锚点或遍历偏移错误。

## 测试约定

- 当前测试按领域分组：
  - `tests/cli/`：CLI 行为与 JSON 输出
  - `tests/comments/`：批注写入、读取、冲突范围、旧批注保留
  - `tests/document/`：路径、结构、线程安全
  - `tests/tables/`：表格、单元格、嵌套表格、合并单元格
  - `tests/text/`：段落文本提取
  - `tests/xml/`：XML / package 合法性
- 新测试尽量放到最贴近语义的目录，不要继续堆在 `tests/` 根目录。
- 行为 bug 修复优先补“能稳定复现问题”的回归测试，尤其是：
  - 局部批注字符范围；
  - `keep_comments` 保留 / 剥离行为；
  - `comments.xml` / rels / `[Content_Types].xml` 包完整性；
  - 超链接等嵌套内容中的锚点；
  - CLI 的退出码与 JSON 结构稳定性。

## CLI 约定

- 入口：`src/docxnote/cli.py` 的 `main(argv=None)`；通过 `[project.scripts] docxnote = "docxnote.cli:main"` 暴露。
- 只使用 stdlib（`argparse` + `json`），不引入新依赖。
- 读命令 (`list` / `show` / `comments`) 必须支持 `--json`，输出 JSON schema 需稳定；新增字段可以，**不要重命名或删除既有字段**（测试里有断言，改动须同步更新）。
- 写命令 (`annotate`) 必须：1) 采用显式 `INPUT OUTPUT` 参数，不支持原地覆盖；2) 任一 op 失败则整体失败、不写出文件；3) 失败时以退出码 `2` 退出并把错误写到 stderr。
- `annotate` 的 `start` / `end` 语义始终基于目标段落 `text` 的字符区间 `[start, end)`；如果实现变了，必须同步更新文档与测试。
- 新加子命令或选项时：同步 `tests/cli/test_cli.py`、`docs/CLI*.md`、`skills/docxnote/cli-usage.md`、`README*.md` 的 CLI 段落。
