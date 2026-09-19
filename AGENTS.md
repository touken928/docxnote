## Docxnote AGENTS 指南

**目标**：保持实现简洁、可靠、可维护，并在代码、测试、README 与 `docs/` 之间同步功能、行为与用法说明。

---

## 文档分工

- **`AGENTS.md`**：仓库内部维护说明，面向开发者 / coding agent，记录依赖、测试、设计约束与同步要求。
- **`README.md`**（英文入口）/**`docs/README_zh.md`**（中文入口）：项目概览、安装、最小快速开始；**不要写冗长 API 细节**。
- **`docs/`**：仓库内权威参考。
  - **`docs/API.md`** / **`docs/API_zh.md`**：完整 Python API、路径语义、批注行为、表格能力与高级用法。
- **`tests/README.md`**：测试目录结构、覆盖范围与常用命令；当测试布局变化时一起更新。

出现以下任一变化时，必须同步更新文档：

- 公共 API 签名或返回值变化；
- `keep_comments`、批注锚点、字符偏移、路径语义等行为变化；
- 测试目录结构或关键测试入口变化。

公共 API 或行为变更时：

- 更新实现与测试（至少覆盖相关目录）；
- 同步更新 `README.md` 与 `docs/README_zh.md` 的简要说明；
- 同步更新 **`docs/API.md`** 与 **`docs/API_zh.md`**（完整 Python API）；
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
  - 超链接等嵌套内容中的锚点。
