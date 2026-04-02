## Docxnote AGENTS 指南

**目标**：保持实现简洁可靠，并在代码、`README.md`（英文）、`README_zh.md`（中文）、`SKILL.md` 之间**同步功能与用法**。

---

## 文档分工

- **`AGENTS.md`**：开发/维护本库时看的内部说明（依赖、测试、设计习惯）。
- **`README.md`**：英文项目概览与核心 API；**`README_zh.md`**：中文对照，与英文同步维护。
- **`SKILL.md`**：给使用本库的对话型 / coding Agent，看如何在回答中**调用 `docxnote`**，内容尽量精简。

公共 API 或行为变更时：

- 更新实现与测试；
- 同步更新 `README.md` 与 `README_zh.md` 的简要说明；
- 同步更新 `SKILL.md` 的示例与推荐用法。

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
