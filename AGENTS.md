## Docxnote AGENTS 指南

**目标**：保持实现简洁可靠，并在代码、`README.md`、`README_zh.md`、`docs/API.md`、`docs/API_zh.md`、`skills/SKILL.md`、`skills/library-usage.md` 之间**同步功能与用法**。

---

## 文档分工

- **`AGENTS.md`**：开发/维护本库时看的内部说明（依赖、测试、设计习惯）。
- **`README.md`** / **`README_zh.md`**：项目概览、安装、极简快速开始；**不写长 API**。详细接口与示例见 **`docs/API.md`**、**`docs/API_zh.md`**（[`docs/README.md`](docs/README.md) 索引）。
- **`skills/SKILL.md`**：Agent 入口，**概述 docxnote 能力与文档索引**，保持简短。
- **`skills/library-usage.md`**：面向 Agent 的 **Python 编码方案**（工作流、示例、常见误解、速查表）；方法签名与段落级说明与 **`docs/API*.md`** 对齐；后续 **CLI** 可另增独立 md（如 `skills/cli-usage.md`）。

公共 API 或行为变更时：

- 更新实现与测试；
- 同步更新 `README.md` 与 `README_zh.md` 的简要说明；
- 同步更新 **`docs/API.md`** 与 **`docs/API_zh.md`**（完整 API）；
- 同步更新 `skills/SKILL.md` 概述（若有整体能力变化）与 `skills/library-usage.md` 中的示例与推荐用法。

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
