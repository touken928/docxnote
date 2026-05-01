---
name: docxnote
description: >
  Lightweight Python library + `docxnote` CLI for DOCX Word comments: parse
  ZIP/XML, traverse paragraphs and tables via plain text, add/read comments by
  character ranges, merged/nested tables, stable address paths
  (resolve/iter_paragraphs, CLI operates on the same paths). Use for DOCX review
  automation or when the user mentions docxnote or Word comments. Structured API:
  repo docs/API.md; CLI: docs/CLI.md; agent patterns: library-usage.md and
  cli-usage.md.
---

# docxnote · 功能概述

**docxnote** 面向「按段落字符串」操控 Word 文档：在 DOCX 上读取正文、添加或读取 **批注**，并在渲染时写入 `comments.xml` 及必要关系。**不暴露 Run/XML**；表格（含合并单元格、嵌套表格）通过 `Table` / `Cell` 与段落递归访问。

## 核心能力

- **解析与写出**：`DocxDocument.parse(bytes)`、`doc.render()` → 新 DOCX 字节。
- **块级遍历**：`doc.blocks()` 得顶层 `Paragraph` / `Table`；单元格内再 `cell.blocks()`。
- **段落文本**：`paragraph.text`（含 `\n` / `\t`）；批注锚点为 **字符区间** `[start, end)`。
- **写批注**：`paragraph.comment(...)`；读批注：`paragraph.comments`、`doc.comments()`。
- **保留旧批注**：解析时 `keep_comments=True`，否则默认剥离文档内原有批注标记后再写。
- **可寻址单元**：`paragraph.path` / `Comment.path` 等；`doc.resolve(path)`、`doc.iter_paragraphs()`。
- **命令行（`docxnote` CLI）**：`list` / `show` / `comments` / `annotate`，与库共用 path 系统，全部子命令支持 `--json`。
- **线程**：同一 `DocxDocument` 实例内部用可重入锁保护，可多线程共用该实例。

## 文档索引

| 文档 | 用途 |
|------|------|
| `docs/API.md` / `docs/API_zh.md` | 仓库内 **完整 Python API**（方法、参数、路径、表格、高级用法）。 |
| `docs/CLI.md` / `docs/CLI_zh.md` | 仓库内 **完整 CLI 参考**（子命令、参数、退出码、JSON 结构）。 |
| **[library-usage.md](library-usage.md)** | Agent 侧：Python **安装、工作流骨架、模式、常见误解、速查表**。 |
| **[cli-usage.md](cli-usage.md)** | Agent 侧：**CLI** 工作流（locate → annotate → verify）、模式、坑点、速查表。 |
| **本文件（SKILL.md）** | 供 Agent 快速加载的**能力摘要**。 |

## 何时用 docxnote

适合：批量自动审阅、按规则打批注、表格内批注、在已有批注上追加（`keep_comments=True`）。不适合需要大量 Low-level OOXML 手工改版的场景（请仍通过本库的高层 API 完成）。
