<h1 align="center">Docxnote</h1>

<p align="center">
  <strong>轻量级 DOCX 批注引擎：在段落纯文本上添加与读取 Word 批注，仅依赖 <code>lxml</code>。</strong>
</p>

<p align="center">
  <a href="https://www.python.org/downloads/"><img src="https://img.shields.io/badge/python-3.12+-blue.svg?style=for-the-badge&logo=python" alt="Python 3.12+"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-Apache%202.0-blue.svg?style=for-the-badge" alt="License: Apache 2.0"></a>
  <a href="https://pypi.org/project/docxnote/"><img src="https://img.shields.io/pypi/v/docxnote.svg?style=for-the-badge&logo=pypi&logoColor=white&label=pypi" alt="PyPI version"></a>
  <a href="https://github.com/touken928/docxnote/stargazers"><img src="https://img.shields.io/github/stars/touken928/docxnote?style=for-the-badge&color=yellow&logo=github" alt="GitHub stars"></a>
</p>

<p align="center">
  <a href="README.md">English</a> &middot; 简体中文
</p>

---

## 概览

**docxnote** 用于自动化 **Word 批注**：遍历 `Paragraph` / `Table` / `Cell`，用 `paragraph.comment(...)` 添加批注，并可通过 `paragraph.comments` 与 `doc.comments()` 读取批注。

**仓库：** [touken928/docxnote](https://github.com/touken928/docxnote)

---

## 安装

```
pip install docxnote
```

使用 [uv](https://github.com/astral-sh/uv)：

```
uv add docxnote
```

---

## 快速开始

```python
from docxnote import DocxDocument, Paragraph, Table

# 读取文档
with open("document.docx", "rb") as f:
    # 默认不保留原有批注（会清空）
    doc = DocxDocument.parse(f.read())

    # 如需保留原有批注并继续添加：
    # doc = DocxDocument.parse(f.read(), keep_comments=True)

# 遍历文档块
for block in doc.blocks():
    if isinstance(block, Paragraph):
        # 为段落添加批注
        if block.text:
            block.comment("请检查表述", end=5, author="reviewer")

    elif isinstance(block, Table):
        # 处理表格
        rows, cols = block.shape()
        for r in range(rows):
            for c in range(cols):
                cell = block[r, c]
                # 为单元格内容添加批注
                for inner in cell.blocks():
                    if isinstance(inner, Paragraph) and inner.text:
                        inner.comment("需复核", end=3, author="reviewer")

# 生成新文档
output = doc.render()
with open("output.docx", "wb") as f:
    f.write(output)
```

---

## API

### DocxDocument

DOCX 文档对象。

#### parse

```python
DocxDocument.parse(docx_bytes, *, keep_comments=False)
```

解析 DOCX 并构建文档对象。

- **keep_comments**: 是否保留原有批注。默认 `False`（清空所有原有批注）。如果你需要在“已有批注的 docx 上继续添加批注”并保留旧批注，请传 `True`。

---

#### blocks

```python
doc.blocks()
```

返回文档中的块级元素：

```python
(Paragraph | Table, ...)
```

顺序与 Word 文档一致。

---

#### render

```python
doc.render()
```

生成新的 DOCX 并返回 `bytes`。

所有批注在此阶段写入文档。

#### 多线程

同一 `DocxDocument` 实例可在多线程中安全使用（内部使用可重入锁串行化访问）；不同实例可并行处理。多进程请各自 `parse` 得到独立实例。

---

### Paragraph

表示 Word 段落。

---

### 批注阅读（Comment）

```python
from docxnote import Comment
```

每个附着在段落上的批注都会以 `Comment` 对象的形式暴露在 `paragraph.comments` 中：

```python
for c in paragraph.comments:
    assert isinstance(c, Comment)
    print(c.start, c.end, c.text, c.author)
```

其中：

- `start` / `end` 是基于 `paragraph.text` 的字符区间，遵循 Python 切片约定 \([start, end)\)。
- `Comment.date` 对应 `comments.xml` 中的 `w:date`（UTC 时间）。

#### text

```python
text = paragraph.text
```

返回段落完整文本，保留换行符（`\n`）和制表符（`\t`）。

---

#### comment

```python
paragraph.comment(
    text,           # 批注内容
    start=0,        # 起始字符位置
    end=None,       # 结束字符位置（None 表示到末尾）
    *,
    author="docxnote",  # 批注作者
    date=None,          # 批注时间（建议带时区）；None 表示当前系统时间
)
```

为段落文本范围添加批注。写入 `comments.xml` 的 `w:date` 为 UTC（`…Z`）。若传入不带时区的 `datetime`，按 UTC 解释。

**示例：**

```python
paragraph.comment("需要修改", start=3, end=8, author="张三")
```

docxnote 会自动处理：

- Run 分割
- 批注锚点
- comments.xml 写入
- 文档关系更新

---

### 文档级批注遍历

```python
comments = doc.comments()
for c in comments:
    # c.paragraph 为所属 Paragraph
    ...
```

`doc.comments()` 会遍历整个文档（包含表格及嵌套表格中的段落），按文档顺序返回所有批注。其行为受 `keep_comments` 影响：

- `keep_comments=False`（默认）：仅暴露当前会话新增的批注，不暴露原始 DOCX 中旧批注。
- `keep_comments=True`：既保留又暴露原有批注，并允许在其基础上继续添加新批注。

---

### Table

表示 Word 表格。

#### shape

```python
rows, cols = table.shape()
```

返回表格尺寸 `(行数, 列数)`。

---

#### 单元格访问

```python
cell = table[row, col]
```

返回 `Cell` 对象。支持访问所有坐标，包括合并单元格覆盖的区域。

---

### Cell

表示表格单元格。

#### blocks

```python
cell.blocks()
```

返回单元格中的块级元素：

```python
(Paragraph | Table, ...)
```

顺序与 Word 文档一致。

---

#### bounds

```python
top, left, bottom, right = cell.bounds()
```

返回单元格边界 `(top, left, bottom, right)`，使用左闭右开区间 `[top, bottom)` 和 `[left, right)`。

对于未合并的单元格，返回 `(r, c, r+1, c+1)`。

---

## 高级用法

### 处理嵌套表格

```python
for block in doc.blocks():
    if isinstance(block, Table):
        rows, cols = block.shape()
        for r in range(rows):
            for c in range(cols):
                cell = block[r, c]
                # 遍历单元格内的块（可能包含嵌套表格）
                for inner_block in cell.blocks():
                    if isinstance(inner_block, Table):
                        # 处理嵌套表格
                        inner_rows, inner_cols = inner_block.shape()
                        # ...
```

### 多个批注

```python
# 为同一段落的不同位置添加多个批注
paragraph.comment("批注1", start=0, end=5, author="张三")
paragraph.comment("批注2", start=10, end=15, author="李四")
paragraph.comment("批注3", start=20, end=25, author="王五")
```

### 处理合并单元格

```python
table = [b for b in doc.blocks() if isinstance(b, Table)][0]

# 访问合并单元格
cell = table[0, 0]
top, left, bottom, right = cell.bounds()

# 如果单元格跨越多行或多列
if bottom - top > 1 or right - left > 1:
    print(f"合并单元格：跨越 {bottom-top} 行，{right-left} 列")
```

---

## 测试

所有测试文档使用 python-docx 动态生成，不依赖外部文件，详见 [tests/README.md](tests/README.md)。

---

## SKILL

本仓库附带 [`SKILL.md`](SKILL.md)，用于指导对话型 / coding Agent 正确调用 `docxnote`。在项目根目录执行（默认分支为 `main`，若不同请改 URL 中的分支名）。Windows PowerShell 下若 `curl` 被解析为 `Invoke-WebRequest`，请改用 `curl.exe`。

**Cursor**

```bash
mkdir -p .cursor/docxnote
curl -fsSL -o .cursor/docxnote/SKILL.md https://raw.githubusercontent.com/touken928/docxnote/main/SKILL.md
```

**Claude Code**

```bash
mkdir -p .claude/docxnote
curl -fsSL -o .claude/docxnote/SKILL.md https://raw.githubusercontent.com/touken928/docxnote/main/SKILL.md
```

在对话环境中使用本库时，让 Agent 优先参考该文件中的安装方式、推荐代码骨架与注意事项。

---

## 许可证

本项目采用 Apache License 2.0 许可证。详见仓库根目录的 `LICENSE`。
