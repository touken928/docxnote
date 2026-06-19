# Python API 参考

**docxnote** 的完整接口与用法说明。极简示例见 [README_zh.md](README_zh.md)。

---

## DocxDocument

DOCX 文档对象。

### parse

```python
DocxDocument.parse(docx_bytes, *, keep_comments=False)
```

解析 DOCX 并构建文档对象。

- **keep_comments**: 是否保留原有批注。默认 `False`（剥离原有批注）。如果你需要在“已有批注的 docx 上继续添加批注”并保留旧批注及其已有批注 XML 元数据，请传 `True`。

---

### blocks

```python
doc.blocks()
```

返回文档中的块级元素：

```python
(Paragraph | Table, ...)
```

顺序与 Word 文档一致。

---

### render

```python
doc.render()
```

生成新的 DOCX 并返回 `bytes`。

所有批注在此阶段写入文档。

### 多线程

同一 `DocxDocument` 实例可在多线程中安全使用（内部使用可重入锁串行化访问）；不同实例可并行处理。多进程请各自 `parse` 得到独立实例。

---

## 可寻址单元（path）

每一个 `Paragraph` / `Table` / `Cell` / `Comment` 都带有一个稳定的字符串 **path**，用于唯一定位它在文档中的位置：

- `p:N`：当前层级的第 N 个段落
- `t:N`：当前层级的第 N 个表格
- `t:N/r:R/c:C`：表格的 `(R, C)` 单元格（合并时指向原点）
- `t:N/r:R/c:C/p:M`：单元格内的段落（可递归进入嵌套表格）
- `<paragraph_path>#<id>`：段落上的一条批注，`<id>` 即 Word 的 `w:id`

```python
for block in doc.blocks():
    print(block.path, type(block).__name__)

# 通过 path 回溯对象
para = doc.resolve("p:0")
cell = doc.resolve("t:0/r:1/c:2")
comment = doc.resolve("t:0/r:1/c:2/p:0#3")
```

`doc.resolve(path)` 可以返回 `Paragraph` / `Table` / `Cell` / `Comment`。`doc.iter_paragraphs()` 会按文档顺序遍历所有段落（含表格与嵌套表格，合并单元格不会重复），每个段落都带有自己的 `path`。

---

## Paragraph

表示 Word 段落。

---

## 批注阅读（Comment）

```python
from docxnote import Comment
```

每个附着在段落上的批注都会以 `Comment` 对象的形式暴露在 `paragraph.comments` 中：

```python
for c in paragraph.comments:
    assert isinstance(c, Comment)
    print(c.path, c.start, c.end, c.text, c.author)
```

字段：

- `path`：批注的可寻址路径，形如 `"t:0/r:0/c:0/p:0#3"`；可直接传回 `doc.resolve(...)`
- `start` / `end`：基于 `paragraph.text` 的字符区间 `[start, end)`
- `text` / `author` / `date`：批注正文、作者、`w:date`（UTC）

### text

```python
text = paragraph.text
```

返回段落完整文本，保留换行符（`\n`）和制表符（`\t`）。

---

### comment

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

为段落文本范围添加批注。范围始终基于 `paragraph.text` 的字符区间，并遵循 Python 切片语义 `[start, end)`。docxnote 会自动处理 Run 拆分与锚点放置，包括超链接等嵌套段落内容中的精确锚点。写入 `comments.xml` 的 `w:date` 为 UTC（`…Z`）。若传入不带时区的 `datetime`，按 UTC 解释。

`paragraph.comment(...)` 会返回新建的 `Comment` 对象，它的 `path` 可以直接回传给 `doc.resolve(...)`。

**示例：**

```python
new = paragraph.comment("需要修改", start=3, end=8, author="张三")
print(new.path)                   # 例如 "p:0#0"
same = doc.resolve(new.path)
assert same.text == "需要修改"
```

docxnote 会自动处理：

- Run 分割
- 批注锚点
- comments.xml 写入
- 文档关系更新

---

## 文档级批注遍历

```python
comments = doc.comments()
for c in comments:
    # c.paragraph 为所属 Paragraph
    ...
```

`doc.comments()` 会遍历整个文档（包含表格及嵌套表格中的段落），按文档顺序返回所有批注。其行为受 `keep_comments` 影响：

- `keep_comments=False`（默认）：仅暴露当前会话新增的批注，不暴露原始 DOCX 中旧批注。
- `keep_comments=True`：既保留又暴露原有批注，并允许在其基础上继续添加新批注；渲染时会继续保留这些已有批注。

---

## Table

表示 Word 表格。

### shape

```python
rows, cols = table.shape()
```

返回表格尺寸 `(行数, 列数)`。

---

### 单元格访问

```python
cell = table[row, col]
```

返回 `Cell` 对象。支持访问所有坐标，包括合并单元格覆盖的区域。

---

## Cell

表示表格单元格。

### blocks

```python
cell.blocks()
```

返回单元格中的块级元素：

```python
(Paragraph | Table, ...)
```

顺序与 Word 文档一致。

---

### bounds

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
