# Python API 参考

[入门](README_zh.md) · [English](API.md) · [Shell 指南](SHELL_zh.md)

**docxnote** 的完整接口与用法说明。极简示例见 [README_zh.md](README_zh.md)。

通过 Python API 解析 DOCX 字节、读取或添加批注，再渲染为输出字节；文件读写由调用方负责。

## 目录

- [DocxDocument](#docxdocument)：解析、遍历、寻址、读取批注、渲染
- [Paragraph](#paragraph)：文本与添加批注
- [Comment](#comment-object)：字段、读取顺序与范围限制
- [Table](#table) / [Cell](#cell)：表格坐标和单元格内容
- [路径](#可寻址单元path)：地址格式与辅助函数
- [高级用法](#高级用法)：嵌套表格、多条批注、合并单元格
- [DocxShell / ShellResult](SHELL_zh.md#python-api)：命令接口的独立参考

## DocxDocument

DOCX 文档对象。

### parse

```python
DocxDocument.parse(docx_bytes, keep_comments=False)  # -> DocxDocument
```

解析 DOCX 字节并构建文档对象；`keep_comments` 只能作为关键字参数传入。

- **keep_comments**: 是否保留原有批注。默认 `False`（剥离原有批注）。如果你需要在“已有批注的 docx 上继续添加批注”并保留旧批注及其已有批注 XML 元数据，请传 `True`。

批注部件通过文档关系定位，支持相对路径和包内绝对路径，不要求文件名为
`word/comments.xml`。`keep_comments=True` 保留原部件路径、文档中的批注关系及
批注部件自身的关系，新批注 ID 从已有最大 ID 之后分配。`False` 会移除旧批注
部件及其关系部件；若追加新批注，则使用解析得到的部件路径，但不复用旧批注
内容及其自身的关系。

### blocks

```python
doc.blocks()
```

返回文档中的块级元素：

```python
(Paragraph | Table, ...)
```

顺序与 Word 文档一致。块级内容控件（`w:sdt`）按原位置展开其
`w:sdtContent` 中的段落和表格，支持嵌套控件。控件不增加路径段，`p:N`、
`t:N` 按该层展开后的块编号。`Cell.blocks()` 采用相同规则。

### iter_paragraphs

```python
doc.iter_paragraphs()  # Iterator[Paragraph]
```

按文档顺序遍历正文范围内的段落，包含表格、嵌套表格和块级内容控件，合并
单元格去重。每个段落带有 `path`。迭代开始时在文档锁内快照段落包装器；
文本范围详见 [text](#text)。

### resolve

```python
doc.resolve(path)  # Paragraph | Table | Cell | Comment
```

按[路径](#可寻址单元path)定位对象。非法路径结构抛出 `ValueError`，合法路径
指向不存在的对象时抛出 `LookupError`；读取不支持的批注范围会抛出
[UnsupportedCommentRangeError](#单段范围限制)。

<a id="文档级批注遍历"></a>

### comments

```python
comments = doc.comments()
for c in comments:
    # c.paragraph 为所属 Paragraph
    ...
```

`doc.comments()` 会遍历整个文档（包含表格及嵌套表格中的段落），按文档顺序返回所有批注。段落包装器每次读取 `comments` 都反映当前 XML 状态，因此旧包装器也能看到通过其他包装器新增的批注。其行为受 `keep_comments` 影响：

- `keep_comments=False`（默认）：仅暴露当前会话新增的批注，不暴露原始 DOCX 中旧批注。
- `keep_comments=True`：既保留又暴露原有批注，并允许在其基础上继续添加新批注；渲染时会继续保留这些已有批注。

若任一段落上存在跨段落或未闭合的批注范围，`doc.comments()` 会抛出 `UnsupportedCommentRangeError`。

### add_comment

```python
doc.add_comment(text, author="docxnote", date=None)  # int
```

仅分配批注 ID 并保存批注正文，不创建段落锚点。一般应用应使用
[`paragraph.comment(...)`](#paragraph-comment)，它会同时创建批注和锚点并返回 `Comment`。
未锚定的批注不会出现在 `doc.comments()` 中。

### render

```python
doc.render()
```

生成新的 DOCX 并返回 `bytes`。

所有批注在此阶段写入文档。

### 多线程

同一 `DocxDocument` 实例可在多线程中安全使用（内部使用可重入锁串行化访问）；不同实例可并行处理。多进程请各自 `parse` 得到独立实例。

## Paragraph

表示 Word 段落。

### text

```python
text = paragraph.text
```

返回段落文本，保留换行符（`\n`）和制表符（`\t`），包含超链接与行内
内容控件。run 内嵌文本框属于独立段落，其文本不计入宿主段落的文本和字符偏移。
文本框在渲染时保留，但不通过 `blocks()`、`iter_paragraphs()` 或批注范围视图
暴露；`DocxShell` 采用相同范围。文本读取、run 拆分及批注锚点读写统一使用
宿主段落的坐标。

<a id="paragraph-comment"></a>

### comment

```python
paragraph.comment(
    text,           # 批注内容
    start=0,        # 起始字符位置
    end=None,       # 结束字符位置（None 表示到末尾）
    author="docxnote",  # 批注作者
    date=None,          # 批注时间（建议带时区）；None 表示当前系统时间
)
```

`author` 与 `date` 只能作为关键字参数传入。没有文本 run 的段落会抛出 `ValueError`。

为段落文本范围添加批注。范围始终基于 `paragraph.text` 的字符区间，并遵循 Python 切片语义 `[start, end)`，包括负数和超出范围的端点；`end=None` 表示段落末尾。规范化后如果 end 早于 start，则变为锚定在规范化 start 的零长度范围。docxnote 会自动处理 Run 拆分与锚点放置，包括超链接等嵌套段落内容中的精确锚点。写入 `comments.xml` 的 `w:date` 为 UTC（`…Z`）。若传入不带时区的 `datetime`，按 UTC 解释。新建批注的 `date` 一定是具体的 `datetime` —— `date=None` 表示当前系统时间（带时区），不会是 `None`。

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

### comments

```python
paragraph.comments  # tuple[Comment, ...]
```

返回当前段落的批注快照。字段、顺序和不支持的范围见 [Comment](#comment-object)。

<a id="批注阅读comment"></a>

<a id="comment-object"></a>

## Comment

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

- `paragraph`：所属的 `Paragraph`；可通过 `comment.paragraph.text[start:end]` 取得被批注文本。
- `path`：批注的可寻址路径，形如 `"t:0/r:0/c:0/p:0#3"`；可直接传回 `doc.resolve(...)`
- `start` / `end`：基于 `paragraph.text` 的字符区间 `[start, end)`
- `text` / `author`：批注正文与作者
- `date`：`datetime | None`；从批注的 `w:date` 解析，保留源时区偏移（源值无时区时结果也无时区）。源属性缺失、空白或非法时为 `None`。`keep_comments=True` 渲染时原样保留源 `w:date` 属性（缺失仍缺失，非法字符串原样保留）。

批注按其 `commentRangeStart` 标记的 XML 文档顺序返回，嵌套批注外层在前。同起点和零长度范围也遵循该顺序，不受闭合顺序影响。

### 单段范围限制

批注范围仅支持单一段落：`commentRangeStart` 与 `commentRangeEnd` 必须位于同一段落内。高层范围视图 —— `paragraph.comments`、`doc.comments()`、`doc.resolve("p:0#N")` —— 在文档中存在跨段落或未闭合的批注范围时会抛出 `UnsupportedCommentRangeError`（`ValueError` 子类，可从 `docxnote` 导入）。`DocxDocument.parse` 与 `doc.render()` 不受影响：这类 XML 会被原样透传，因此仍可以用 `keep_comments=False` 剥离、或用 `keep_comments=True` 原样保留，而不检查范围。

## Table

表示 Word 表格。

### shape

```python
rows, cols = table.shape()
```

返回表格尺寸 `(行数, 列数)`。

### 单元格访问

```python
cell = table[row, col]
```

返回 `Cell` 对象。支持访问所有坐标，包括合并单元格覆盖的区域。

表格使用 `tblGrid` 定义的逻辑网格。行通过 `gridBefore` 或 `gridAfter`
省略的首尾坐标仍然可以访问，并返回合成的空 `Cell` 对象。合成单元格的
`blocks()` 返回 `()`，`bounds()` 返回一个单元格的左闭右开范围
`(row, col, row + 1, col + 1)`，`path` 使用正常的逻辑坐标路径（例如
`t:0/r:0/c:0`）。该路径可以传给 `doc.resolve()`，并解析到同一逻辑坐标的
合成单元格。

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

顺序与 Word 文档一致，块级内容控件按原位置展开，编号规则与 `doc.blocks()` 相同。

### bounds

```python
top, left, bottom, right = cell.bounds()
```

返回单元格边界 `(top, left, bottom, right)`，使用左闭右开区间 `[top, bottom)` 和 `[left, right)`。

对于未合并的单元格，返回 `(r, c, r+1, c+1)`。

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

编号从零开始，段落与表格分别计数。批注路径会按规范形式解析，因此路径段之间的空白或多余分隔符不会影响定位。合并单元格覆盖坐标会解析为原点单元格，其 `path` 使用原点坐标。遍历规则见 [iter_paragraphs](#iter_paragraphs)。

### 路径辅助函数

以下函数可从 `docxnote` 直接导入：

| 调用 | 返回值 |
| --- | --- |
| `build_segment("p", 2)` | `"p:2"` |
| `join_path("t:0", "r:1", "c:2")` | `"t:0/r:1/c:2"`，忽略空字符串 |
| `comment_path("p:0", 3)` | `"p:0#3"` |
| `parse_path("p:0#3")` | `([("p", 0)], 3)`；无批注后缀时第二项为 `None` |

编号必须非负。`parse_path` 解析语法，`doc.resolve` 检查路径结构及对象是否存在。

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
