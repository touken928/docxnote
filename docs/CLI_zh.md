# `docxnote` 命令行参考

安装后会注册 `docxnote` 命令，由 [`src/docxnote/cli.py`](../src/docxnote/cli.py) 提供。
所有读命令支持 `--json` 便于对接 LLM 流水线、脚本和表格；写命令必须显式指定输出文件。

```
docxnote --help
docxnote <子命令> --help
```

所有命令都基于 **可寻址路径**（`paragraph.path` / `Comment.path` / `Table.path` / `Cell.path`），
与 Python 侧 `DocxDocument.resolve(path)` 完全一致。路径语法见 [API_zh.md](API_zh.md#可寻址单元paths)。

不安装也可在仓库里直接跑：

```
uv run docxnote <子命令> ...
# 或
python -m docxnote.cli <子命令> ...
```

---

## `docxnote list FILE`

列出文档里所有段落路径（包括嵌套表格），自动对合并单元格去重。

| 参数 | 默认 | 说明 |
|------|------|------|
| `--text` | 关 | 纯文本模式下额外输出段落文字 |
| `--text-limit N` | `80` | 截断预览长度；`0` 表示不截断 |
| `--json` | 关 | 输出 `{path, text}` 的 JSON 数组 |
| `--keep-comments` / `--no-keep-comments` | 剥离 | 解析时是否保留原有 Word 批注 |

示例：

```
$ docxnote list report.docx --text --text-limit 40
p:0     季度报告
p:1     本季度营收同比增长 12%……
t:0/r:0/c:0/p:0 区域
t:0/r:0/c:1/p:0 营收
```

JSON 稳定、UTF-8 安全：

```
$ docxnote list report.docx --json | jq '.[0]'
{
  "path": "p:0",
  "text": "季度报告"
}
```

---

## `docxnote show FILE PATH`

把 `PATH` resolve 成 `Paragraph` / `Table` / `Cell` / `Comment` 并打印。
`show` 是读操作，默认 `--keep-comments` 开启。

JSON 结构：

```jsonc
// 段落
{"type": "paragraph", "path": "p:0", "text": "...", "comments": [ /* 见下 */ ]}

// 表格
{"type": "table", "path": "t:0", "shape": [rows, cols]}

// 单元格
{"type": "cell", "path": "t:0/r:1/c:2",
 "bounds": [top, left, bottom, right],
 "blocks": [{"kind": "paragraph", "path": "t:0/r:1/c:2/p:0"}, ...]}

// 批注
{"type": "comment", "path": "p:1#3", "paragraph": "p:1",
 "start": 0, "end": 5, "text": "...", "author": "...", "date": "ISO-8601"}
```

纯文本模式适合 `less` / `grep`。

---

## `docxnote comments FILE`

按文档顺序列出全部 Word 批注。默认 `--keep-comments` 开启。

纯文本每行：`<批注路径>\t<start>:<end>\t<author>\t<预览>`。

```
$ docxnote comments report.docx --json | jq '.[0]'
{
  "path": "p:1#0",
  "paragraph": "p:1",
  "start": 0,
  "end": 7,
  "text": "请注明出处。",
  "author": "reviewer",
  "date": "2026-04-22T09:12:00+00:00"
}
```

---

## `docxnote annotate INPUT OUTPUT`

在 `INPUT` 上新增一条或多条批注，写入 `OUTPUT`。**永不原地覆盖**。

### 单条

```
docxnote annotate in.docx out.docx \
  --path "p:1" \
  --text "请修改开头。" \
  --start 0 --end 5 \
  --author reviewer
```

`--path` 与 `--text` 必须同时给出。`--start` 默认 `0`，`--end` 默认 *段末*。两者都基于目标段落 `text` 的字符偏移，遵循 `[start, end)` 语义。

### 批量（`--spec` JSON）

`--spec ops.json` 读取数组，每项是一条操作。可与 `--path/--text` 组合，单条会追加到末尾。

```json
[
  {"path": "p:0", "text": "标题缺少年份", "author": "editor"},
  {"path": "t:0/r:1/c:2/p:0", "text": "核对数字", "start": 0, "end": 3}
]
```

### 行为

- `--keep-comments` / `--no-keep-comments`：控制**输入文档**原有批注是否在写入前保留；
  默认**剥离**（与 `DocxDocument.parse` 默认一致）。
- 使用 `--keep-comments` 时，输出文件会继续保留原有批注以及它们已有的批注 XML 元数据。
- 任一操作的 `path` 没解析到段落，则以状态码 `2` 退出，**不写任何文件**。
- `--json` 时 stdout 为 `{"output": "<路径>", "added": [<批注对象>, ...]}`。

---

## 退出码

| 码 | 含义 |
|----|------|
| `0` | 成功 |
| `2` | 用户错误：参数错、文件不存在、路径无法解析、spec 非法等 |

---

## 典型模式

**定位 → 批注 循环**（shell）：

```sh
# 1. 浏览段落
docxnote list input.docx --text --json > index.json

# 2. 挑选目标并生成 ops.json（jq / LLM / 脚本皆可）
jq '[.[] | select(.text | test("TODO")) | {path, text: "处理 TODO"}]' \
   index.json > ops.json

# 3. 应用批注
docxnote annotate input.docx annotated.docx --spec ops.json --json
```

**回环自检**：

```
docxnote annotate in.docx tmp.docx --path p:0 --text "hi"
docxnote comments tmp.docx --json
```
