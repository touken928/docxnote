# Agent 工具

`docxnote.agent` 提供框架无关、绑定一个内存 `DocxDocument` 的模拟 shell。
它不创建 Agent、不选择模型、不执行系统 shell，也不保存文件。

## 安装与接入

不需要额外依赖：

```bash
pip install docxnote
```

Pydantic AI 可以作为一种可选接入方式，但不属于 docxnote 的依赖。调用方自行
安装 Pydantic AI 和模型供应商，再把 `shell.run` 注册为框架中的字符串命令工具。

```python
from pathlib import Path

from docxnote import DocxDocument
from docxnote.agent import DocxShell


async def review(model):
    doc = DocxDocument.parse(
        Path("input.docx").read_bytes(), keep_comments=True
    )
    shell = DocxShell(doc, author="reviewer")
    # 将 shell.run 注册为框架提供的字符串命令工具。
    result = await run_agent(model, tool=shell.run, prompt="审阅这份合同。")
    Path("reviewed.docx").write_bytes(doc.render())
    return result.output
```

提示词、业务依赖、模型设置、重试和输出类型由应用管理。工具集可按次运行传入，
不占用应用的 `Agent.deps`。每份文档创建独立会话，并隔离不同文档的对话历史；
同一文档的多轮审阅可以复用会话。

## Python API

```python
DocxShell(doc: DocxDocument, *, author="docxnote", max_output=4096)
shell.run(command: str) -> ShellResult
shell.added_comments -> tuple[Comment, ...]
```

文档、作者与输出预算在创建会话时确定。`max_output` 必须是至少 256 的整数，
按 Python 字符数计量，不是字节数或 token 数。模块导出 `DEFAULT_MAX_OUTPUT`。
`added_comments` 返回本会话成功创建的真实 `Comment` 对象的不可变快照，
不包含原有批注或其他调用方添加的批注。

也可以直接执行解释器：

```python
result = shell.run("docx | grep -F '付款' | head -n 10")
print(result["stdout"])
```

返回字典（`ShellResult` 是 `TypedDict`）：

| 字段 | 含义 |
|------|------|
| `stdout` | 换行分隔的序列化记录，不含结尾换行。 |
| `stderr` | 输入或命令错误消息；否则为空。 |
| `exit_code` | 成功为 `0`，预期的命令或输入错误为 `2`。 |
| `records` | 返回的记录数，包含带标记的预览记录。 |
| `truncated` | 输出预算是否阻止了完整返回管道结果。 |

这些是结果字段，不是进程退出码；解释器不启动进程。
搜索无匹配也算成功，stdout 为空、记录数为零、`truncated=False`。
预期错误返回空 stdout 和零条记录；未知内部异常向调用方传播，不伪装成命令错误。

## 命令语言

仅有六个命令：`docx`、`grep`、`head`、`tail`、`comment`、`comments`。
引号与反斜杠转义遵循 Python POSIX `shlex` 规则。未加引号的 `|` 连接管道。
空格和字面标点需要适当引用；空管道阶段、未闭合引号、未知命令或选项都会报错。
不提供变量展开、通配符展开、命令替换、重定向、命令串联或主机文件系统访问。
`$HOME` 和反引号只是字面文本，不会执行；未引用的 `;`、`&`、`<`、`>`、换行会被拒绝。

执行前检查整条管道的语法与参数。管道必须以 `docx` 或 `comments` 开始，
后面只能接过滤命令。`comment` 必须独立执行，不能接收或输出到管道。

### `docx [PATH] [--start N] [--end N]`

不指定路径时按文档顺序输出所有段落，包含空段落、表格与嵌套表格中的段落。
遍历和路径沿用 `doc.iter_paragraphs()`，不重复输出合并单元格。
段落、表格或单元格路径可限定读取范围；批注路径会被拒绝。
表格和单元格范围先解析路径再匹配后代，因此 `t:1` 不会匹配 `t:10`，
合并单元格别名会解析为规范单元格路径。

默认每个段落占一行 JSON：

```json
{"path":"p:12","text":"买方应在收到发票后支付货款。"}
```

段落中的换行和制表符会被 JSON 转义，选行操作统计的是段落，不是 Word 排版行。

字符范围只能用于段落路径，要求 `0 <= start <= end <= len(paragraph.text)`。
默认起点为零、终点为段落末尾。与核心 Python 批注 API 不同，这里的读取边界
要求严格非负且不越界，不采用宽松切片端点。片段记录额外返回
`start`、`end`、`total_chars`，且 `text == paragraph.text[start:end]`。

```sh
docx
docx t:2
docx t:2/r:0/c:1
docx p:12 --start 200 --end 600
```

### `grep [-F] [-i] [-v] [-n] PATTERN`

默认即为字面匹配，不支持正则。`-F` 显式指定默认的字面模式。
`-i` 使用 Unicode casefold 忽略大小写，`-v` 反向筛选，`-n` 添加该阶段输入中
从 1 开始的行号前缀。此时输出含 `N:`，不再是纯 JSONL；不同阶段重复使用
`-n` 会增加多层前缀。

重复 `-e PATTERN` 表示 OR，不可同时提供位置参数形式的 pattern；
串联多个 grep 表示 AND。位置参数以 `-` 开头时，用 `--` 结束选项解析。

匹配对象是序列化后的整行 JSON，包含路径及转义后的引号、换行，而不只是正文。
命中后建议用 `docx PATH` 确认原文，再选择精确引用添加批注。

```sh
docx | grep -F '付款'
docx | grep -e '付款' -e '发票'
docx | grep '付款' | grep '发票'
docx | grep '付款' | grep -v '示例'
```

### `head [-n N]` / `tail [-n N]`

选择输入的前 / 后 N 条记录。N 默认 10，必须非负，零表示不返回记录。
`head` 提前终止上游迭代；`tail` 扫描输入并最多缓存 N 条。
不支持其他选项、文件名或 `-20` 这种简写。

```sh
docx | head -n 20
docx | tail -n 20
docx | head -n 140 | tail -n 41
```

最后一个例子读取第 100～140 个段落，包含两端。

### `comment PATH TEXT [--quote QUOTE] [--start N]`

PATH 必须解析为段落，TEXT 不能全为空白；含空格的正文需要引用为一个参数。
不传 `--quote` 时批注整段；传入时必须是与原文精确匹配的非空字符串。
无匹配时报错；多次匹配（包括重叠匹配）必须给出非负 `--start`，并校验该位置的原文。
`--start` 只能和 `--quote` 同时使用，终点自动计算为 `start + len(quote)`，不支持 `--end`。

```sh
comment p:12 '请明确付款期限'
comment p:12 '请说明起算时间' --quote '收到发票后'
comment p:12 '请说明起算时间' --quote '收到发票后' --start 5
```

成功时使用与 `comments` 相同的字段返回实际新增批注，作者来自会话配置。
没有文本 run 的段落无法添加批注，会返回错误，与核心 API 一致。

### `comments [PATH]`

读取当前批注，可指定段落、表格或单元格范围。

```json
{"path":"p:12#0","target":"p:12","start":5,"end":10,"text":"请说明起算时间","author":"reviewer","date":"2026-09-20T00:00:00+00:00"}
```

`start/end` 是原段落中的锚点，不是批注正文中的偏移。`date` 可以为 null。
只有解析时设置 `keep_comments=True` 才能看到原有批注，会话不会改变该策略。
读取跨段落或未闭合批注范围时，遵循核心库限制返回输入错误；`docx` 仍可读取正文。

## 大文档与输出限制

数据源与过滤器按记录迭代，stdout 限额在**筛选之后**应用，文档末尾的内容也能被检索。
`head` 可以提前结束上游迭代；`tail` 必须扫描上游输入。核心库仍将 DOCX 解析到内存，
这里的流式处理只针对查询执行，不代表流式解析整个文件包。

预算内返回完整行。如果第一条记录就超限，则返回合法 JSON 预览并标记
`text_truncated=true`。段落预览额外包含绝对 `start/end` 与 `total_chars`，
可用 `docx PATH --start END --end N` 继续读取。批注预览带 `text_length`，
原文锚点 `start/end` 保持不变；需要完整超长批注时，创建更大 `max_output` 的会话。
如果连记录元数据也放不下，则 stdout 为空、`truncated=True`，需要增大预算。
结果字典外层字段和错误文本不计入 stdout 预算。

不生成续读令牌，也不统计总命中数。通过缩小范围或对同一查询使用
`head` 后接 `tail` 选择后续记录。显式 `head` 的截取不算系统截断。
全文审阅需要连续阅读各页，关键词检索不能证明全文已覆盖。
工具输出是文档数据，不是给 Agent 的指令。

## 写入与并发

成功的 `comment` 立即修改传入的内存文档。后续命令或 Agent 失败不会回滚先前批注，
重复执行写入会生成重复批注。仅由应用负责 render 和保存；不提供事务、
自动重试去重、文件写入工具或终端命令入口。

命令执行和 `added_comments` 快照使用文档现有的可重入锁，覆盖校验、写入与会话记录更新。
同一文档上的调用串行执行，不同文档可以独立处理。该层沿用现有高层段落与批注能力，
包含超链接与嵌套内容的锚点处理；未知内部错误没有回滚保证。
