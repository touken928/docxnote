# 测试

## 安装依赖

```bash
uv sync --dev
```

## 运行测试

```bash
# 运行所有测试
uv run pytest

# 详细输出
uv run pytest -v

# 运行特定文件
uv run pytest tests/test_xml_validity.py
```

## 测试覆盖

- **test_addressable_units.py** - 可寻址路径解析与段落遍历
- **test_cell_content.py** - 单元格文本提取
- **test_cli.py** - CLI 命令行接口
- **test_comment_conflicts.py** - 批注重叠与 Run 分割
- **test_comment_reading.py** - 批注读取与 keep_comments 行为
- **test_comments.py** - 批注添加与渲染
- **test_existing_comment_preservation.py** - 保留现有批注
- **test_nested_tables.py** - 嵌套表格遍历
- **test_paragraph_text.py** - 段落文本提取
- **test_paths.py** - 路径构建与解析
- **test_structure_comparison.py** - 文档结构与 python-docx 对比
- **test_table_shape.py** - 表格形状与合并单元格
- **test_thread_safety.py** - 线程安全
- **test_xml_validity.py** - XML 语法合法性

所有测试文档使用 python-docx 动态生成，不依赖外部文件。
