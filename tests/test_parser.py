# tests/test_parser.py
import pytest
from md2docx.parser import (
    HeadingElement, ParagraphElement, MermaidBlockElement,
    FormulaBlockElement, TableElement, CodeBlockElement,
    ImageElement, PageBreakElement, AbstractElement,
    MarkdownParser, parse_markdown
)


class TestElementTypes:
    def test_heading_element(self):
        elem = HeadingElement(level=2, text="第 1 章 绪论")
        assert elem.level == 2
        assert elem.text == "第 1 章 绪论"

    def test_paragraph_element(self):
        elem = ParagraphElement(text="测试段落", has_inline_math=False)
        assert elem.text == "测试段落"
        assert elem.has_inline_math == False

    def test_mermaid_block_element(self):
        elem = MermaidBlockElement(
            code="flowchart TB\n    A --> B",
            caption="图 1-1 系统架构"
        )
        assert "flowchart" in elem.code
        assert elem.caption == "图 1-1 系统架构"

    def test_formula_block_element(self):
        elem = FormulaBlockElement(latex=r"\frac{a}{b}", tag="1-1")
        assert elem.latex == r"\frac{a}{b}"
        assert elem.tag == "1-1"

    def test_table_element(self):
        elem = TableElement(
            headers=["参数", "值"],
            rows=[["A", "100"], ["B", "200"]],
            caption="表 1-1 参数对比"
        )
        assert len(elem.headers) == 2
        assert len(elem.rows) == 2

    def test_code_block_element(self):
        elem = CodeBlockElement(code="print('hello')", language="python")
        assert elem.language == "python"

    def test_image_element(self):
        elem = ImageElement(path="/path/to/image.png", caption="图 1-2 示例")
        assert elem.path == "/path/to/image.png"

    def test_page_break_element(self):
        elem = PageBreakElement()
        assert isinstance(elem, PageBreakElement)

    def test_abstract_element(self):
        elem = AbstractElement(
            lang="zh",
            body="这是摘要内容...",
            keywords="关键词1；关键词2"
        )
        assert elem.lang == "zh"


class TestMarkdownParser:
    def test_parse_heading(self):
        md = "## 第 1 章 绪论\n"
        elements = parse_markdown(md)
        assert len(elements) == 1
        assert isinstance(elements[0], HeadingElement)
        assert elements[0].level == 2
        assert elements[0].text == "第 1 章 绪论"

    def test_parse_paragraph(self):
        md = "这是一段普通文本。\n"
        elements = parse_markdown(md)
        assert len(elements) == 1
        assert isinstance(elements[0], ParagraphElement)

    def test_parse_paragraph_with_inline_math(self):
        md = "公式 $E=mc^2$ 的意义。\n"
        elements = parse_markdown(md)
        assert len(elements) == 1
        assert elements[0].has_inline_math == True

    def test_parse_mermaid_block(self):
        md = """**图 1-1 系统架构图**

```mermaid
flowchart TB
    A[开始] --> B[结束]
```
"""
        elements = parse_markdown(md)
        mermaid_elems = [e for e in elements if isinstance(e, MermaidBlockElement)]
        assert len(mermaid_elems) == 1
        assert mermaid_elems[0].caption == "图 1-1 系统架构图"

    def test_parse_formula_block(self):
        md = r"""
$$
\frac{\partial u}{\partial t} = \alpha \nabla^2 u
\tag{1-1}
$$
"""
        elements = parse_markdown(md)
        formula_elems = [e for e in elements if isinstance(e, FormulaBlockElement)]
        assert len(formula_elems) == 1
        assert formula_elems[0].tag == "1-1"

    def test_parse_table(self):
        md = """**表 1-1 参数对比**

| 参数 | 值 |
|------|-----|
| A    | 100 |
| B    | 200 |
"""
        elements = parse_markdown(md)
        table_elems = [e for e in elements if isinstance(e, TableElement)]
        assert len(table_elems) == 1
        assert table_elems[0].headers == ["参数", "值"]
        assert len(table_elems[0].rows) == 2
        assert table_elems[0].caption == "表 1-1 参数对比"

    def test_parse_code_block(self):
        md = """```python
def hello():
    print("Hello")
```
"""
        elements = parse_markdown(md)
        code_elems = [e for e in elements if isinstance(e, CodeBlockElement)]
        assert len(code_elems) == 1
        assert code_elems[0].language == "python"

    def test_parse_page_break(self):
        md = "文本\n\n<!-- pagebreak -->\n\n更多文本\n"
        elements = parse_markdown(md)
        break_elems = [e for e in elements if isinstance(e, PageBreakElement)]
        assert len(break_elems) == 1

    def test_parse_complex_document(self):
        md_content = """## 第 1 章 绪论

### 1.1 研究背景

滑坡、泥石流等地质灾害频发。

**图 1-1 系统架构**

```mermaid
flowchart TB
    A[输入] --> B[处理] --> C[输出]
```

公式如下：

$$
E = mc^2
\tag{1-1}
$$

| 参数 | 值 |
|------|-----|
| 质量 | 10 |

<!-- pagebreak -->

## 第 2 章 方法
"""
        elements = parse_markdown(md_content)
        assert len(elements) >= 5

    def test_parser_with_file_path(self, tmp_path):
        md_file = tmp_path / "test.md"
        md_file.write_text("## 标题\n\n段落内容\n", encoding="utf-8")

        parser = MarkdownParser()
        elements = parser.parse_file(str(md_file))
        assert len(elements) >= 1
