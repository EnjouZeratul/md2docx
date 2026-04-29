"""Markdown parser for md2docx.

Parses markdown text into structured element list.
"""

import re
from pathlib import Path
from dataclasses import dataclass, field
from typing import List, Optional, Tuple


@dataclass
class HeadingElement:
    """Heading element.

    level: 2=章, 3=节, 4=小节
    """
    level: int
    text: str


@dataclass
class ParagraphElement:
    """Paragraph element."""
    text: str
    has_inline_math: bool = False


@dataclass
class MermaidBlockElement:
    """Mermaid diagram block."""
    code: str
    caption: str = ""


@dataclass
class FormulaBlockElement:
    """LaTeX formula block."""
    latex: str
    tag: str = ""  # Formula number from \tag{}


@dataclass
class TableElement:
    """Table element."""
    headers: List[str] = field(default_factory=list)
    rows: List[List[str]] = field(default_factory=list)
    caption: str = ""


@dataclass
class CodeBlockElement:
    """Code block element."""
    code: str
    language: str = ""


@dataclass
class ImageElement:
    """Image element."""
    path: str
    caption: str = ""


@dataclass
class PageBreakElement:
    """Page break element."""
    pass


@dataclass
class AbstractElement:
    """Abstract/摘要 element."""
    lang: str  # "zh" or "en"
    body: str
    keywords: str = ""


# Element type alias for type hints
Element = (
    HeadingElement | ParagraphElement | MermaidBlockElement |
    FormulaBlockElement | TableElement | CodeBlockElement |
    ImageElement | PageBreakElement | AbstractElement
)


class MarkdownParser:
    """Markdown parser that converts MD text to element list."""

    def __init__(self, base_path: str = "."):
        """Initialize parser.

        Args:
            base_path: Base path for resolving relative image paths
        """
        self.base_path = Path(base_path).resolve()

    def parse(self, text: str) -> List[Element]:
        """Parse markdown text to element list.

        Args:
            text: Markdown text content

        Returns:
            List of Element objects
        """
        elements = []
        lines = text.split('\n')
        i = 0

        # State variables
        in_code_block = False
        code_block_lang = ""
        code_block_content = []
        in_formula_block = False
        formula_content = []

        # Caption tracking (for figures/tables)
        pending_caption = None

        while i < len(lines):
            line = lines[i]

            # Handle code blocks
            if line.strip().startswith('```'):
                if not in_code_block:
                    # Start of code block
                    in_code_block = True
                    code_block_lang = line.strip()[3:].strip()
                    code_block_content = []
                    i += 1
                    continue
                else:
                    # End of code block
                    in_code_block = False
                    code = '\n'.join(code_block_content)

                    if code_block_lang == 'mermaid':
                        elements.append(MermaidBlockElement(
                            code=code,
                            caption=pending_caption or ""
                        ))
                        pending_caption = None
                    else:
                        elements.append(CodeBlockElement(
                            code=code,
                            language=code_block_lang
                        ))
                    code_block_lang = ""
                    code_block_content = []
                    i += 1
                    continue

            if in_code_block:
                code_block_content.append(line)
                i += 1
                continue

            # Handle formula blocks ($$ on its own line)
            if line.strip() == '$$':
                if not in_formula_block:
                    in_formula_block = True
                    formula_content = []
                else:
                    in_formula_block = False
                    formula_text = '\n'.join(formula_content)

                    # Extract tag if present
                    tag_match = re.search(r'\\tag\{([^}]+)\}', formula_text)
                    tag = tag_match.group(1) if tag_match else ""
                    # Remove tag from latex
                    latex = re.sub(r'\\tag\{[^}]+\}', '', formula_text).strip()

                    elements.append(FormulaBlockElement(
                        latex=latex,
                        tag=tag
                    ))
                    formula_content = []
                i += 1
                continue

            if in_formula_block:
                formula_content.append(line)
                i += 1
                continue

            # Skip empty lines
            if not line.strip():
                i += 1
                continue

            # Check for page break
            if '<!-- pagebreak -->' in line:
                elements.append(PageBreakElement())
                i += 1
                continue

            # Check for heading
            heading_match = re.match(r'^(#{1,6})\s+(.+)$', line)
            if heading_match:
                level = len(heading_match.group(1))
                text = heading_match.group(2).strip()
                elements.append(HeadingElement(level=level, text=text))
                i += 1
                continue

            # Check for figure/table caption (Bold text starting with 图/表)
            caption_match = re.match(r'^\*\*(图\s+\d+-\d+[^*]*|表\s+\d+-\d+[^*]*)\*\*', line)
            if caption_match:
                pending_caption = caption_match.group(1)
                i += 1
                continue

            # Check for table
            if '|' in line and not line.strip().startswith('|---'):
                # Look ahead for separator line
                if i + 1 < len(lines) and re.match(r'^\|[-:|]+\|', lines[i + 1]):
                    # Parse table
                    table, consumed = self._parse_table(lines, i)
                    if table:
                        table.caption = pending_caption or ""
                        pending_caption = None
                        elements.append(table)
                        i += consumed
                        continue

            # Check for image
            img_match = re.match(r'!\[([^\]]*)\]\(([^)]+)\)', line)
            if img_match:
                img_path = img_match.group(2)
                # Convert to absolute path if relative
                if not Path(img_path).is_absolute():
                    img_path = str(self.base_path / img_path)
                elements.append(ImageElement(path=img_path, caption=pending_caption or ""))
                pending_caption = None
                i += 1
                continue

            # Check for horizontal rule (abstract boundary)
            if line.strip() == '---':
                i += 1
                continue

            # Otherwise, treat as paragraph
            # Collect paragraph lines
            para_lines = [line]
            i += 1
            while i < len(lines) and lines[i].strip() and not self._is_special_line(lines[i]):
                para_lines.append(lines[i])
                i += 1

            para_text = ' '.join(para_lines)
            has_inline_math = '$' in para_text

            elements.append(ParagraphElement(text=para_text, has_inline_math=has_inline_math))

        return elements

    def _parse_table(self, lines: List[str], start: int) -> Tuple[Optional[TableElement], int]:
        """Parse a table starting at given line index.

        Returns:
            Tuple of (TableElement or None, number of lines consumed)
        """
        i = start

        # Get header row
        header_line = lines[i].strip()
        if not header_line.startswith('|'):
            return None, 0

        headers = [cell.strip() for cell in header_line.split('|') if cell.strip()]
        i += 1

        # Skip separator
        if i >= len(lines) or not re.match(r'^\|[-:|]+\|', lines[i]):
            return None, 0
        i += 1

        # Parse data rows
        rows = []
        while i < len(lines):
            line = lines[i].strip()
            if not line.startswith('|'):
                break
            row = [cell.strip() for cell in line.split('|') if cell.strip()]
            if row:
                rows.append(row)
            i += 1

        if not headers:
            return None, 0

        return TableElement(headers=headers, rows=rows), i - start

    def _is_special_line(self, line: str) -> bool:
        """Check if line is a special markdown element."""
        line = line.strip()
        if not line:
            return True
        if line.startswith('#'):
            return True
        if line.startswith('```'):
            return True
        if line.startswith('$$'):
            return True
        if line.startswith('|'):
            return True
        if line.startswith('!['):
            return True
        if line == '---':
            return True
        if '<!-- pagebreak -->' in line:
            return True
        return False

    def parse_file(self, file_path: str) -> List[Element]:
        """Parse markdown file.

        Args:
            file_path: Path to markdown file

        Returns:
            List of Element objects
        """
        file_path = Path(file_path)
        self.base_path = file_path.parent

        with open(file_path, 'r', encoding='utf-8') as f:
            text = f.read()

        return self.parse(text)


def parse_markdown(text: str, base_path: str = ".") -> List[Element]:
    """Convenience function to parse markdown text.

    Args:
        text: Markdown text content
        base_path: Base path for resolving relative paths

    Returns:
        List of Element objects
    """
    parser = MarkdownParser(base_path=base_path)
    return parser.parse(text)
