"""Command-line interface for md2docx."""

import argparse
import sys
from pathlib import Path

from md2docx import __version__
from md2docx.config import load_config, ConfigError
from md2docx.parser import MarkdownParser
from md2docx.renderer.docx_builder import DocxBuilder


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        prog='md2docx',
        description='Convert Markdown thesis to DOCX with Mermaid and LaTeX support'
    )

    parser.add_argument('--version', action='version', version=f'%(prog)s {__version__}')
    parser.add_argument('-v', '--verbose', action='count', default=0,
                       help='Increase verbosity')

    subparsers = parser.add_subparsers(dest='command', help='Commands')

    # Convert command
    convert_parser = subparsers.add_parser('convert', help='Convert MD to DOCX')
    convert_parser.add_argument('input', help='Input Markdown file')
    convert_parser.add_argument('-o', '--output', help='Output DOCX file')
    convert_parser.add_argument('-t', '--template', help='Template config file')

    # Validate command
    validate_parser = subparsers.add_parser('validate', help='Validate MD syntax')
    validate_parser.add_argument('input', help='Input Markdown file')

    # Preview command
    preview_parser = subparsers.add_parser('preview', help='Preview parsed elements')
    preview_parser.add_argument('input', help='Input Markdown file')

    args = parser.parse_args()

    if args.command is None:
        parser.print_help()
        return 0

    try:
        if args.command == 'convert':
            return cmd_convert(args)
        elif args.command == 'validate':
            return cmd_validate(args)
        elif args.command == 'preview':
            return cmd_preview(args)
    except ConfigError as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        if args.verbose > 1:
            import traceback
            traceback.print_exc()
        return 1

    return 0


def cmd_convert(args) -> int:
    """Execute convert command."""
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"[ERROR] File not found: {input_path}", file=sys.stderr)
        return 1

    if args.output:
        output_path = Path(args.output)
    else:
        output_path = input_path.with_suffix('.docx')

    config = load_config(args.template)

    if args.verbose:
        print(f"[INFO] Loading config: {args.template or 'default'}")

    if args.verbose:
        print(f"[INFO] Parsing {input_path}...")

    parser = MarkdownParser(base_path=str(input_path.parent))
    elements = parser.parse_file(str(input_path))

    if args.verbose:
        print(f"[INFO] Found {len(elements)} elements")

    if args.verbose:
        print(f"[INFO] Building DOCX...")

    builder = DocxBuilder(config, verbose=args.verbose)
    builder.build(elements, str(output_path))

    print(f"[SUCCESS] Output: {output_path}")
    return 0


def cmd_validate(args) -> int:
    """Execute validate command."""
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"[ERROR] File not found: {input_path}", file=sys.stderr)
        return 1

    print(f"[INFO] Checking {input_path}")

    with open(input_path, 'r', encoding='utf-8') as f:
        content = f.read()

    errors = []
    warnings = []

    code_block_count = content.count('```')
    if code_block_count % 2 != 0:
        errors.append("Line X: Unclosed code block")

    import re
    headings = re.findall(r'^(#{2,4})\s+(.+)$', content, re.MULTILINE)
    print(f"[OK] 标题层级: 发现 {len(headings)} 个标题")

    mermaid_count = content.count('```mermaid')
    print(f"[OK] Mermaid 图: 发现 {mermaid_count} 个图表")

    formula_count = len(re.findall(r'^\$\$$', content, re.MULTILINE))
    print(f"[OK] 公式块: 发现 {formula_count // 2} 个块级公式")

    if errors:
        for e in errors:
            print(f"[ERROR] {e}")
        return 1

    if warnings:
        for w in warnings:
            print(f"[WARN] {w}")

    print("[OK] Validation complete")
    return 0


def cmd_preview(args) -> int:
    """Execute preview command."""
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"[ERROR] File not found: {input_path}", file=sys.stderr)
        return 1

    print(f"[INFO] Parsing {input_path}")

    parser = MarkdownParser(base_path=str(input_path.parent))
    elements = parser.parse_file(str(input_path))

    print("\n## 元素列表\n")

    for i, elem in enumerate(elements, 1):
        elem_type = type(elem).__name__
        if hasattr(elem, 'text'):
            preview = elem.text[:50] + '...' if len(elem.text) > 50 else elem.text
            print(f"[{i}] {elem_type}: \"{preview}\"")
        elif hasattr(elem, 'code'):
            preview = elem.code[:30] + '...' if len(elem.code) > 30 else elem.code
            caption = getattr(elem, 'caption', '')
            print(f"[{i}] {elem_type}: \"{preview}\" (caption: \"{caption}\")")
        elif hasattr(elem, 'latex'):
            preview = elem.latex[:30] + '...' if len(elem.latex) > 30 else elem.latex
            tag = getattr(elem, 'tag', '')
            print(f"[{i}] {elem_type}: \"{preview}\" (tag: \"{tag}\")")
        else:
            print(f"[{i}] {elem_type}")

    print("\n## 统计\n")
    from collections import Counter
    type_counts = Counter(type(e).__name__ for e in elements)
    for t, c in sorted(type_counts.items()):
        print(f"- {t}: {c} 个")

    return 0


if __name__ == '__main__':
    sys.exit(main())
