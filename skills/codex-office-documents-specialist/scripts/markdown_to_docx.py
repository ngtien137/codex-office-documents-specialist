#!/usr/bin/env python3
"""Create a DOCX document from Markdown content."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from docx_python_tools import create_document, add_toc, render_markdown_into_document, set_document_metadata, set_header_footer


def read_markdown(path: str | None, inline_markdown: str | None) -> str:
    if inline_markdown is not None:
        return inline_markdown
    if path is None:
        return sys.stdin.read()
    return Path(path).read_text(encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Convert Markdown content into a DOCX document.")
    parser.add_argument("input", nargs="?", help="Optional Markdown file path. If omitted, reads stdin unless --markdown is used.")
    parser.add_argument("--markdown", help="Inline Markdown content.")
    parser.add_argument("--output", required=True, help="Output DOCX path.")
    parser.add_argument("--template", help="Optional DOCX template path.")
    parser.add_argument("--title", help="Document title metadata.")
    parser.add_argument("--author", help="Document author metadata.")
    parser.add_argument("--subject", help="Document subject metadata.")
    parser.add_argument("--header", help="Header text. Supports {page} and {pages}.")
    parser.add_argument("--footer", help="Footer text. Supports {page} and {pages}.")
    parser.add_argument("--toc", action="store_true", help="Insert a table of contents at the top of the document.")
    parser.add_argument("--summary", action="store_true", help="Print a JSON summary after writing the document.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    markdown = read_markdown(args.input, args.markdown)
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    document = create_document(args.template)
    set_document_metadata(document, title=args.title, author=args.author, subject=args.subject)

    if args.toc:
        add_toc(document)
    if args.header:
        set_header_footer(document, args.header, kind="header")
    if args.footer:
        set_header_footer(document, args.footer, kind="footer")

    stats = render_markdown_into_document(document, markdown)
    document.save(str(output_path))

    if args.summary:
        print(
            json.dumps(
                {
                    "path": str(output_path.resolve()),
                    "template": str(Path(args.template).resolve()) if args.template else None,
                    "stats": stats,
                },
                ensure_ascii=False,
                indent=2,
            )
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
