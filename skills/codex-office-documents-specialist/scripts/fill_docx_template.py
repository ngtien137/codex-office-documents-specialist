#!/usr/bin/env python3
"""Fill DOCX templates containing {{placeholders}}."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from docx_python_tools import create_document, replace_placeholders_in_document, set_document_metadata


def load_context(data_path: str | None, inline_json: str | None) -> dict:
    if inline_json is not None:
        return json.loads(inline_json)
    if data_path is None:
        raise ValueError("Either --data or --json is required.")
    return json.loads(Path(data_path).read_text(encoding="utf-8"))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Fill a DOCX template using placeholder values from JSON.")
    parser.add_argument("template", help="Path to the DOCX template.")
    parser.add_argument("--output", required=True, help="Output DOCX path.")
    parser.add_argument("--data", help="Path to a JSON file with placeholder values.")
    parser.add_argument("--json", dest="inline_json", help="Inline JSON object with placeholder values.")
    parser.add_argument("--title", help="Document title metadata.")
    parser.add_argument("--author", help="Document author metadata.")
    parser.add_argument("--subject", help="Document subject metadata.")
    parser.add_argument("--summary", action="store_true", help="Print a JSON summary after writing the document.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    template_path = Path(args.template)
    if not template_path.exists():
        raise FileNotFoundError(f"Template not found: {template_path}")

    context = load_context(args.data, args.inline_json)
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    document = create_document(str(template_path))
    replace_placeholders_in_document(document, {key: "" if value is None else str(value) for key, value in context.items()})
    set_document_metadata(document, title=args.title, author=args.author, subject=args.subject)
    document.save(str(output_path))

    if args.summary:
        print(
            json.dumps(
                {
                    "path": str(output_path.resolve()),
                    "template": str(template_path.resolve()),
                    "keys": sorted(context.keys()),
                },
                ensure_ascii=False,
                indent=2,
            )
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
