#!/usr/bin/env python3
"""Inspect DOCX/DOCM files and emit a JSON summary."""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path
from typing import Any, Dict, List, Optional
from xml.etree import ElementTree as ET

NS = {
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}


def local_name(tag: str) -> str:
    if "}" in tag:
        return tag.rsplit("}", 1)[1]
    return tag


def read_xml(archive: zipfile.ZipFile, name: str) -> Optional[ET.Element]:
    try:
        return ET.fromstring(archive.read(name))
    except KeyError:
        return None


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def preview(text: str, limit: int = 180) -> str:
    compact = clean_text(text)
    if len(compact) <= limit:
        return compact
    return compact[: limit - 3].rstrip() + "..."


def paragraph_text(paragraph: ET.Element) -> str:
    parts: List[str] = []
    for node in paragraph.iter():
        name = local_name(node.tag)
        if name == "t" and node.text:
            parts.append(node.text)
        elif name == "tab":
            parts.append("\t")
        elif name in {"br", "cr"}:
            parts.append("\n")
    return "".join(parts)


def paragraph_style_id(paragraph: ET.Element) -> Optional[str]:
    p_pr = paragraph.find("w:pPr", NS)
    if p_pr is None:
        return None
    style = p_pr.find("w:pStyle", NS)
    if style is None:
        return None
    return style.attrib.get(f"{{{NS['w']}}}val")


def get_style_map(styles_root: Optional[ET.Element]) -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    if styles_root is None:
        return mapping
    for style in styles_root.findall("w:style", NS):
        style_id = style.attrib.get(f"{{{NS['w']}}}styleId")
        name_node = style.find("w:name", NS)
        if style_id and name_node is not None:
            style_name = name_node.attrib.get(f"{{{NS['w']}}}val")
            if style_name:
                mapping[style_id] = style_name
    return mapping


def cell_text(cell: ET.Element) -> str:
    texts = [paragraph_text(p) for p in cell.findall(".//w:p", NS)]
    return "\n".join(filter(None, texts)).strip()


def extract_table(table: ET.Element, max_rows: int = 4, max_cols: int = 4) -> Dict[str, Any]:
    rows = table.findall("w:tr", NS)
    preview_rows: List[List[str]] = []
    max_column_count = 0
    for row in rows[:max_rows]:
        cells = row.findall("w:tc", NS)
        max_column_count = max(max_column_count, len(cells))
        preview_rows.append([preview(cell_text(cell), 80) for cell in cells[:max_cols]])
    if rows:
        max_column_count = max(max_column_count, max(len(row.findall("w:tc", NS)) for row in rows))
    return {
        "row_count": len(rows),
        "column_count": max_column_count,
        "preview_rows": preview_rows,
    }


def extract_text_from_part(root: Optional[ET.Element], limit: int) -> List[str]:
    if root is None:
        return []
    lines: List[str] = []
    for paragraph in root.findall(".//w:p", NS):
        text = clean_text(paragraph_text(paragraph))
        if text:
            lines.append(text)
        if len(lines) >= limit:
            break
    return lines


def extract_comments(root: Optional[ET.Element], limit: int) -> List[Dict[str, Any]]:
    if root is None:
        return []
    comments: List[Dict[str, Any]] = []
    for comment in root.findall("w:comment", NS)[:limit]:
        comment_id = comment.attrib.get(f"{{{NS['w']}}}id")
        author = comment.attrib.get(f"{{{NS['w']}}}author")
        text = clean_text(" ".join(paragraph_text(p) for p in comment.findall(".//w:p", NS)))
        comments.append(
            {
                "id": comment_id,
                "author": author,
                "text": preview(text, 220),
            }
        )
    return comments


def extract_properties(core_root: Optional[ET.Element], app_root: Optional[ET.Element]) -> Dict[str, Any]:
    props: Dict[str, Any] = {}
    if core_root is not None:
        for key, xpath in {
            "title": "dc:title",
            "subject": "dc:subject",
            "creator": "dc:creator",
            "description": "dc:description",
            "last_modified_by": "cp:lastModifiedBy",
            "created": "dcterms:created",
            "modified": "dcterms:modified",
        }.items():
            node = core_root.find(xpath, NS)
            if node is not None and node.text:
                props[key] = node.text
    if app_root is not None:
        for key, xpath in {
            "application": "ep:Application",
            "pages": "ep:Pages",
            "words": "ep:Words",
            "characters": "ep:Characters",
            "lines": "ep:Lines",
            "paragraphs": "ep:Paragraphs",
            "company": "ep:Company",
        }.items():
            node = app_root.find(xpath, NS)
            if node is not None and node.text:
                props[key] = node.text
    return props


def build_report(path: Path, max_paragraphs: int, max_tables: int) -> Dict[str, Any]:
    with zipfile.ZipFile(path) as archive:
        document_root = read_xml(archive, "word/document.xml")
        if document_root is None:
            raise ValueError("word/document.xml was not found in the archive.")

        styles_root = read_xml(archive, "word/styles.xml")
        core_root = read_xml(archive, "docProps/core.xml")
        app_root = read_xml(archive, "docProps/app.xml")
        comments_root = read_xml(archive, "word/comments.xml")
        style_map = get_style_map(styles_root)

        all_paragraphs = document_root.findall(".//w:p", NS)
        body = document_root.find("w:body", NS)
        if body is None:
            raise ValueError("Document body was not found.")

        style_counts: Counter[str] = Counter()
        for paragraph in all_paragraphs:
            style_id = paragraph_style_id(paragraph) or "Normal"
            style_counts[style_id] += 1

        block_samples: List[Dict[str, Any]] = []
        table_samples: List[Dict[str, Any]] = []
        paragraph_index = 0
        table_index = 0
        for child in list(body):
            name = local_name(child.tag)
            if name == "p" and len(block_samples) < max_paragraphs:
                text = paragraph_text(child)
                style_id = paragraph_style_id(child)
                block_samples.append(
                    {
                        "index": paragraph_index,
                        "type": "paragraph",
                        "style_id": style_id,
                        "style_name": style_map.get(style_id or "", style_id or "Normal"),
                        "text": preview(text),
                    }
                )
                paragraph_index += 1
            elif name == "tbl" and len(table_samples) < max_tables:
                table_samples.append({"index": table_index, **extract_table(child)})
                table_index += 1

        headers = []
        footers = []
        media_files = []
        for member in archive.namelist():
            if member.startswith("word/header") and member.endswith(".xml"):
                headers.append({"part": member, "lines": extract_text_from_part(read_xml(archive, member), 8)})
            elif member.startswith("word/footer") and member.endswith(".xml"):
                footers.append({"part": member, "lines": extract_text_from_part(read_xml(archive, member), 8)})
            elif member.startswith("word/media/"):
                media_files.append(member)

        section_count = len(document_root.findall(".//w:sectPr", NS)) or 1
        non_empty_paragraphs = sum(1 for p in all_paragraphs if clean_text(paragraph_text(p)))

        return {
            "path": str(path.resolve()),
            "properties": extract_properties(core_root, app_root),
            "summary": {
                "paragraph_count": len(all_paragraphs),
                "non_empty_paragraph_count": non_empty_paragraphs,
                "table_count": len(document_root.findall(".//w:tbl", NS)),
                "section_count": section_count,
                "comment_count": len(comments_root.findall(".//w:comment", NS)) if comments_root is not None else 0,
                "media_file_count": len(media_files),
            },
            "styles_used": [
                {
                    "style_id": style_id,
                    "style_name": style_map.get(style_id, style_id),
                    "count": count,
                }
                for style_id, count in style_counts.most_common()
            ],
            "block_samples": block_samples,
            "table_samples": table_samples,
            "headers": headers,
            "footers": footers,
            "comments": extract_comments(comments_root, 12),
            "media_files": media_files,
        }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Inspect a DOCX or DOCM file and emit a JSON summary.")
    parser.add_argument("path", help="Path to the DOCX or DOCM file.")
    parser.add_argument("--max-paragraphs", type=int, default=40, help="Maximum number of top-level paragraphs to sample.")
    parser.add_argument("--max-tables", type=int, default=10, help="Maximum number of top-level tables to sample.")
    parser.add_argument("--output", help="Optional output JSON file path.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    path = Path(args.path)
    if not path.exists():
        print(f"File not found: {path}", file=sys.stderr)
        return 1
    if not zipfile.is_zipfile(path):
        print(f"Not a ZIP-based Office file: {path}", file=sys.stderr)
        return 1

    try:
        report = build_report(path, args.max_paragraphs, args.max_tables)
    except Exception as exc:
        print(f"Inspection failed: {exc}", file=sys.stderr)
        return 1

    output = json.dumps(report, ensure_ascii=False, indent=2)
    if args.output:
        Path(args.output).write_text(output, encoding="utf-8")
    else:
        print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
