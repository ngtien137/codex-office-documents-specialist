#!/usr/bin/env python3
"""Profile a DOCX/DOCM file by structure, likely document type, and topic."""

from __future__ import annotations

import argparse
import json
import unicodedata
from pathlib import Path
from typing import Any, Dict, Iterable, List, Tuple

from inspect_docx import build_report


def normalize_text(text: str) -> str:
    lowered = text.lower()
    decomposed = unicodedata.normalize("NFD", lowered)
    stripped = "".join(ch for ch in decomposed if unicodedata.category(ch) != "Mn")
    return " ".join(stripped.split())


def joined_text(chunks: Iterable[str]) -> str:
    return " ".join(chunk for chunk in chunks if chunk).strip()


def count_matches(text: str, keywords: Iterable[str]) -> int:
    return sum(1 for keyword in keywords if keyword in text)


def document_kind_scores(text: str) -> Dict[str, int]:
    return {
        "academic_thesis": count_matches(
            text,
            [
                "luan van",
                "luan an",
                "tieu luan",
                "do an",
                "bao cao tot nghiep",
                "chuong 1",
                "tai lieu tham khao",
                "phu luc",
                "loi cam on",
                "tom tat",
                "ket luan va kien nghi",
            ],
        ),
        "academic_report": count_matches(
            text,
            [
                "bao cao",
                "nghien cuu",
                "phuong phap nghien cuu",
                "ket qua va thao luan",
                "muc tieu",
                "doi tuong va pham vi",
            ],
        ),
        "business_report": count_matches(
            text,
            [
                "executive summary",
                "market share",
                "kpi",
                "quarterly",
                "financial",
                "revenue",
                "cost",
                "recommendation",
                "bao cao kinh doanh",
            ],
        ),
        "technical_report": count_matches(
            text,
            [
                "system architecture",
                "api",
                "implementation",
                "testing",
                "deployment",
                "giai phap",
                "kien truc",
                "he thong",
                "thuat toan",
                "ky thuat",
            ],
        ),
        "proposal_plan": count_matches(
            text,
            [
                "proposal",
                "de xuat",
                "ke hoach",
                "roadmap",
                "scope",
                "timeline",
                "muc tieu du an",
            ],
        ),
    }


def guess_document_kind(text: str) -> Tuple[str, int]:
    scores = document_kind_scores(text)
    best_kind = max(scores, key=scores.get)
    return best_kind, scores[best_kind]


def topic_scores(text: str) -> Dict[str, int]:
    return {
        "environment": count_matches(
            text,
            [
                "moi truong",
                "tai nguyen",
                "waste",
                "solid waste",
                "gis",
                "chat thai",
                "nuoc thai",
                "sustainability",
                "climate",
            ],
        ),
        "education": count_matches(
            text,
            [
                "giao duc",
                "dai hoc",
                "hoc sinh",
                "sinh vien",
                "chuong trinh dao tao",
                "learning",
                "curriculum",
                "pedagogy",
            ],
        ),
        "business_finance": count_matches(
            text,
            [
                "business",
                "market",
                "finance",
                "revenue",
                "cost",
                "doanh thu",
                "loi nhuan",
                "tai chinh",
                "khach hang",
            ],
        ),
        "technology_engineering": count_matches(
            text,
            [
                "software",
                "system",
                "api",
                "database",
                "engineering",
                "kien truc",
                "he thong",
                "thiet ke",
                "implementation",
            ],
        ),
        "health_biology": count_matches(
            text,
            [
                "health",
                "medical",
                "benh",
                "suc khoe",
                "biology",
                "cell",
                "clinical",
                "y hoc",
            ],
        ),
        "law_policy": count_matches(
            text,
            [
                "law",
                "policy",
                "phap luat",
                "quy dinh",
                "nghi dinh",
                "compliance",
                "governance",
            ],
        ),
    }


def guess_topic(text: str) -> Tuple[str, int]:
    scores = topic_scores(text)
    best_topic = max(scores, key=scores.get)
    return best_topic, scores[best_topic]


def heading_counts(styles_used: List[Dict[str, Any]]) -> Dict[str, int]:
    counts: Dict[str, int] = {}
    for item in styles_used:
        name = str(item.get("style_name") or item.get("style_id") or "")
        if name.startswith("Heading"):
            counts[name] = int(item.get("count", 0))
    return counts


def advice_for_kind(kind: str, report: Dict[str, Any]) -> Dict[str, Any]:
    styles = report.get("styles_used", [])
    headings = heading_counts(styles)
    has_toc = any(str(item.get("style_name", "")).startswith("TOC") for item in styles)
    summary = report.get("summary", {})

    if kind in {"academic_thesis", "academic_report"}:
        return {
            "should_offer_academic_update": True,
            "format_advice": [
                "Offer academic-formatting defaults before editing.",
                "Check chapter headings, subsection hierarchy, and whether numbering stays within three levels.",
                "Verify front matter vs main matter page numbering using sections.",
                "Check tables, figures, references, and appendix structure against academic-formatting.md.",
            ],
            "suggested_change_items": [
                "Normalize page setup, margins, font, paragraph spacing, and first-line indent to the academic guide.",
                "Normalize chapter and subsection headings into real heading styles with consistent numbering.",
                "Split front matter and main matter page numbering if sections are not configured correctly.",
                "Standardize table and figure captions, chapter-based numbering, and update the table of contents.",
            ],
            "notes": {
                "has_toc_styles": has_toc,
                "heading_counts": headings,
                "section_count": summary.get("section_count"),
            },
        }

    if kind == "business_report":
        return {
            "should_offer_academic_update": False,
            "format_advice": [
                "Check whether the document needs an executive summary near the front.",
                "Keep headings concise, tables readable, and charts labeled for decision-makers.",
                "Prefer short conclusion and action-oriented recommendations.",
            ],
            "suggested_change_items": [
                "Normalize heading hierarchy and tighten spacing for a cleaner business-report layout.",
                "Add or polish an executive summary if the user wants decision-oriented framing.",
                "Standardize table, chart, and appendix labeling for easier scanning.",
            ],
            "notes": {
                "has_toc_styles": has_toc,
                "heading_counts": headings,
            },
        }

    if kind == "technical_report":
        return {
            "should_offer_academic_update": False,
            "format_advice": [
                "Check whether the document clearly separates problem, method, implementation, results, and references.",
                "Use code, tables, and diagrams consistently and keep captions explicit.",
                "Prefer heading-driven navigation and a clean appendix for technical details.",
            ],
            "suggested_change_items": [
                "Normalize heading structure and technical section order.",
                "Standardize tables, diagrams, and appendix references.",
                "Improve readability of implementation and results sections with spacing and captions.",
            ],
            "notes": {
                "has_toc_styles": has_toc,
                "heading_counts": headings,
            },
        }

    return {
        "should_offer_academic_update": False,
        "format_advice": [
            "Inspect heading structure, spacing, captions, and page numbering before editing.",
            "Keep the format aligned with the apparent document purpose rather than applying academic rules by default.",
        ],
        "suggested_change_items": [
            "Normalize heading hierarchy and spacing.",
            "Update captions, page numbering, and navigation aids only where needed.",
        ],
        "notes": {
            "has_toc_styles": has_toc,
            "heading_counts": headings,
        },
    }


def advice_for_topic(topic: str) -> List[str]:
    mapping = {
        "environment": [
            "Check whether terms, figures, and maps use consistent environmental naming.",
            "If the report uses GIS or waste-management data, verify map, table, and unit captions carefully.",
        ],
        "education": [
            "Check whether literature review, educational context, and comparative discussion are clearly separated.",
            "If the document compares systems or institutions, make tables and references easy to scan.",
        ],
        "business_finance": [
            "Make findings and recommendations easier to scan with concise headings and summary tables.",
            "If financial metrics appear, check number formatting, units, and time periods for consistency.",
        ],
        "technology_engineering": [
            "Check whether architecture, method, implementation, and validation sections are clearly distinct.",
            "If diagrams or code-adjacent tables appear, ensure captions and references are explicit.",
        ],
        "health_biology": [
            "Check scientific naming, units, and method/result separation carefully.",
            "Ensure tables and figure captions carry enough context to stand alone.",
        ],
        "law_policy": [
            "Check citation style, reference ordering, and whether normative terms stay consistent throughout.",
            "Keep section numbering stable because legal and policy readers depend on navigation precision.",
        ],
    }
    return mapping.get(
        topic,
        [
            "Check whether the section order and caption style support the apparent subject matter.",
            "Keep formatting changes aligned with document purpose, not just visual cleanup.",
        ],
    )


def build_profile(path: Path) -> Dict[str, Any]:
    report = build_report(path, max_paragraphs=80, max_tables=12)
    properties = report.get("properties", {})
    block_samples = report.get("block_samples", [])
    headers = report.get("headers", [])
    footers = report.get("footers", [])

    text_sources = [
        str(properties.get("title") or ""),
        *(str(item.get("text") or "") for item in block_samples),
        *(joined_text(item.get("lines", [])) for item in headers),
        *(joined_text(item.get("lines", [])) for item in footers),
    ]
    normalized = normalize_text(joined_text(text_sources))

    kind, kind_score = guess_document_kind(normalized)
    topic, topic_score = guess_topic(normalized)
    kind_advice = advice_for_kind(kind, report)
    topic_advice = advice_for_topic(topic)

    confidence = "low"
    if kind_score >= 6 or topic_score >= 5:
        confidence = "high"
    elif kind_score >= 3 or topic_score >= 2:
        confidence = "medium"

    return {
        "path": str(path.resolve()),
        "document_kind_guess": kind,
        "topic_guess": topic,
        "confidence": confidence,
        "scores": {
            "document_kind": document_kind_scores(normalized),
            "topic": topic_scores(normalized),
        },
        "format_advice": kind_advice["format_advice"],
        "topic_advice": topic_advice,
        "should_offer_academic_update": kind_advice["should_offer_academic_update"],
        "suggested_change_items": kind_advice["suggested_change_items"],
        "notes": kind_advice["notes"],
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Profile a DOCX or DOCM file by type, topic, and formatting advice.")
    parser.add_argument("path", help="Path to the DOCX or DOCM file.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    profile = build_profile(Path(args.path))
    output = json.dumps(profile, ensure_ascii=False, indent=2)
    try:
        print(output)
    except UnicodeEncodeError:
        import sys

        sys.stdout.buffer.write(output.encode("utf-8"))
        sys.stdout.buffer.write(b"\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
