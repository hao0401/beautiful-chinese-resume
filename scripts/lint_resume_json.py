#!/usr/bin/env python
"""Lint structured resume JSON before building a DOCX."""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any

try:
    from build_resume_docx import load_json, normalize_sections, validate_resume_data
except ModuleNotFoundError:  # pragma: no cover
    from scripts.build_resume_docx import load_json, normalize_sections, validate_resume_data


REQUIRED_SECTIONS = ["教育背景", "项目经历", "实习经历", "技能证书"]
ECOMMERCE_HINTS = ["Amazon", "Listing", "ASIN", "BSR", "竞品", "Excel"]
EXAGGERATED_TERMS = ["精通", "显著提升", "大幅增长", "全面负责", "主导", "行业领先"]
MAX_BULLET_CHARS = 86


def flatten_text(value: Any) -> str:
    if isinstance(value, dict):
        return "\n".join(flatten_text(v) for v in value.values())
    if isinstance(value, list):
        return "\n".join(flatten_text(v) for v in value)
    return str(value)


def iter_bullets(data: dict[str, Any]):
    for section in normalize_sections(data):
        for item in section.get("items") or []:
            if isinstance(item, str):
                yield section["title"], item
                continue
            if not isinstance(item, dict):
                continue
            details = item.get("details") or item.get("bullets") or []
            if isinstance(details, str):
                details = [details]
            for detail in details:
                yield section["title"], str(detail)


def lint_resume_json(data: Any) -> dict[str, Any]:
    errors = validate_resume_data(data)
    warnings: list[str] = []

    if not isinstance(data, dict):
        return {"ok": False, "errors": errors, "warnings": warnings}

    if not str(data.get("name") or "").strip():
        warnings.append("name is missing")
    if not str(data.get("target_role") or data.get("role") or "").strip():
        warnings.append("target_role is missing")

    contact = data.get("contact") or []
    if isinstance(contact, str):
        contact_items = [contact]
    elif isinstance(contact, list):
        contact_items = [str(item) for item in contact if str(item).strip()]
    else:
        contact_items = []
    if not contact_items:
        warnings.append("contact is missing")

    section_titles = [section["title"] for section in normalize_sections(data)] if not errors else []
    missing_sections = [title for title in REQUIRED_SECTIONS if title not in section_titles]
    if missing_sections:
        warnings.append("missing recommended sections: " + ", ".join(missing_sections))

    for section_title, bullet in iter_bullets(data):
        clean = re.sub(r"\s+", "", bullet)
        if len(clean) > MAX_BULLET_CHARS:
            warnings.append(f"{section_title} bullet is long ({len(clean)} chars): {bullet[:36]}...")

    all_text = flatten_text(data)
    for term in EXAGGERATED_TERMS:
        if term in all_text:
            warnings.append(f"possibly exaggerated term: {term}")

    target_role = str(data.get("target_role") or data.get("role") or "")
    if any(hint in target_role for hint in ["电商", "运营", "Amazon", "跨境"]):
        matched = [hint for hint in ECOMMERCE_HINTS if hint in all_text]
        if len(matched) < 2:
            warnings.append("ecommerce/operations role has few supported keywords")

    return {"ok": not errors, "errors": errors, "warnings": warnings}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input_json")
    parser.add_argument("--strict", action="store_true", help="Treat warnings as failures.")
    args = parser.parse_args()

    data = load_json(args.input_json)
    result = lint_resume_json(data)
    print(json.dumps(result, ensure_ascii=False, indent=2))

    if result["errors"]:
        return 1
    if args.strict and result["warnings"]:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
