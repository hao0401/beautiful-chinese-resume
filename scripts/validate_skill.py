#!/usr/bin/env python
"""Lightweight repository-local validation for this Codex skill."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


def parse_frontmatter(skill_md: Path) -> dict[str, str]:
    text = skill_md.read_text(encoding="utf-8-sig")
    if not text.startswith("---\n"):
        raise ValueError("SKILL.md must start with YAML frontmatter")

    lines = text.splitlines()
    end = None
    for index, line in enumerate(lines[1:], start=1):
        if line == "---":
            end = index
            break
    if end is None:
        raise ValueError("SKILL.md frontmatter is not closed")

    frontmatter: dict[str, str] = {}
    for line in lines[1:end]:
        if not line.strip():
            continue
        if ":" not in line:
            raise ValueError(f"unsupported frontmatter line: {line}")
        key, value = line.split(":", 1)
        frontmatter[key.strip()] = value.strip()
    return frontmatter


def validate(skill_dir: Path) -> list[str]:
    errors: list[str] = []
    skill_md = skill_dir / "SKILL.md"
    openai_yaml = skill_dir / "agents" / "openai.yaml"

    if not skill_md.exists():
        return ["missing SKILL.md"]
    if not openai_yaml.exists():
        errors.append("missing agents/openai.yaml")

    try:
        frontmatter = parse_frontmatter(skill_md)
    except ValueError as exc:
        return [str(exc)]

    expected_name = skill_dir.name
    if frontmatter.get("name") != expected_name:
        errors.append(f"frontmatter name must be {expected_name!r}")
    if not frontmatter.get("description"):
        errors.append("frontmatter description is required")

    if openai_yaml.exists():
        metadata = openai_yaml.read_text(encoding="utf-8-sig")
        for required in ("display_name", "short_description", "default_prompt"):
            if required not in metadata:
                errors.append(f"agents/openai.yaml missing {required}")
        if f"${expected_name}" not in metadata:
            errors.append("default_prompt should mention the skill name")

    return errors


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("skill_dir", nargs="?", default=".")
    args = parser.parse_args()

    skill_dir = Path(args.skill_dir).resolve()
    errors = validate(skill_dir)
    if errors:
        for error in errors:
            print(f"ERROR: {error}", file=sys.stderr)
        return 1

    print(f"skill ok: {skill_dir.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
