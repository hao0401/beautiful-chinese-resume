from __future__ import annotations

from pathlib import Path

from scripts.validate_skill import validate


ROOT = Path(__file__).resolve().parents[1]


def test_skill_metadata_is_valid() -> None:
    assert validate(ROOT) == []


def test_readme_mentions_real_limitations() -> None:
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    assert "Word" in readme
    assert "不编造经历" in readme
