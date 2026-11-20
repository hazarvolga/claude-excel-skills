"""Ortak test yardımcıları: skill ağacını geçici bir araç klasörüne kurar."""
from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

import pytest

DEPO = Path(__file__).resolve().parent.parent
SKILLS = DEPO / "skills"
SCRIPTS = SKILLS / "excel-design-system" / "scripts"

if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))


def _kur(kok: Path, kip: str) -> Path:
    kok.mkdir(parents=True, exist_ok=True)
    for skill in SKILLS.iterdir():
        if not (skill / "SKILL.md").is_file():
            continue
        hedef = kok / skill.name
        if kip == "bagla":
            os.symlink(skill, hedef)
        else:
            shutil.copytree(skill, hedef, ignore=shutil.ignore_patterns("__pycache__"))
    return kok


@pytest.fixture
def skill_kopyasi(tmp_path):
    """skill_kopyasi("bagla" | "kopya") -> araç klasörü kökü (içinde her skill)."""
    def kur(kip: str) -> Path:
        return _kur(tmp_path / f"arac-{kip}" / "skills", kip)
    return kur
