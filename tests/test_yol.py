"""Skill'ler hiçbir araç klasörüne sabitlenmez; kütüphane kendi konumundan bulunur."""
from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys

from conftest import DEPO, SKILLS

YASAK = ("~/.claude", "~/.Codex", "~/.codex", ".agents/skills")


README_YASAK = ("~/.claude/skills/excel", "~/.Codex", "cp -R")  # README araç klasörü tablosunu gösterebilir


def _taranacaklar():
    yield from (p for p in SKILLS.rglob("*") if p.suffix in (".md", ".py") and "__pycache__" not in p.parts)


def test_sabit_arac_yolu_yok():
    bulunan = [f"{p.relative_to(DEPO)}: {y}" for p in _taranacaklar()
               for y in YASAK if y in p.read_text(encoding="utf-8")]
    readme = (DEPO / "README.md").read_text(encoding="utf-8")
    bulunan += [f"README.md: {y}" for y in README_YASAK if y in readme]
    assert bulunan == []


def _aciklama(skill: str) -> str:
    metin = (SKILLS / skill / "SKILL.md").read_text(encoding="utf-8")
    eslesme = re.search(r'^description:\s*"?(.*?)"?\s*$', metin, re.M)
    assert eslesme, f"{skill}: description yok"
    return eslesme.group(1)


def test_aciklamalar_use_when():
    for skill in ("excel-design-system", "excel-charts", "excel-dashboard"):
        aciklama = _aciklama(skill)
        assert aciklama.startswith("Use when"), skill
        for sozcuk in ("xlkit", "Base layer", "Builds on", "Uses "):
            assert sozcuk not in aciklama, f"{skill}: {sozcuk}"


def _pano(kok, cikti):
    betik = kok / "excel-dashboard" / "scripts" / "ornek_pano.py"
    env = {**os.environ, "HOME": str(cikti.parent / "bos-home"), "USERPROFILE": str(cikti.parent / "bos-home")}
    return subprocess.run([sys.executable, "-I", str(betik), str(cikti), "minimal"],
                          capture_output=True, text=True, timeout=120, env=env)


def test_ornek_pano_kopya_kurulumda_calisir(skill_kopyasi, tmp_path):
    sonuc = _pano(skill_kopyasi("kopya"), tmp_path / "p.xlsx")
    assert sonuc.returncode == 0, sonuc.stderr
    assert (tmp_path / "p.xlsx").is_file()


def test_ornek_pano_bagla_kurulumda_calisir(skill_kopyasi, tmp_path):
    sonuc = _pano(skill_kopyasi("bagla"), tmp_path / "p.xlsx")
    assert sonuc.returncode == 0, sonuc.stderr
    assert (tmp_path / "p.xlsx").is_file()


def test_kardes_yoksa_acik_hata(skill_kopyasi, tmp_path):
    kok = skill_kopyasi("kopya")
    shutil.rmtree(kok / "excel-design-system")
    sonuc = _pano(kok, tmp_path / "p.xlsx")
    assert sonuc.returncode != 0
    assert "excel-design-system bulunamadı" in sonuc.stderr
