"""ortam.py: Python, paketler ve motorlar için okunur rapor."""
from __future__ import annotations

import json
import subprocess
import sys

from conftest import SCRIPTS

import ortam

BETIK = SCRIPTS / "ortam.py"


def test_rapor_anahtarlari():
    r = ortam.rapor()
    assert set(r) == {"python", "paketler", "motorlar", "eksikler"}
    assert isinstance(r["python"], str)
    assert {"openpyxl", "pymupdf", "PIL", "win32com"} <= set(r["paketler"])
    assert isinstance(r["motorlar"], list) and isinstance(r["eksikler"], list)


def test_openpyxl_yoksa_cikis_1():
    sarmal = (f"import sys, runpy; sys.modules['openpyxl'] = None; sys.argv = ['ortam.py']; "
              f"sys.path.insert(0, {str(SCRIPTS)!r}); runpy.run_path({str(BETIK)!r}, run_name='__main__')")
    sonuc = subprocess.run([sys.executable, "-I", "-c", sarmal], capture_output=True, text=True, timeout=60)
    assert sonuc.returncode == 1
    assert "openpyxl" in sonuc.stdout + sonuc.stderr


def test_json_cikti_gecerli():
    sonuc = subprocess.run([sys.executable, "-I", str(BETIK), "--json"], capture_output=True, text=True, timeout=60)
    veri = json.loads(sonuc.stdout)
    assert "motorlar" in veri
