"""Ortam raporu: Python sürümü, gerekli paketler ve bulunan hesaplama motorları.

Kullanım:
    python3 scripts/ortam.py          # okunur Türkçe rapor
    python3 scripts/ortam.py --json   # makine için JSON
Zorunlu paket (openpyxl) yoksa çıkış kodu 1.
"""
from __future__ import annotations

import importlib
import json
import os
import platform
import sys

_BURA = os.path.dirname(os.path.abspath(__file__))
if _BURA not in sys.path:
    sys.path.insert(0, _BURA)
import motor  # noqa: E402  (aynı klasördeki motor.py; openpyxl gerektirmez)

PAKETLER = ("openpyxl", "pymupdf", "PIL", "win32com")
ZORUNLU = ("openpyxl",)
PIP = "python3 -m pip install --user openpyxl pymupdf pillow"
UV = "uv run --no-project --with openpyxl --with pymupdf --with pillow python -I betik.py"


def _surum(ad: str) -> str | None:
    try:
        modul = importlib.import_module(ad)
    except ImportError:
        return None
    return str(getattr(modul, "__version__", getattr(modul, "VersionBind", "kurulu")))


def rapor() -> dict:
    paketler = {ad: _surum(ad) for ad in PAKETLER}
    motorlar = motor.bulunan_motorlar()
    eksikler = []
    if not paketler["openpyxl"]:
        eksikler.append(f"openpyxl zorunlu. Kurun: {PIP}  (ya da uv ile: {UV})")
    if not paketler["pymupdf"] or not paketler["PIL"]:
        eksikler.append(f"Önizleme için pymupdf ve pillow gerekir: {PIP}")
    if sys.platform == "win32" and not paketler["win32com"]:
        eksikler.append("Windows Excel otomasyonu için: python -m pip install --user pywin32")
    if not motorlar:
        eksikler.append("Görsel denetim için Microsoft Excel ya da LibreOffice kurun: https://www.libreoffice.org")
    elif motorlar == ["libreoffice"]:
        eksikler.append("Yalnız LibreOffice var: önizleme çalışır, Excel'de grafik görünümü küçük farklar gösterebilir")
    return {"python": platform.python_version(), "paketler": paketler, "motorlar": motorlar, "eksikler": eksikler}


def _yaz(r: dict):
    print(f"Python      {r['python']} ({sys.platform})")
    for ad, surum in r["paketler"].items():
        print(f"{ad:<11} {surum or 'YOK'}")
    print(f"Motorlar    {', '.join(r['motorlar']) or 'YOK'}")
    for satir in r["eksikler"]:
        print(f"- {satir}")


def main(argv: list[str]) -> int:
    r = rapor()
    if "--json" in argv:
        print(json.dumps(r, ensure_ascii=False, indent=2))
    else:
        _yaz(r)
    return 1 if any(not r["paketler"][ad] for ad in ZORUNLU) else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
