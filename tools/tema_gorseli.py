"""README görseli: örnek panoyu yedi temada üretir, ilk sayfaları etiketli tek PNG'de birleştirir.

Geliştirici betiği; skill'lerin parçası değildir. Bir hesaplama motoru (Excel ya da LibreOffice) gerekir.
Kullanım:
    uv run --no-project --with openpyxl --with pymupdf --with pillow python -I tools/tema_gorseli.py docs/temalar.png
"""
from __future__ import annotations

import subprocess
import sys
import tempfile
from pathlib import Path

DEPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(DEPO / "skills" / "excel-design-system" / "scripts"))
import xlkit as xk  # noqa: E402
from PIL import Image, ImageDraw, ImageFont  # noqa: E402

PANO = DEPO / "skills" / "excel-dashboard" / "scripts" / "ornek_pano.py"
SUTUN = 4
GENISLIK = 900
BOSLUK = 24
ETIKET = 56
YAZI_ADAYLARI = ("/System/Library/Fonts/Supplemental/Arial Bold.ttf", r"C:\Windows\Fonts\arialbd.ttf",
                 "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf")


def _yazi_tipi():
    for yol in YAZI_ADAYLARI:
        if Path(yol).is_file():
            return ImageFont.truetype(yol, 30)
    return ImageFont.load_default()


def _onizleme(tema: str, dizin: Path) -> Image.Image:
    xlsx = dizin / f"{tema}.xlsx"
    subprocess.run([sys.executable, "-I", str(PANO), str(xlsx), tema], check=True, capture_output=True)
    sayfalar = xk.onizle(str(xlsx), str(dizin / tema))
    if not sayfalar:
        sys.exit("Hesaplama motoru yok: Excel ya da LibreOffice kurun (python3 .../ortam.py).")
    im = Image.open(sayfalar[0]).convert("RGB")
    return im.resize((GENISLIK, round(im.height * GENISLIK / im.width)), Image.LANCZOS)


def main(cikti: str) -> None:
    temalar = list(xk.TEMALAR)
    with tempfile.TemporaryDirectory(prefix="tema-gorseli-") as gecici:
        resimler = [_onizleme(t, Path(gecici)) for t in temalar]
    yukseklik = max(r.height for r in resimler) + ETIKET
    satir = -(-len(resimler) // SUTUN)
    tuval = Image.new("RGB", (SUTUN * (GENISLIK + BOSLUK) + BOSLUK, satir * (yukseklik + BOSLUK) + BOSLUK), "white")
    kalem, yazi_tipi = ImageDraw.Draw(tuval), _yazi_tipi()
    for i, (ad, resim) in enumerate(zip(temalar, resimler)):
        x = BOSLUK + (i % SUTUN) * (GENISLIK + BOSLUK)
        y = BOSLUK + (i // SUTUN) * (yukseklik + BOSLUK)
        kalem.text((x, y), ad, fill="black", font=yazi_tipi)
        tuval.paste(resim, (x, y + ETIKET))
    tuval.save(cikti, optimize=True)
    print(cikti, tuval.size)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else str(DEPO / "docs" / "temalar.png"))
