"""Hesaplatma ve PDF motorları: macOS Excel, Windows Excel, LibreOffice.

Sıra: excel-mac → excel-win → libreoffice. Otomatik modda bir motor hata verirse nedeni uyarı olarak yazılır
ve sıradakine geçilir; hiçbiri yoksa uyarı verilip None döner. Motor açıkça seçilirse yalnız o denenir ve
hatası yükseltilir. LibreOffice otomatik modda özgün dosyaya yazmaz (biçimi bozabilir); yalnız
motor="libreoffice" açıkça verilince kaydeder.
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
import warnings
from pathlib import Path

MOTORLAR = ("excel-mac", "excel-win", "libreoffice")
LO_ZAMAN_ASIMI = 180
EXCEL_ZAMAN_ASIMI = 300
MAC_EXCEL = Path("/Applications/Microsoft Excel.app")
LO_ADAYLARI = (
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
)
YOK_UYARISI = "Hiçbir hesaplama motoru bulunamadı: görsel denetim yapılamadı"


class MotorHatasi(RuntimeError):
    """Motor çalıştı ama beklenen çıktıyı üretmedi."""


# ---------------------------------------------------------------- bulma


def _soffice() -> str | None:
    for ad in ("soffice", "libreoffice"):
        yol = shutil.which(ad)
        if yol:
            return yol
    return next((a for a in LO_ADAYLARI if os.path.isfile(a)), None)


def _win_excel_var() -> bool:
    try:
        import winreg  # noqa: PLC0415
        import win32com.client  # noqa: F401, PLC0415
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Excel.Application"))
        return True
    except (ImportError, OSError):
        return False


def bulunan_motorlar() -> list[str]:
    bulunan = []
    if sys.platform == "darwin" and MAC_EXCEL.exists():
        bulunan.append("excel-mac")
    if sys.platform == "win32" and _win_excel_var():
        bulunan.append("excel-win")
    if _soffice():
        bulunan.append("libreoffice")
    return bulunan


# ---------------------------------------------------------------- macOS Excel (AppleScript)


def _applescript(betik: str, zaman: int = EXCEL_ZAMAN_ASIMI):
    """İzin penceresi ya da kurtarma penceresi Excel'i kilitlerse süreç zaman aşımıyla düşer, sıradaki motora geçilir."""
    subprocess.run(["osascript", "-e", f"with timeout of {zaman} seconds\n{betik}\nend timeout"],
                   check=True, capture_output=True, text=True, timeout=zaman + 30)


def _excel_betigi(yol: str, ara_adim: str) -> str:
    """Aynı adlı açık kitap kaydedilmişse kapatır (yoksa Excel eski açık kopyayı kullanır); kaydedilmemişse
    kullanıcının işini kaybetmemek için kapatmaz, hata verir. Kendi açtığı kitabı hata olsa da kapatır."""
    ad = os.path.basename(yol)
    return f'''tell application "Microsoft Excel"
  if (exists workbook "{ad}") then
    if saved of workbook "{ad}" then
      close workbook "{ad}" saving no
    else
      error "Excel'de aynı adlı kaydedilmemiş bir kitap açık ({ad}): kaydedip kapatın"
    end if
  end if
  set wbk to open workbook workbook file name (POSIX file "{yol}" as text)
  try
    calculate full
    {ara_adim}
  on error hata
    close wbk saving no
    error hata
  end try
  close wbk saving no
end tell'''


def _mac_kaydet(xlsx: str):
    _applescript(_excel_betigi(os.path.abspath(xlsx), "save wbk"))


def _mac_pdf(xlsx: str, pdf: str):
    adim = f'save as active sheet of wbk filename (POSIX file "{pdf}" as text) file format PDF file format'
    try:
        _applescript(_excel_betigi(os.path.abspath(xlsx), adim))
    except subprocess.CalledProcessError as hata:
        if "(-50)" not in (hata.stderr or ""):
            raise
        _applescript(_excel_betigi(os.path.abspath(xlsx), adim))  # Excel PDF kaydında aralıklı -50 verir


# ---------------------------------------------------------------- Windows Excel (COM)


def _excel_win(xlsx: str, pdf: str | None = None, dispatch=None):
    """Excel'i görünmez açar, tam hesaplatır; pdf verilirse PDF'e aktarır, yoksa kaydeder.
    Yalnız açtığı kitabı kapatır ve kendi başlattığı Excel örneğinden çıkar."""
    if dispatch is None:
        from win32com.client import DispatchEx  # noqa: PLC0415
        dispatch = DispatchEx
    uygulama = dispatch("Excel.Application")
    uygulama.Visible = False
    uygulama.DisplayAlerts = False
    kitap = None
    try:
        kitap = uygulama.Workbooks.Open(os.path.abspath(xlsx))
        uygulama.CalculateFull()
        if pdf:
            kitap.ExportAsFixedFormat(0, os.path.abspath(pdf))
        else:
            kitap.Save()
    except (OSError, MotorHatasi):
        raise
    except Exception as hata:  # pywintypes.com_error OSError'dan türemez; sıradaki motora geçilebilsin
        raise MotorHatasi(f"Excel (COM): {hata}") from hata
    finally:
        if kitap is not None:
            kitap.Close(False)
        uygulama.Quit()


def _win_kaydet(xlsx: str):
    _excel_win(xlsx)


def _win_pdf(xlsx: str, pdf: str):
    _excel_win(xlsx, pdf=pdf)


# ---------------------------------------------------------------- LibreOffice (headless)

_XCU = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" \
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse">\
<value>0</value></prop></item>
<item oor:path="/org.openoffice.Setup/L10N"><prop oor:name="ooSetupSystemLocale" oor:op="fuse">\
<value>tr-TR</value></prop></item>
</oor:items>
"""


def _lo_donustur(xlsx: str, hedef_tur: str, uzanti: str) -> tuple[Path, tempfile.TemporaryDirectory]:
    """Yalıtılmış profille dönüştürür (kullanıcının açık LibreOffice oturumuna takılmaz, formüller her zaman
    yeniden hesaplanır, sayılar Türkçe yerel ayarla gösterilir). Döndürür: üretilen dosya ve silinmesi gereken geçici klasör."""
    soffice = _soffice()
    if not soffice:
        raise MotorHatasi("LibreOffice (soffice) bulunamadı")
    gecici = tempfile.TemporaryDirectory(prefix="xlkit-lo-")
    kok = Path(gecici.name)
    (kok / "profil" / "user").mkdir(parents=True)
    (kok / "profil" / "user" / "registrymodifications.xcu").write_text(_XCU, encoding="utf-8")
    cikis = kok / "cikti"
    cikis.mkdir()
    subprocess.run([soffice, f"-env:UserInstallation={(kok / 'profil').as_uri()}", "--headless", "--norestore",
                    "--convert-to", hedef_tur, "--outdir", str(cikis), os.path.abspath(xlsx)],
                   check=True, capture_output=True, timeout=LO_ZAMAN_ASIMI)
    uretilen = cikis / f"{Path(xlsx).stem}.{uzanti}"
    if not uretilen.is_file():
        gecici.cleanup()
        raise MotorHatasi(f"LibreOffice {uzanti} üretmedi")
    return uretilen, gecici


def _lo_kaydet(xlsx: str):
    uretilen, gecici = _lo_donustur(xlsx, "xlsx:Calc MS Excel 2007 XML", "xlsx")  # liste argümanı: tırnaksız
    try:
        shutil.move(str(uretilen), os.path.abspath(xlsx))
    finally:
        gecici.cleanup()


def _lo_pdf(xlsx: str, pdf: str):
    uretilen, gecici = _lo_donustur(xlsx, "pdf", "pdf")
    try:
        shutil.move(str(uretilen), os.path.abspath(pdf))
    finally:
        gecici.cleanup()


# ---------------------------------------------------------------- ortak seçim

_KAYDET = {"excel-mac": "_mac_kaydet", "excel-win": "_win_kaydet", "libreoffice": "_lo_kaydet"}
_PDF = {"excel-mac": "_mac_pdf", "excel-win": "_win_pdf", "libreoffice": "_lo_pdf"}
_HATALAR = (MotorHatasi, subprocess.CalledProcessError, subprocess.TimeoutExpired, OSError)


def _calistir(tablo: dict, args: tuple, motor: str | None, atla: tuple = ()) -> str | None:
    if motor is not None:
        if motor not in MOTORLAR:
            raise ValueError(f"Bilinmeyen motor: {motor}. Seçenekler: {', '.join(MOTORLAR)}")
        getattr(sys.modules[__name__], tablo[motor])(*args)
        return motor
    for ad in bulunan_motorlar():
        if ad in atla:
            warnings.warn("LibreOffice otomatik kaydetmede kullanılmaz (biçimi bozabilir); dosya hesaplanmamış "
                          f"değerlerle bırakıldı. Gerekirse motor=\"{ad}\" ile açıkça seçin", UserWarning, stacklevel=3)
            continue
        try:
            getattr(sys.modules[__name__], tablo[ad])(*args)
            return ad
        except _HATALAR as hata:
            warnings.warn(f"{ad} çalışmadı ({hata}); sıradaki motor deneniyor", UserWarning, stacklevel=3)
    warnings.warn(YOK_UYARISI, UserWarning, stacklevel=3)
    return None


def hesapla_kaydet(xlsx: str, motor: str | None = None) -> str | None:
    """Formülleri hesaplatıp dosyayı kaydeder. Döndürür: kullanılan motor ya da None."""
    return _calistir(_KAYDET, (xlsx,), motor, atla=("libreoffice",))


def pdf_al(xlsx: str, pdf: str, motor: str | None = None) -> str | None:
    """Kitabı PDF'e aktarır; özgün dosyaya dokunmaz. Döndürür: kullanılan motor ya da None."""
    return _calistir(_PDF, (xlsx, pdf), motor)
