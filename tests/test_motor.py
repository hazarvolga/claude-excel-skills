"""Motor katmanı: sırayla Excel (Mac/Windows) ve LibreOffice; sessiz geçiş yok."""
from __future__ import annotations

import subprocess
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

import motor
import xlkit as xk


def _kitap(yol: Path) -> Path:
    yol.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws["A1"] = 2
    ws["A2"] = "=A1*21"
    wb.save(yol)
    return yol


def test_otomatik_sira_ve_gecis(monkeypatch, tmp_path):
    monkeypatch.setattr(motor, "bulunan_motorlar", lambda: ["excel-mac", "libreoffice"])

    def mac_hata(xlsx, pdf):
        raise subprocess.CalledProcessError(1, "osascript")

    def lo_pdf(xlsx, pdf):
        Path(pdf).write_bytes(b"%PDF-1.4")

    monkeypatch.setattr(motor, "_mac_pdf", mac_hata)
    monkeypatch.setattr(motor, "_lo_pdf", lo_pdf)
    x = _kitap(tmp_path / "k.xlsx")
    with pytest.warns(UserWarning, match="excel-mac"):
        assert motor.pdf_al(str(x), str(tmp_path / "k.pdf")) == "libreoffice"


def test_motor_yoksa_uyari_ve_bos(monkeypatch, tmp_path):
    monkeypatch.setattr(motor, "bulunan_motorlar", lambda: [])
    x = _kitap(tmp_path / "k.xlsx")
    with pytest.warns(UserWarning, match="görsel denetim yapılamadı"):
        assert xk.onizle(str(x), str(tmp_path / "on")) == []
    with pytest.warns(UserWarning, match="görsel denetim yapılamadı"):
        assert xk.excel_ile_kaydet(str(x)) is None


def test_libreoffice_otomatik_kaydetmede_atlanir(monkeypatch, tmp_path):
    monkeypatch.setattr(motor, "bulunan_motorlar", lambda: ["libreoffice"])
    monkeypatch.setattr(motor, "_lo_kaydet", lambda xlsx: pytest.fail("LibreOffice özgün dosyaya yazmamalı"))
    x = _kitap(tmp_path / "k.xlsx")
    once = x.read_bytes()
    with pytest.warns(UserWarning, match="LibreOffice"):
        assert motor.hesapla_kaydet(str(x)) is None
    assert x.read_bytes() == once


def test_libreoffice_komutu_yalitilmis_profil(monkeypatch, tmp_path):
    yakalanan = {}

    def sahte_run(args, **kw):
        yakalanan["args"] = args
        profil = next(a for a in args if a.startswith("-env:UserInstallation=file://"))
        xcu = Path(profil.split("=", 1)[1].replace("file://", "")) / "user" / "registrymodifications.xcu"
        yakalanan["xcu"] = xcu.read_text(encoding="utf-8")
        cikti = Path(args[args.index("--outdir") + 1]) / "k.pdf"
        cikti.write_bytes(b"%PDF-1.4")
        return subprocess.CompletedProcess(args, 0, "", "")

    monkeypatch.setattr(motor, "_soffice", lambda: "/yok/soffice")
    monkeypatch.setattr(motor.subprocess, "run", sahte_run)
    x = _kitap(tmp_path / "k.xlsx")
    motor._lo_pdf(str(x), str(tmp_path / "sonuc.pdf"))
    assert "--headless" in yakalanan["args"]
    assert "OOXMLRecalcMode" in yakalanan["xcu"] and "<value>0</value>" in yakalanan["xcu"]
    assert "ooSetupSystemLocale" in yakalanan["xcu"] and "<value>tr-TR</value>" in yakalanan["xcu"]
    assert (tmp_path / "sonuc.pdf").is_file()


class _SahteCom:
    def __init__(self, open_hata=False):
        self.cagrilar = []
        self.open_hata = open_hata
        self.Visible = True
        self.DisplayAlerts = True
        sahte = self

        class _Kitap:
            def ExportAsFixedFormat(self, tur, yol):
                sahte.cagrilar.append(("ExportAsFixedFormat", tur, yol))

            def Save(self):
                sahte.cagrilar.append(("Save",))

            def Close(self, kaydet):
                sahte.cagrilar.append(("Close", kaydet))

        class _Kitaplar:
            def Open(self, yol):
                sahte.cagrilar.append(("Open", yol))
                if sahte.open_hata:
                    raise OSError("açılamadı")
                return _Kitap()

        self.Workbooks = _Kitaplar()

    def CalculateFull(self):
        self.cagrilar.append(("CalculateFull",))

    def Quit(self):
        self.cagrilar.append(("Quit",))


def test_excel_win_cagri_sirasi(tmp_path):
    com = _SahteCom()
    motor._excel_win(str(tmp_path / "k.xlsx"), pdf=str(tmp_path / "k.pdf"), dispatch=lambda ad: com)
    assert [c[0] for c in com.cagrilar] == ["Open", "CalculateFull", "ExportAsFixedFormat", "Close", "Quit"]
    assert com.cagrilar[2][1:] == (0, str(tmp_path / "k.pdf"))
    assert com.cagrilar[3] == ("Close", False)

    hatali = _SahteCom(open_hata=True)
    with pytest.raises(OSError):
        motor._excel_win(str(tmp_path / "k.xlsx"), dispatch=lambda ad: hatali)
    assert hatali.cagrilar[-1] == ("Quit",)


TURKCE = Path("Teklif 2026 ş") / "çıktı.xlsx"


@pytest.mark.skipif("libreoffice" not in motor.bulunan_motorlar(), reason="LibreOffice yok")
def test_libreoffice_gercek_hesaplama_turkce_yol(tmp_path):
    x = _kitap(tmp_path / TURKCE)
    assert motor.hesapla_kaydet(str(x), motor="libreoffice") == "libreoffice"
    assert load_workbook(x, data_only=True).active["A2"].value == 42
    pdf = tmp_path / "sonuc.pdf"
    assert motor.pdf_al(str(x), str(pdf), motor="libreoffice") == "libreoffice"
    assert pdf.stat().st_size > 0


@pytest.mark.skipif("excel-mac" not in motor.bulunan_motorlar(), reason="macOS Excel yok")
def test_excel_mac_onizle_turkce_yol(tmp_path):
    x = _kitap(tmp_path / TURKCE)
    assert len(xk.onizle(str(x), str(tmp_path / "on"), motor="excel-mac")) >= 1
