"""Temalar: yedi tema, geriye uyum, kontrast, yeni alanlar."""
from __future__ import annotations

from dataclasses import astuple

import xlkit as xk

YEDI = {"kurumsal", "minimal", "gece", "mimari", "teknik", "baski", "sunum"}

# 1.1.0 değerleri: bu temaların çıktısı değişmemeli
ESKI = {
    "kurumsal": ("kurumsal", "1B2A3A", "F5F7FA", "EEF1F5", "C9D1DB", "E6EAF0", "6B7785", "1F6FEB",
                 ("1F6FEB", "1B2A3A", "8A97A6", "C9D1DB", "5FA8D3"), "2E7D32", "C98A12", "B3261E",
                 "Calibri Light", "Calibri", "FFF8E1"),
    "minimal": ("minimal", "202124", "FAFAFA", "F1F3F4", "DADCE0", "ECEDEF", "70757A", "0F9D58",
                ("0F9D58", "202124", "9AA0A6", "DADCE0", "7BC8A4"), "2E7D32", "C98A12", "B3261E",
                "Calibri", "Calibri", "FFF8E1"),
    "gece": ("gece", "E8EAED", "1E2227", "262B31", "3C434C", "30363D", "9AA4AF", "F2A33A",
             ("F2A33A", "E8EAED", "7D8790", "4F5862", "E07B53"), "2E7D32", "C98A12", "B3261E",
             "Calibri", "Calibri", "FFF8E1"),
}


def _parlaklik(hex_renk: str) -> float:
    def kanal(c: float) -> float:
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = (int(hex_renk[i:i + 2], 16) / 255 for i in (0, 2, 4))
    return 0.2126 * kanal(r) + 0.7152 * kanal(g) + 0.0722 * kanal(b)


def kontrast(a: str, b: str) -> float:
    la, lb = sorted((_parlaklik(a), _parlaklik(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


def test_yedi_tema_var():
    assert set(xk.TEMALAR) == YEDI


def test_eski_temalar_degismedi():
    for ad, beklenen in ESKI.items():
        t = xk.TEMALAR[ad]
        assert astuple(t)[:len(beklenen)] == beklenen, ad
        assert (t.yazi_olcegi, t.rakam_yazi, t.renksiz) == (1.0, None, False), ad


def test_vurgu_kontrasti():
    # Yalnız yeni temalar: kurumsal ve minimal 1.0'dan beri kuralın altında, geriye uyum için değişmez
    for ad in YEDI - set(ESKI):
        t = xk.TEMALAR[ad]
        assert kontrast(t.vurgu, t.kagit) >= 4.5, ad


def test_yeni_tema_alanlari():
    assert xk.TEMALAR["teknik"].rakam_yazi == "Consolas"
    assert xk.TEMALAR["baski"].renksiz is True
    assert xk.TEMALAR["sunum"].yazi_olcegi == 1.3
    assert xk.TEMALAR["mimari"].baslik_yazi == "Cambria"


def test_ozel_vurgu_yeni_temada():
    assert xk.tema("mimari", vurgu="C8102E").seriler[0] == "C8102E"


# ---------------------------------------------------------------- yazı ölçeği ve rakam yazısı

from openpyxl import Workbook  # noqa: E402
from openpyxl.chart import BarChart, Reference  # noqa: E402


def _kpi(ad: str):
    ws = Workbook().active
    xk.kpi_karti(ws, xk.tema(ad), "B2", 4, "Toplam ciro", "=1000", "sayi")
    return ws


def test_sunum_olcegi():
    ws = _kpi("sunum")
    assert ws["B3"].font.sz == 28.5
    assert ws.row_dimensions[3].height == 46.8
    ws = _kpi("kurumsal")
    assert ws["B3"].font.sz == 22
    assert ws.row_dimensions[3].height == 36


def test_teknik_rakam_yazisi():
    ws = _kpi("teknik")
    assert ws["B3"].font.name == "Consolas"
    assert ws["B2"].font.name == "Calibri"


def _grafik_baslik_boyutu(ad: str) -> int:
    ws = Workbook().active
    for i in range(1, 4):
        ws.append([f"k{i}", i])
    ch = BarChart()
    ch.add_data(Reference(ws, min_col=2, min_row=1, max_row=3))
    xk.grafik_stili(ch, xk.tema(ad), "Başlık")
    return ch.title.tx.rich.p[0].r[0].rPr.sz


def test_grafik_metni_olcekli():
    assert _grafik_baslik_boyutu("sunum") == 1430
    assert _grafik_baslik_boyutu("kurumsal") == 1100


def test_sunum_eksen_etiketi_olceklenmez():
    # Ölçekli eksen yazısı dar grafikte ay etiketlerini iki satıra böler (görsel denetimde görüldü)
    ws = Workbook().active
    for i in range(1, 4):
        ws.append([f"k{i}", i, i * 2])
    ch = BarChart()
    ch.add_data(Reference(ws, min_col=2, max_col=3, min_row=1, max_row=3))
    xk.grafik_stili(ch, xk.tema("sunum"), "Başlık")
    assert ch.x_axis.txPr.p[0].pPr.defRPr.sz == 900
    assert ch.legend.txPr.p[0].pPr.defRPr.sz == 1170


# ---------------------------------------------------------------- renksiz baskı

from openpyxl.chart import DoughnutChart, LineChart  # noqa: E402


def _renk(dolgu) -> str | None:
    """openpyxl solidFill'i bellekte bazen düz metin, bazen ColorChoice tutar."""
    if dolgu is None or isinstance(dolgu, str):
        return dolgu
    rgb = dolgu.srgbClr
    return rgb if isinstance(rgb, str) or rgb is None else rgb.val


def _grafik(tur, t, seri_sayisi=2, satir=3):
    ws = Workbook().active
    for i in range(1, satir + 1):
        ws.append([f"k{i}"] + [i * (j + 1) for j in range(seri_sayisi)])
    ch = tur()
    ch.add_data(Reference(ws, min_col=2, max_col=1 + seri_sayisi, min_row=1, max_row=satir))
    xk.grafik_stili(ch, t, "Başlık")
    return ch


def test_renksiz_cubuk_desenli():
    ch = _grafik(BarChart, xk.tema("baski"))
    s0, s1 = (s.graphicalProperties for s in ch.series)
    assert _renk(s0.solidFill) == "000000"
    assert s1.pattFill.prst == "ltUpDiag"
    assert s1.solidFill is None


def test_renksiz_cizgi_kesikli():
    ch = _grafik(LineChart, xk.tema("baski"))
    s0, s1 = (s.graphicalProperties.ln for s in ch.series)
    assert s1.prstDash == "dash"
    assert _renk(s0.solidFill) == _renk(s1.solidFill) == "000000"


def test_renksiz_yedi_dilim_doner():
    ch = _grafik(DoughnutChart, xk.tema("baski"), seri_sayisi=1, satir=7)
    dilimler = ch.series[0].dPt
    assert _renk(dilimler[6].graphicalProperties.solidFill) == "000000"   # DESENLER[0] düz mürekkep
    assert dilimler[1].graphicalProperties.pattFill.prst == xk.DESENLER[1]


def test_renksiz_ozel_vurgu():
    ch = _grafik(BarChart, xk.tema("baski", vurgu="C8102E"))
    assert _renk(ch.series[0].graphicalProperties.solidFill) == "000000"


def test_renksiz_kpi_alt_bicim_renksiz():
    ws = Workbook().active
    xk.kpi_karti(ws, xk.tema("baski"), "B2", 4, "Ciro", "=1", "sayi", alt_formul="=0.05")
    kod = ws["B4"].number_format
    assert "[Color" not in kod and "[Red]" not in kod and "▲" in kod


def _durum_font_renkleri(ad):
    ws = Workbook().active
    xk.durum_renkleri(ws, xk.tema(ad), "B2:B5", "$B2>1", "$B2>0", "$B2<0")
    return [k.dxf.font.color.rgb[-6:] for k in ws.conditional_formatting["B2:B5"]]


def test_renksiz_durum_murekkep():
    assert set(_durum_font_renkleri("baski")) == {"000000"}
    assert _durum_font_renkleri("kurumsal")[0] == "2E7D32"


def test_renksiz_desenli_dilim_etiketi_zeminli():
    # Görsel denetim: sık desen üstündeki yüzde etiketi okunmuyordu
    ch = _grafik(DoughnutChart, xk.tema("baski"), seri_sayisi=1, satir=5)
    etiketler = ch.series[0].dLbls.dLbl
    assert _renk(etiketler[1].spPr.solidFill) == "FFFFFF"
    assert etiketler[0].spPr is None


# ---------------------------------------------------------------- antet

import subprocess  # noqa: E402
import sys  # noqa: E402

import pytest  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

from conftest import SKILLS  # noqa: E402

ALANLAR = [("Proje", "Örnek"), ("Belge no", "PNO-001"), ("Rev.", "A"), ("Tarih", "=TODAY()")]


def test_antet_hucreleri():
    ws = Workbook().active
    xk.antet(ws, xk.tema("teknik"), "B2:I4", ALANLAR)
    assert ws["B2"].value == "PROJE" and ws["D2"].value == "BELGE NO"
    assert ws["B3"].value == "Örnek" and ws["H3"].value == "=TODAY()"
    birlesik = {str(r) for r in ws.merged_cells.ranges}
    assert {"B2:C2", "D2:E2", "B3:C4", "H3:I4"} <= birlesik
    assert ws["B2"].border.left.style == "medium"
    assert ws["B3"].font.name == "Consolas"


def test_antet_dar_alan_hata():
    with pytest.raises(ValueError, match="dar"):
        xk.antet(Workbook().active, xk.tema("teknik"), "B2:C4", ALANLAR)


def _pano_degerleri(tmp_path, tema):
    cikti = tmp_path / f"{tema}.xlsx"
    subprocess.run([sys.executable, "-I", str(SKILLS / "excel-dashboard/scripts/ornek_pano.py"), str(cikti), tema],
                   check=True, capture_output=True, timeout=120)
    ws = load_workbook(cikti)["Pano"]
    return {c.value for satir in ws.iter_rows() for c in satir if isinstance(c.value, str)}


def test_ornek_pano_teknik_antet(tmp_path):
    assert "BELGE NO" in _pano_degerleri(tmp_path, "teknik")
    assert "BELGE NO" not in _pano_degerleri(tmp_path, "kurumsal")


# ---------------------------------------------------------------- son inceleme düzeltmeleri


def test_renksiz_yigilmis_cubuk_etiketi_okunur():
    ws = Workbook().active
    for i in range(1, 4):
        ws.append([f"k{i}", i, i * 2, i * 3])
    ch = BarChart(grouping="stacked", overlap=100)
    ch.add_data(Reference(ws, min_col=2, max_col=4, min_row=1, max_row=3))
    xk.grafik_stili(ch, xk.tema("baski"), "Yığılmış")
    duz, desenli = ch.series[0].dLbls, ch.series[1].dLbls
    assert _renk(duz.txPr.p[0].pPr.defRPr.solidFill) == "FFFFFF"       # düz mürekkep çubukta kâğıt rengi
    assert _renk(desenli.txPr.p[0].pPr.defRPr.solidFill) == "000000"   # desenli çubukta mürekkep
    assert _renk(desenli.spPr.solidFill) == "FFFFFF"                   # ve kâğıt zemin


def test_antet_tek_hucre_birlestirmez():
    ws = Workbook().active
    xk.antet(ws, xk.tema("teknik"), "Q2:T3", ALANLAR)
    assert all(":" in str(r) for r in ws.merged_cells.ranges)
    assert ws["R3"].value == "PNO-001"


def test_turkce_buyuk_harf():
    assert xk.buyuk_harf("Teklif tarihi") == "TEKLİF TARİHİ"
    assert xk.buyuk_harf("Bütçe kullanımı") == "BÜTÇE KULLANIMI"
    ws = Workbook().active
    xk.antet(ws, xk.tema("teknik"), "B2:C3", [("Teklif no", "T-1")])
    assert ws["B2"].value == "TEKLİF NO"
    ws = Workbook().active
    xk.kpi_karti(ws, xk.tema("kurumsal"), "B2", 4, "Bitiş tarihi", "=1")
    assert ws["B2"].value == "BİTİŞ TARİHİ"
