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
