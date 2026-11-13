"""xlkit: openpyxl üstünde profesyonel Excel tasarım yardımcıları.

Üç skill (excel-design-system, excel-charts, excel-dashboard) bu tek dosyayı kullanır.
Kullanım:
    import sys; sys.path.insert(0, os.path.expanduser("~/.claude/skills/excel-design-system/scripts"))
    import xlkit as xk
    t = xk.tema("kurumsal")

Çalıştırma: uv run --no-project --with openpyxl --with pymupdf --with pillow python -I betik.py
"""
from __future__ import annotations

import io
import os
import subprocess
from dataclasses import dataclass, field, replace

from openpyxl.chart.axis import ChartLines
from openpyxl.chart.data_source import NumFmt
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, RegularTextRun
from openpyxl.drawing.text import Font as DFont
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

# ---------------------------------------------------------------- tema


@dataclass(frozen=True)
class Tema:
    ad: str
    murekkep: str          # ana metin, başlık bandı
    kagit: str             # sayfa zemini (KPI kartı, bölüm zemini)
    yuzey: str             # formül / salt okunur hücre zemini
    cizgi: str             # ince kenarlık, eksen
    izgara: str            # grafik yatay ızgara (çizgiden açık)
    ara: str               # ikincil metin, açıklama
    vurgu: str             # tek vurgu rengi (marka)
    seriler: tuple = ()    # grafik serileri sırası: vurgu önce
    iyi: str = "2E7D32"    # durum renkleri anlam taşır: temayla değişmez
    uyari: str = "C98A12"
    kotu: str = "B3261E"
    baslik_yazi: str = "Georgia"
    govde_yazi: str = "Calibri"
    girdi: str = "FFF8E1"  # kullanıcının dolduracağı hücre zemini


TEMALAR = {
    "kurumsal": Tema("kurumsal", "1B2A3A", "F5F7FA", "EEF1F5", "C9D1DB", "E6EAF0", "6B7785", "1F6FEB",
                     seriler=("1F6FEB", "1B2A3A", "8A97A6", "C9D1DB", "5FA8D3"), baslik_yazi="Calibri Light"),
    "minimal": Tema("minimal", "202124", "FAFAFA", "F1F3F4", "DADCE0", "ECEDEF", "70757A", "0F9D58",
                    seriler=("0F9D58", "202124", "9AA0A6", "DADCE0", "7BC8A4"), baslik_yazi="Calibri"),
    "gece": Tema("gece", "E8EAED", "1E2227", "262B31", "3C434C", "30363D", "9AA4AF", "F2A33A",
                 seriler=("F2A33A", "E8EAED", "7D8790", "4F5862", "E07B53"), baslik_yazi="Calibri"),
}


def tema(ad: str = "kurumsal", **degisiklik) -> Tema:
    """Hazır temayı döndürür; marka rengi için tema("kurumsal", vurgu="C8102E", seriler=(...))."""
    if ad not in TEMALAR:
        raise ValueError(f"Bilinmeyen tema: {ad}. Seçenekler: {', '.join(TEMALAR)}")
    t = TEMALAR[ad]
    if "vurgu" in degisiklik and "seriler" not in degisiklik:
        degisiklik["seriler"] = (degisiklik["vurgu"],) + t.seriler[1:]
    return replace(t, **degisiklik) if degisiklik else t


# ---------------------------------------------------------------- sayı biçimleri (Türkçe Excel'de de doğru görünür)

BICIM = {
    "sayi": "#,##0",
    "sayi1": "#,##0.0",
    "sayi2": "#,##0.00",
    "yuzde": "0%",
    "yuzde1": "0.0%",
    "tl": '#,##0 "₺"',
    "tl2": '#,##0.00 "₺"',
    "bin": '#,##0, "B"',
    "milyon": '#,##0.0,, "M"',
    "tarih": "dd.mm.yyyy",
    "ay": "[$-41F]mmm yy",       # Türkçe ay adı, Excel dili ne olursa olsun
    "ay_kisa": "[$-41F]mmm",      # tek yıllık dar grafik ekseni
    "ay_uzun": "[$-41F]mmmm yyyy",
    "delta": '+0.0%;-0.0%;0.0%',  # KPI değişim
    "delta_ok": '[Color10]▲ 0.0%;[Red]▼ 0.0%;0.0%',
    # Grafik veri etiketleri için güvenli biçimler (bkz. etiket_bicimi): yalnız 0 % + - ve "metin".
    # "etiket_bin" / "etiket_milyon" hücrede ÖLÇEKLENMİŞ TAM SAYI ister: =ROUND(K7/1000,0) gibi.
    "etiket_bin": '0" B"',
    "etiket_milyon": '0" M"',
    "etiket_sayi": '0',
    "etiket_yuzde": '0%',
    "etiket_delta": '+0%;-0%;0%',
}


def etiket_bicimi(anahtar: str) -> str:
    """Grafik verisi hücre biçimi: sıfırlar gizlenir (üçüncü bölüm boş), etikette de görünmez.

    Ölçülen davranış (Excel for Mac 16, Türkçe, PDF çıktısı, Kasım 2026):
    - Etiket kaynak hücrenin biçimini gösterir; etikete yazılan numFmt kodu yok sayılır.
    - Bazı kitaplarda Excel etiket biçimini TÜRKÇE YEREL sözdizimiyle okur: virgül ondalık, nokta binlik,
      "General" harf harf kod. "#,##0.0,, M" -> "5200000,00,, M", "+0.0%" -> "+04%",
      "General M" -> "5190001900r1l". Aynı biçim hücrede ve eksende doğru görünür.
    - Güvenli olanlar her denemede doğru çıktı: yalnız 0, %, +, - ve tırnaklı metin ("+0%", '0" B"').
    Bu yüzden etiketin okuduğu hücrelerde ölçeği formül yapsın, biçim yalnız güvenli karakterlerden olsun:
        ws["O7"] = "=ROUND(K7/1000,0)"; ws["O7"].number_format = xk.BICIM["etiket_bin"]   # 5200 B
    Ondalık gereken değeri etikette değil eksende gösterin (eksen biçimi doğru çizilir).
    """
    temel = BICIM.get(anahtar, anahtar)
    return temel if ";" in temel else f"{temel};-{temel};;"


# ---------------------------------------------------------------- hücre stilleri


def ince(renk: str) -> Side:
    return Side(style="thin", color=renk)


def yazi(t: Tema, boyut=10, kalin=False, renk=None, baslik=False, italik=False) -> Font:
    return Font(name=t.baslik_yazi if baslik else t.govde_yazi, sz=boyut, b=kalin, i=italik, color=renk or t.murekkep)


def dolgu(renk: str) -> PatternFill:
    """Düz dolgu. Hücrede de koşullu biçimde de çalışır: hücre dolgusu rengi fgColor'dan, koşullu biçim (dxf)
    dolgusu rengi bgColor'dan okunur. Yalnız fgColor yazılırsa koşullu biçim dolgusu Excel'de görünmez
    (LibreOffice'te görünür), bu yüzden ikisi de yazılır."""
    return PatternFill("solid", fgColor=renk, bgColor=renk)


def zemin(ws, t: Tema, alan: str = "A1:Z80"):
    """Izgarayı kapatır ve alanı kâğıt rengine boyar (koyu temada zorunlu)."""
    ws.sheet_view.showGridLines = False
    if t.kagit.upper() in ("FFFFFF", "FAFAFA"):
        return
    for satir in ws[alan]:
        for c in satir:
            if c.__class__.__name__ != "MergedCell":
                c.fill = dolgu(t.kagit)


def sayfa_basligi(ws, t: Tema, metin: str, alt: str = "", hucre: str = "B2", genislik: int = 12):
    """Sayfa başlığı + alt satır. Başlık serif/ince, alt başlık ikincil renkte."""
    c = ws[hucre]
    c.value = metin
    c.font = yazi(t, 20, baslik=True)
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[c.row].height = 32
    if alt:
        a = ws.cell(c.row + 1, c.column, alt)
        a.font = yazi(t, 10, renk=t.ara)
    # İnce vurgu çizgisi
    for col in range(c.column, c.column + genislik):
        ws.cell(c.row + 2, col).border = Border(top=Side(style="medium", color=t.vurgu))


def bant(ws, t: Tema, alan: str, metin: str, koyu=True):
    """Bölüm başlık bandı (birleşik hücre)."""
    ws.merge_cells(alan)
    c = ws[alan.split(":")[0]]
    c.value = metin
    c.font = yazi(t, 11, kalin=True, renk=t.kagit if koyu else t.murekkep)
    c.fill = dolgu(t.murekkep if koyu else t.yuzey)
    c.alignment = Alignment(vertical="center", indent=1)


def tablo(ws, alan: str, ad: str, stil: str = "TableStyleLight1", seritli=True) -> Table:
    """Gerçek Excel Tablosu: filtre, yapısal başvuru, otomatik genişleme. Ad boşluksuz olmalı."""
    tb = Table(displayName=ad, ref=alan)
    tb.tableStyleInfo = TableStyleInfo(name=stil, showRowStripes=seritli, showColumnStripes=False)
    ws.add_table(tb)
    return tb


def tablo_basligi(ws, t: Tema, satir: int, sutun1: int, basliklar: list[tuple[str, float]]):
    """Başlık satırı + sütun genişlikleri. basliklar = [(metin, genişlik), ...]"""
    for i, (metin, gen) in enumerate(basliklar):
        c = ws.cell(satir, sutun1 + i, metin)
        c.font = yazi(t, 10, kalin=True, renk=t.kagit)
        c.fill = dolgu(t.murekkep)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(sutun1 + i)].width = gen
    ws.row_dimensions[satir].height = 30


def girdi_hucresi(ws, t: Tema, alan: str):
    """Kullanıcı girişi: açık sarı zemin + kilit kapalı (sayfa korumasında düzenlenebilir)."""
    from openpyxl.styles import Protection
    for satir in ws[alan]:
        for c in satir:
            c.fill = dolgu(t.girdi)
            c.protection = Protection(locked=False)


def acilir_liste(ws, alan: str, kaynak: str, baslik="Seçin", mesaj=""):
    """Veri doğrulamalı açılır liste. kaynak: "=Listeler!$A$2:$A$20" veya '"A,B,C"'."""
    dv = DataValidation(type="list", formula1=kaynak, allow_blank=True, showDropDown=False)
    dv.promptTitle, dv.prompt = baslik, mesaj
    dv.showInputMessage = bool(mesaj)
    ws.add_data_validation(dv)
    dv.add(alan)
    return dv


def durum_renkleri(ws, t: Tema, alan: str, iyi: str, uyari: str, kotu: str):
    """Formül tabanlı trafik ışığı. Formüller alanın sol üst hücresine göre yazılır, örn. iyi="$F5>=0.95"."""
    for formul, renk in ((iyi, t.iyi), (uyari, t.uyari), (kotu, t.kotu)):
        ws.conditional_formatting.add(alan, FormulaRule(formula=[formul], font=Font(color=renk, b=True)))


def veri_cubugu(ws, t: Tema, alan: str):
    ws.conditional_formatting.add(alan, DataBarRule(start_type="num", start_value=0, end_type="max", color=t.vurgu,
                                                    showValue=True))


def negatif_kirmizi(ws, t: Tema, alan: str):
    ws.conditional_formatting.add(alan, CellIsRule(operator="lessThan", formula=["0"], font=Font(color=t.kotu)))


# ---------------------------------------------------------------- KPI kartı


def kpi_karti(ws, t: Tema, sol_ust: str, sutun: int, etiket: str, formul: str, bicim: str = "sayi",
              alt_formul: str | None = None, alt_bicim: str = "delta_ok", alt_metin: str = ""):
    """3 satırlık kart: etiket / büyük değer / değişim. sutun = kart kaç sütun kaplar.

    Kart yalnız hücre stilidir (şekil değil): yazdırılır, sıralamada bozulmaz, formülle güncellenir.
    """
    c1, r1, _, _ = range_boundaries(sol_ust + ":" + sol_ust)
    c2 = c1 + sutun - 1
    harf1, harf2 = get_column_letter(c1), get_column_letter(c2)
    satirlar = [r1, r1 + 1, r1 + 2]
    for r in satirlar:
        ws.merge_cells(f"{harf1}{r}:{harf2}{r}")
    kenar = ince(t.cizgi)
    for r in satirlar:
        for col in range(c1, c2 + 1):
            c = ws.cell(r, col)
            c.fill = dolgu(t.yuzey)
            c.border = Border(left=Side(style="thick", color=t.vurgu) if col == c1 else None,
                              right=kenar if col == c2 else None,
                              top=kenar if r == r1 else None, bottom=kenar if r == r1 + 2 else None)
    e = ws.cell(r1, c1, etiket.upper())
    e.font = yazi(t, 9, kalin=True, renk=t.ara)
    e.alignment = Alignment(indent=1, vertical="bottom")
    d = ws.cell(r1 + 1, c1, formul)
    d.font = yazi(t, 22, kalin=True, baslik=True)
    d.number_format = BICIM.get(bicim, bicim)
    d.alignment = Alignment(horizontal="left", indent=1, vertical="center", shrink_to_fit=True)
    a = ws.cell(r1 + 2, c1, alt_formul if alt_formul else alt_metin)
    a.font = yazi(t, 9, renk=t.ara)
    a.alignment = Alignment(horizontal="left", indent=1, vertical="top")
    if alt_formul:
        a.number_format = BICIM.get(alt_bicim, alt_bicim)
    ws.row_dimensions[r1].height = 20
    ws.row_dimensions[r1 + 1].height = 36
    ws.row_dimensions[r1 + 2].height = 18


# ---------------------------------------------------------------- grafikler


def _metin(t: Tema, boyut: int, renk=None, kalin=False) -> RichText:
    oz = CharacterProperties(sz=boyut, b=kalin, solidFill=renk or t.ara, latin=DFont(typeface=t.govde_yazi))
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=oz), endParaRPr=oz)])


def _baslik(t: Tema, metin: str) -> Title:
    oz = CharacterProperties(sz=1100, b=True, solidFill=t.murekkep, latin=DFont(typeface=t.govde_yazi))
    p = Paragraph(pPr=ParagraphProperties(defRPr=oz), r=[RegularTextRun(rPr=oz, t=metin)])
    return Title(tx=Text(rich=RichText(p=[p])), overlay=False)


def grafik_stili(chart, t: Tema, baslik: str, deger: str = "sayi", etiket: bool = True, ay_ekseni: bool | str = False,
                 lejant: bool | None = None, yatay: bool = False, etiket_kodu: str | None = None):
    """Tek tasarım kuralı: başlık çizim alanı dışında, lejant altta, yalnız yatay açık ızgara,
    tick yok, seri renkleri temadan, sıfır etiketleri gizli, yığılmışta koyu çubuğa beyaz etiket.

    deger: BICIM anahtarı ("sayi", "tl", "milyon", "yuzde"...) yalnız DEĞER EKSENİ için. Veri etiketleri
    kaynak hücre biçimini gösterir: etiket biçimini hücreye verin (etiket_bicimi()); etiket_kodu verilmezse
    deger'den türetilir ve numFmt elemanı her zaman yazılır.
    ay_ekseni: True -> "ay" (Oca 26), ya da BICIM anahtarı ("ay_kisa").
    yatay: BarChart için yatay çubuk; ilk kategori üstte, değer ekseni altta.
    """
    tur = type(chart).__name__
    etiket_kodu = etiket_kodu or etiket_bicimi(deger)
    if tur == "BarChart" and yatay:
        chart.type = "bar"                         # BarChart(type=...) kurucuda yok sayılır
        chart.x_axis.scaling.orientation = "maxMin"
        chart.y_axis.crosses = "max"               # ters kategori ekseninde değer ekseni altta kalsın
    chart.title = _baslik(t, baslik)
    chart.style = None
    chart.roundedCorners = False
    tek_seri = len(chart.series) == 1 and tur != "PieChart" and tur != "DoughnutChart"
    if lejant is None:
        lejant = not tek_seri
    if lejant and chart.legend is not None:
        chart.legend.position = "b"
        chart.legend.overlay = False
        chart.legend.txPr = _metin(t, 900)
    else:
        chart.legend = None
    chart.graphical_properties = GraphicalProperties(solidFill=t.kagit, ln=LineProperties(solidFill=t.cizgi, w=6350))
    pasta = tur in ("PieChart", "DoughnutChart")
    # Çizim alanı varsayılanı beyazdır: koyu temada açık renkli seri ve etiketler kaybolur.
    cizim_zemini = t.kagit if (_koyu_mu(t.kagit) or pasta) else "FFFFFF"  # halkada beyaz kare kalmasın
    chart.plot_area.graphicalProperties = GraphicalProperties(solidFill=cizim_zemini, ln=LineProperties(noFill=True))
    if not pasta:
        alt_pay = 0.18 if lejant else 0.08
        chart.plot_area.layout = Layout(manualLayout=ManualLayout(x=0.03, y=0.14, w=0.94, h=1 - 0.14 - alt_pay - 0.04))
        for eksen in (chart.x_axis, chart.y_axis):
            eksen.delete = False
            eksen.txPr = _metin(t, 900)
            eksen.spPr = GraphicalProperties(ln=LineProperties(solidFill=t.cizgi, w=6350))
            eksen.majorTickMark = "none"
            eksen.minorTickMark = "none"
        chart.x_axis.majorGridlines = None
        chart.y_axis.majorGridlines = ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill=t.izgara, w=6350)))
        chart.y_axis.numFmt = NumFmt(formatCode=BICIM.get(deger, deger), sourceLinked=False)
        chart.x_axis.tickLblPos = "low"   # negatif değerde etiketler altta kalır
        if ay_ekseni:
            anahtar = "ay" if ay_ekseni is True else ay_ekseni
            chart.x_axis.numFmt = NumFmt(formatCode=BICIM[anahtar], sourceLinked=False)
    yigin = getattr(chart, "grouping", None) in ("stacked", "percentStacked")
    for i, seri in enumerate(chart.series):
        renk = t.seriler[i % len(t.seriler)]
        if tur in ("LineChart", "AreaChart", "ScatterChart"):
            seri.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=renk, w=25400))
            if tur == "AreaChart":
                seri.graphicalProperties.solidFill = renk
            seri.smooth = False
            if etiket and tur == "LineChart" and len(chart.series) == 1:
                seri.dLbls = DataLabelList(showVal=True, showSerName=False, showCatName=False, showLegendKey=False,
                                           numFmt=etiket_kodu, dLblPos="t")
                seri.dLbls.txPr = _metin(t, 800, t.murekkep)
        elif tur == "BarChart":
            seri.graphicalProperties = GraphicalProperties(solidFill=renk, ln=LineProperties(noFill=True))
            seri.invertIfNegative = False              # negatif çubuk beyaza dönmesin
            if etiket:
                koyu = _koyu_mu(renk)
                seri.dLbls = DataLabelList(showVal=True, showSerName=False, showCatName=False, showLegendKey=False,
                                           numFmt=etiket_kodu, dLblPos="ctr" if yigin else "outEnd")
                seri.dLbls.txPr = _metin(t, 800, "FFFFFF" if (yigin and koyu) else t.murekkep, yigin)
        elif pasta:
            for j in range(12):
                n = DataPoint(idx=j)
                n.graphicalProperties = GraphicalProperties(solidFill=t.seriler[j % len(t.seriler)],
                                                            ln=LineProperties(solidFill=t.kagit, w=12700))
                seri.dPt.append(n)
    if tur == "BarChart":
        chart.gapWidth = 60
        if getattr(chart, "grouping", None) == "clustered":
            chart.overlap = -10
        else:
            chart.overlap = 100
    if pasta:
        # Dilim başına etiket rengi: koyu dilimde beyaz, açık dilimde mürekkep.
        from openpyxl.chart.label import DataLabel
        ortak = dict(showPercent=True, showVal=False, showCatName=False, showSerName=False, showLegendKey=False)
        for seri in chart.series:
            etiketler = []
            for j in range(12):
                renk = t.seriler[j % len(t.seriler)]
                yazi_rengi = "FFFFFF" if _koyu_mu(renk) else "1B1F23"
                etiketler.append(DataLabel(idx=j, txPr=_metin(t, 900, yazi_rengi, True), **ortak))
            seri.dLbls = DataLabelList(dLbl=etiketler, **ortak)
    return chart


def _koyu_mu(renk: str) -> bool:
    r, g, b = (int(renk[i:i + 2], 16) for i in (0, 2, 4))
    return (0.299 * r + 0.587 * g + 0.114 * b) < 140


def yerlestir(ws, chart, alan: str):
    """Grafiği tam bir hücre aralığına oturtur ("B12:H28"). Sütun/satır boyutu değişince grafik de uyar."""
    c1, r1, c2, r2 = range_boundaries(alan)
    chart.anchor = TwoCellAnchor(_from=AnchorMarker(col=c1 - 1, row=r1 - 1), to=AnchorMarker(col=c2, row=r2))
    ws.add_chart(chart)
    return chart


def _dolu(ws, r, c):
    v = ws.cell(r, c).value
    return v is not None and str(v).strip() != ""


def bos_alana_sigdir(ws, sag_pay: int = 2, min_satir: int = 12, max_satir: int = 20):
    """add_chart(ch, "B5") ile tek hücreye bağlanmış grafikleri, sağında ve altında içerik başlayana
    kadar uzanan boş alana yayar. Var olan, elle tasarlanmamış kitaplar için."""
    en_sag = max((c.anchor._from.col + 1 for c in ws._charts if hasattr(c.anchor, "_from")), default=0)
    sag_sinir = max(ws.max_column, en_sag) + sag_pay
    for chart in ws._charts:
        if isinstance(chart.anchor, str):
            c0, r0, _, _ = range_boundaries(chart.anchor + ":" + chart.anchor)
        else:
            c0, r0 = chart.anchor._from.col + 1, chart.anchor._from.row + 1
        c1 = sag_sinir
        for c in range(c0, sag_sinir + 1):
            if any(_dolu(ws, r, c) for r in range(r0, r0 + 6)):
                c1 = c - 1
                break
        if c1 - c0 < 4:
            c1 = min(sag_sinir, c0 + 7)
        r1 = r0 + max_satir
        for r in range(r0 + 1, r0 + max_satir + 1):
            if any(_dolu(ws, r, c) for c in range(c0, c1 + 1)):
                r1 = r - 1
                break
        r1 = max(r1, r0 + min_satir)
        chart.anchor = TwoCellAnchor(_from=AnchorMarker(col=c0 - 1, row=r0 - 1), to=AnchorMarker(col=c1, row=r1))


def cakisma_denetimi(ws) -> list[str]:
    """Grafikler birbirinin ya da dolu hücrelerin üstüne biniyor mu? Boş liste = temiz."""
    kutular, sorunlar = [], []
    for ch in ws._charts:
        a = ch.anchor
        if not hasattr(a, "to"):
            sorunlar.append(f"{_ad(ch)}: tek hücre bağlantısı; yerlestir() ile aralık verin")
            continue
        kutu = (a._from.col + 1, a._from.row + 1, a.to.col, a.to.row)  # 1 tabanlı, kapsayıcı
        for ad, k in kutular:
            if not (kutu[2] < k[0] or k[2] < kutu[0] or kutu[3] < k[1] or k[3] < kutu[1]):
                sorunlar.append(f"{_ad(ch)} ile {ad} üst üste")
        for r in range(kutu[1], kutu[3] + 1):
            for c in range(kutu[0], kutu[2] + 1):
                if _dolu(ws, r, c):
                    sorunlar.append(f"{_ad(ch)} dolu hücrenin üstünde: {get_column_letter(c)}{r}")
                    break
            else:
                continue
            break
        kutular.append((_ad(ch), kutu))
    return sorunlar


def _ad(ch) -> str:
    try:
        return "'" + "".join(r.t for r in ch.title.tx.rich.p[0].r) + "'"
    except (AttributeError, IndexError, TypeError):
        return type(ch).__name__


# ---------------------------------------------------------------- pano ızgarası ve baskı


def izgara(ws, sutun_sayisi: int = 24, genislik: float = 4.6, kenar: float = 2.0):
    """Pano için eşit sütun ızgarası: A kenar boşluğu, B.. eşit sütunlar. 24 sütun = 2/3/4/6 bölünür."""
    ws.column_dimensions["A"].width = kenar
    for i in range(2, sutun_sayisi + 2):
        ws.column_dimensions[get_column_letter(i)].width = genislik
    ws.column_dimensions[get_column_letter(sutun_sayisi + 2)].width = kenar


def hucre_araligi(sutun1: int, satir1: int, sutun_say: int, satir_say: int) -> str:
    """Izgara koordinatından aralık: hucre_araligi(2, 8, 12, 16) -> 'B8:M23'."""
    return f"{get_column_letter(sutun1)}{satir1}:{get_column_letter(sutun1 + sutun_say - 1)}{satir1 + satir_say - 1}"


def baski(ws, t: Tema, ust_metin: str = "&A", alt_metin: str = "", yatay: bool = True, alan: str | None = None,
          tek_sayfa: bool = False):
    """A4 baskı. tek_sayfa=True panolar içindir (genişlik ve yükseklik bir sayfaya); tablolar için False."""
    ws.page_setup.orientation = "landscape" if yatay else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1 if tek_sayfa else 0
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.print_options.horizontalCentered = True
    if alan:
        ws.print_area = alan
    ws.oddHeader.left.text, ws.oddHeader.left.size = ust_metin, 8
    ws.oddFooter.right.text, ws.oddFooter.right.size = "&P / &N", 8
    if alt_metin:
        ws.oddFooter.left.text, ws.oddFooter.left.size = alt_metin, 8


def sekme(ws, renk: str):
    ws.sheet_properties.tabColor = renk


def gezinme(wb, t: Tema, menu_sayfa: str, hucre: str = "A1", metin: str = "← Menü"):
    """Her sayfanın hucre'sine menüye dönüş bağlantısı koyar."""
    for ws in wb.worksheets:
        if ws.title == menu_sayfa:
            continue
        c = ws[hucre]
        c.value = metin
        c.hyperlink = f"#'{menu_sayfa}'!A1"
        c.font = Font(name=t.govde_yazi, sz=9, u="single", color=t.vurgu)


def dondur(ws, hucre: str):
    ws.freeze_panes = hucre


# ---------------------------------------------------------------- Excel ile hesapla, önizle


def _applescript(betik: str, zaman=300):
    subprocess.run(["osascript", "-e", f"with timeout of {zaman} seconds\n{betik}\nend timeout"], check=True)


def _excel_betigi(yol: str, ara_adim: str) -> str:
    """Aynı adlı açık kitabı önce kapatır (yoksa Excel eski açık kopyayı kullanır), hata olsa da kapatır."""
    ad = os.path.basename(yol)
    return f'''tell application "Microsoft Excel"
  try
    close workbook "{ad}" saving no
  end try
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


def excel_ile_kaydet(xlsx: str, temizle: bool = True):
    """macOS + Microsoft Excel: formülleri hesaplatıp kaydeder. openpyxl hesaplanmış değer yazmaz;
    Quick Look, telefon, bazı okuyucular (exceljs) bu adım olmadan boş veya bozuk görür.
    Excel kaydederken kullanıcı/şirket adını ve tam dosya yolunu yazar; temizle=True bunları siler."""
    _applescript(_excel_betigi(os.path.abspath(xlsx), "save wbk"))
    if temizle:
        meta_temizle(xlsx)


def meta_temizle(xlsx: str, yazar: str = ""):
    """Paylaşılacak dosyadan kişisel izleri siler: dc:creator, cp:lastModifiedBy, Company, Manager ve
    workbook.xml içindeki absPath (bilgisayardaki tam klasör yolu). Diğer parçalara dokunmaz."""
    import re
    import shutil
    import tempfile
    import zipfile

    desenler = {
        "docProps/core.xml": [(r"(<dc:creator>)[^<]*(</dc:creator>)", rf"\g<1>{yazar}\g<2>"),
                              (r"(<cp:lastModifiedBy>)[^<]*(</cp:lastModifiedBy>)", rf"\g<1>{yazar}\g<2>")],
        "docProps/app.xml": [(r"<Company>[^<]*</Company>", "<Company></Company>"),
                             (r"<Manager>[^<]*</Manager>", "<Manager></Manager>")],
        "xl/workbook.xml": [(r"<mc:AlternateContent[^>]*>(?:(?!</mc:AlternateContent>).)*absPath(?:(?!</mc:AlternateContent>).)*</mc:AlternateContent>", ""),
                            (r"<x15ac:absPath[^>]*/>", "")],
    }
    fd, gecici = tempfile.mkstemp(suffix=".xlsx", dir=os.path.dirname(os.path.abspath(xlsx)))
    os.close(fd)
    with zipfile.ZipFile(xlsx) as kaynak, zipfile.ZipFile(gecici, "w", zipfile.ZIP_DEFLATED) as hedef:
        for oge in kaynak.infolist():
            veri = kaynak.read(oge.filename)
            if oge.filename in desenler:
                metin = veri.decode("utf-8")
                for desen, yerine in desenler[oge.filename]:
                    metin = re.sub(desen, yerine, metin, flags=re.S)
                veri = metin.encode("utf-8")
            hedef.writestr(oge, veri)
    shutil.move(gecici, xlsx)


def onizle(xlsx: str, cikti_dizin: str, dpi: int = 150, max_genislik: int = 2000) -> list[str]:
    """Kitabı Excel ile PDF'e kaydeder, her sayfayı PNG çizer (pymupdf + Pillow). Döndürür: PNG yolları.
    Görsel denetim bu görüntülerle yapılır: taşma, kırpılma, üst üste binme, okunmayan etiket."""
    import pymupdf
    from PIL import Image

    os.makedirs(cikti_dizin, exist_ok=True)
    pdf = os.path.join(os.path.abspath(cikti_dizin), "onizleme.pdf")
    for eski in os.listdir(cikti_dizin):
        if eski == "onizleme.pdf" or (eski.startswith("sayfa-") and eski.endswith(".png")):
            os.remove(os.path.join(cikti_dizin, eski))
    adim = f'save as active sheet of wbk filename (POSIX file "{pdf}" as text) file format PDF file format'
    try:
        _applescript(_excel_betigi(os.path.abspath(xlsx), adim))
    except subprocess.CalledProcessError:
        _applescript(_excel_betigi(os.path.abspath(xlsx), adim))  # Excel PDF kaydında aralıklı -50 verir
    if not os.path.exists(pdf) or os.path.getmtime(pdf) < os.path.getmtime(xlsx):
        raise RuntimeError("PDF üretilmedi ya da kitaptan eski: Excel izin penceresini kontrol edin")
    cikanlar = []
    doc = pymupdf.open(pdf)
    for i, sayfa in enumerate(doc, 1):
        pix = sayfa.get_pixmap(dpi=dpi)
        img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
        img = _kirp(img)
        if img.width > max_genislik:
            img = img.resize((max_genislik, round(img.height * max_genislik / img.width)), Image.LANCZOS)
        hedef = os.path.join(cikti_dizin, f"sayfa-{i:02d}.png")
        img.save(hedef)
        cikanlar.append(hedef)
    return cikanlar


def _kirp(img, pay=24):
    gri = img.convert("L").point(lambda v: 255 if v < 250 else 0)
    kutu = gri.getbbox()
    if not kutu:
        return img
    x0, y0, x1, y1 = kutu
    return img.crop((max(0, x0 - pay), max(0, y0 - pay), min(img.width, x1 + pay), min(img.height, y1 + pay)))


def turkce_bugun() -> str:
    """Türkçe Excel'de TEXT(TODAY();"dd.mm.yyyy") '00.mm.2026' gösterir (yyyy/dd harfleri yerelleşir).
    Dilden bağımsız tarih metni için bu formül parçasını kullanın."""
    return 'TEXT(DAY(TODAY()),"00")&"."&TEXT(MONTH(TODAY()),"00")&"."&YEAR(TODAY())'
