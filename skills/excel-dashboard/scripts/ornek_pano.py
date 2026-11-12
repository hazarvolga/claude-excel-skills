"""Örnek pano: Veri sayfası + Pano sayfası (4 KPI kartı, 4 grafik, filtre, tablo).

Yeni bir pano için bu dosyayı kopyalayın, VERI ve pano_kur() içindeki yerleşimi değiştirin.
Kullanım:
    uv run --no-project --with openpyxl --with pymupdf --with pillow python -I ornek_pano.py CIKTI.xlsx [tema] [--onizle DIZIN]
"""
import os
import sys
from datetime import date

sys.path.insert(0, os.path.expanduser("~/.claude/skills/excel-design-system/scripts"))
import xlkit as xk  # noqa: E402
from openpyxl import Workbook  # noqa: E402
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference  # noqa: E402
from openpyxl.styles import Alignment  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402

AYLAR = [date(2026, m, 1) for m in range(1, 13)]
PLAN = [820, 860, 900, 940, 980, 1010, 1040, 1060, 1080, 1100, 1120, 1150]
GERCEK = [790, 875, 920, 905, 1010, 1050, 1015, 1090, 1120, 1085, 1160, 1190]
BOLGE = [("Marmara", 4200), ("İç Anadolu", 2950), ("Ege", 2300), ("Akdeniz", 1650), ("Karadeniz", 1110)]
URUN = [("Elektronik", 5200, 4400), ("Mobilya", 3100, 2700), ("Tekstil", 1900, 1750), ("Kırtasiye", 1200, 900), ("Ambalaj", 800, 720)]


def veri_kur(wb, t):
    ws = wb.create_sheet("Veri")
    xk.zemin(ws, t)
    xk.sayfa_basligi(ws, t, "Veri", "Panodaki tüm grafikler bu sayfadaki tablolardan beslenir.", genislik=6)
    xk.tablo_basligi(ws, t, 6, 2, [("Ay", 12), ("Plan", 12), ("Gerçek", 12), ("Sapma", 12)])
    for i, (ay, p, g) in enumerate(zip(AYLAR, PLAN, GERCEK)):
        r = 7 + i
        ws.cell(r, 2, ay).number_format = xk.BICIM["ay"]
        ws.cell(r, 3, p * 1000).number_format = xk.BICIM["tl"]
        ws.cell(r, 4, g * 1000).number_format = xk.BICIM["tl"]
        s = ws.cell(r, 5, f"=D{r}/C{r}-1")
        s.number_format = xk.BICIM["etiket_delta"]  # grafik etiketi olarak da kullanılır: ondalıksız
    xk.tablo(ws, "B6:E18", "AylikSatis")
    xk.negatif_kirmizi(ws, t, "E7:E18")

    xk.tablo_basligi(ws, t, 6, 7, [("Bölge", 14), ("Ciro", 14)])
    for i, (b, v) in enumerate(BOLGE):
        ws.cell(7 + i, 7, b)
        ws.cell(7 + i, 8, v * 1000).number_format = xk.BICIM["tl"]
    xk.tablo(ws, "G6:H11", "Bolge")

    xk.tablo_basligi(ws, t, 6, 10, [("Ürün", 14), ("Bütçe", 14), ("Harcanan", 14), ("Kullanım", 12)])
    for i, (u, b, h) in enumerate(URUN):
        r = 7 + i
        ws.cell(r, 10, u)
        ws.cell(r, 11, b * 1000).number_format = xk.BICIM["tl"]
        ws.cell(r, 12, h * 1000).number_format = xk.BICIM["tl"]
        ws.cell(r, 13, f"=L{r}/K{r}").number_format = xk.BICIM["yuzde"]
    xk.tablo(ws, "J6:M11", "Urun")
    # Grafik verisi bloğu: etiket kaynak hücre biçimini gösterir ve Türkçe Excel etiket biçimini yerel
    # sözdizimiyle okuyabilir. Ölçeği formül yapar, biçim yalnız güvenli karakterlerden (0" B" -> 5200 B).
    xk.tablo_basligi(ws, t, 6, 15, [("Bütçe", 10), ("Harcanan", 10)])
    for i in range(5):
        for j, kaynak in enumerate(("K", "L")):
            c = ws.cell(7 + i, 15 + j, f"=ROUND({kaynak}{7 + i}/1000,0)")
            c.number_format = xk.BICIM["etiket_bin"]
    ws.cell(5, 15, "Grafik verisi (bin ₺)").font = xk.yazi(t, 9, renk=t.ara)
    xk.veri_cubugu(ws, t, "M7:M11")
    xk.dondur(ws, "A7")
    xk.baski(ws, t, alt_metin="Örnek pano · xlkit")
    return ws


def pano_kur(wb, t):
    ws = wb.active
    ws.title = "Pano"
    xk.izgara(ws, 23)
    xk.zemin(ws, t, "A1:Z60")
    xk.sayfa_basligi(ws, t, "Satış ve Maliyet Panosu", "2026 · tüm değerler Veri sayfasından formülle gelir",
                     hucre="B2", genislik=23)
    # Filtre satırı (açılır liste + seçime bağlı formül)
    ws["B6"].value = "Bölge"
    ws["B6"].font = xk.yazi(t, 9, kalin=True, renk=t.ara)
    ws.merge_cells("D6:H6")
    ws["D6"].value = "Marmara"
    xk.girdi_hucresi(ws, t, "D6:H6")
    xk.acilir_liste(ws, "D6", "=Veri!$G$7:$G$11", "Bölge", "Bölge seçin; sağdaki değer güncellenir")
    # Sayıyı TEXT() ile metne çevirmeyin: biçim dizesi Excel diline göre değişir. Ayrı hücre + number_format.
    ws["J6"].value = "Seçilen bölge cirosu"
    ws["J6"].font = xk.yazi(t, 9, kalin=True, renk=t.ara)
    ws.merge_cells("N6:R6")
    ws["N6"].value = "=INDEX(Veri!$H$7:$H$11,MATCH(D6,Veri!$G$7:$G$11,0))"
    ws["N6"].number_format = xk.BICIM["tl"]
    ws["N6"].font = xk.yazi(t, 11, kalin=True)
    ws["N6"].alignment = Alignment(horizontal="left", vertical="center")

    # KPI satırı: 4 kart x 5 sütun, aralarında 1 boş sütun (B-F, H-L, N-R, T-X)
    kartlar = [
        ("Toplam ciro", "=SUM(Veri!D7:D18)", "milyon", "=SUM(Veri!D7:D18)/SUM(Veri!C7:C18)-1", "plana göre"),
        ("Aylık ortalama", "=AVERAGE(Veri!D7:D18)", "tl", "=Veri!D18/Veri!D17-1", "son ay değişimi"),
        ("Planı tutan ay", '=COUNTIF(Veri!E7:E18,">=0")&" / 12"', "General", None, "gerçek ≥ plan"),
        ("Bütçe kullanımı", "=SUM(Veri!L7:L11)/SUM(Veri!K7:K11)", "yuzde", None, "harcanan / bütçe"),
    ]
    for i, (etiket, formul, bicim, alt, alt_metin) in enumerate(kartlar):
        xk.kpi_karti(ws, t, f"{get_column_letter(2 + i * 6)}8", 5, etiket, formul, bicim,
                     alt_formul=alt, alt_metin=alt_metin)

    # Grafik ızgarası: satır 12-29 iki grafik (15 + 7 sütun), satır 31-48 iki grafik (11 + 11)
    line = LineChart()
    line.add_data(Reference(wb["Veri"], min_col=3, max_col=4, min_row=6, max_row=18), titles_from_data=True)
    line.set_categories(Reference(wb["Veri"], min_col=2, min_row=7, max_row=18))
    xk.grafik_stili(line, t, "Aylık ciro: plan ve gerçek", deger="milyon", ay_ekseni=True)
    xk.yerlestir(ws, line, xk.hucre_araligi(2, 12, 15, 18))

    pasta = DoughnutChart(holeSize=55)
    pasta.add_data(Reference(wb["Veri"], min_col=8, min_row=6, max_row=11), titles_from_data=True)
    pasta.set_categories(Reference(wb["Veri"], min_col=7, min_row=7, max_row=11))
    xk.grafik_stili(pasta, t, "Bölge payı")
    xk.yerlestir(ws, pasta, xk.hucre_araligi(18, 12, 7, 18))

    bar = BarChart()
    bar.add_data(Reference(wb["Veri"], min_col=15, max_col=16, min_row=6, max_row=11), titles_from_data=True)
    bar.set_categories(Reference(wb["Veri"], min_col=10, min_row=7, max_row=11))
    xk.grafik_stili(bar, t, "Ürün grubu: bütçe ve harcanan (bin ₺)", deger="sayi", yatay=True)
    xk.yerlestir(ws, bar, xk.hucre_araligi(2, 31, 11, 18))

    sapma = BarChart(type="col")
    sapma.add_data(Reference(wb["Veri"], min_col=5, min_row=6, max_row=18), titles_from_data=True)
    sapma.set_categories(Reference(wb["Veri"], min_col=2, min_row=7, max_row=18))
    xk.grafik_stili(sapma, t, "Aylık sapma (gerçek / plan − 1)", deger="delta", ay_ekseni="ay_kisa")
    xk.yerlestir(ws, sapma, xk.hucre_araligi(14, 31, 11, 18))

    xk.baski(ws, t, alt_metin="Örnek pano · xlkit", alan="A1:Y49", tek_sayfa=True)
    return ws


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    cikti = os.path.abspath(args[0] if args else "ornek-pano.xlsx")
    t = xk.tema(args[1] if len(args) > 1 else "kurumsal")
    wb = Workbook()
    veri_kur(wb, t)
    pano = pano_kur(wb, t)
    xk.sekme(pano, t.vurgu)
    xk.sekme(wb["Veri"], t.murekkep)
    sorunlar = xk.cakisma_denetimi(pano)
    if sorunlar:
        raise SystemExit("Yerleşim sorunu:\n" + "\n".join(sorunlar))
    wb.save(cikti)
    if "--onizle" in sys.argv:
        dizin = sys.argv[sys.argv.index("--onizle") + 1]
        for p in xk.onizle(cikti, dizin):
            print(p)
    print(cikti)


if __name__ == "__main__":
    main()
