"""Birim fiyat teklif cetvelini KİK012 e-formu ve Kamu İhale Genel Tebliği 16.5 / 39 kurallarına göre denetler.

Şablon değildir: cetveli kurmaz, kurulmuş ve HESAPLATILMIŞ bir kitabı (xk.excel_ile_kaydet ya da LibreOffice
kopyası) okur ve uymayan her noktayı satırıyla listeler. Kurallar 20.11.2026 itibarıyla doğrulandı
(bkz. references/kamu-ekap.md).

Kullanım:
    python3 scripts/kik012_denetim.py CETVEL.xlsx --idare IDARE.json [--bf-hane 2] [--sayfa Cetvel]
IDARE.json: idarenin A grubu, [[sıra, iş kalemi no, adı ve kısa açıklaması, birimi, miktarı], ...]
Çıkış kodu: 0 uygun, 1 hata var, 2 kullanım hatası.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from decimal import ROUND_HALF_UP, Decimal

from openpyxl import load_workbook

BASLIK = "BİRİM FİYAT TEKLİF CETVELİ"
IKN = "İhale Kayıt Numarası (İKN)"
TOPLAM = "TOPLAM TUTAR (KDV Hariç)"
ALT_BILGI = "KİK012 Birim Fiyat Teklif Cetveli (Mal/Yapım/Danışmanlık)"
SUTUNLAR = ("Sıra No", "İş Kalemi No", "Adı ve Kısa Açıklaması", "Birimi", "Miktarı", "Teklif Edilen Birim Fiyat",
            "Tutarı")
DIPNOTLAR = (
    "a) Kısmi teklif verilmesine izin verilen ihalelerde kısımlar ihale dokümanına ve ikincil ihale mevzuatına uygun "
    "olarak düzenlenecektir. b) Konsorsiyumların ihaleye teklif verebileceklerinin öngörülmesi halinde, bu cetvel işin "
    "uzmanlık gerektiren bölümleri esas alınarak idarece ayrı ayrı düzenlenecektir.",
    "Bu sütun idarece hazırlanacaktır.",
    "Bu sütun isteklilerce doldurulacaktır.",
    "Bu sütuna sadece yapım işleri ihalelerinde yer verilecektir.",
    "Hesaplama yapılırken her bir iş kaleminin miktar sütunundaki rakam ile teklif edilen birim fiyat çarpılarak o iş "
    "kalemine ait teklif tutarı bulunacaktır.",
)
HASSASIYET = Decimal("1e-9")  # Excel'in kayan nokta gürültüsü için


def yuvarla_kurus(deger: Decimal) -> Decimal:
    """Tebliğ 39.2 / 16.5.2: en yakın iki ondalık; yarım kuruş ve üstü bir kuruşa tamamlanır."""
    return deger.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def _ondalik(deger) -> Decimal:
    return Decimal(str(deger)).quantize(HASSASIYET).normalize()


def _hane(deger: Decimal) -> int:
    us = deger.normalize().as_tuple().exponent
    return max(0, -us) if isinstance(us, int) else 0


def _gosterilen_hane(bicim: str) -> int:
    """Sayı biçiminin ilk bölümünde ondalık ayıracından sonra gösterilen en çok hane (0 ve #)."""
    ilk = re.sub(r'"[^"]*"|\[[^\]]*\]', "", (bicim or "General").split(";")[0])
    if ilk.strip().lower() == "general":
        return 10
    if "." not in ilk:
        return 0
    return len(re.match(r"[0#]*", ilk.split(".", 1)[1]).group(0))


def _metin(deger) -> str:
    return re.sub(r"\s+", " ", str(deger or "")).strip()


def _sayfa_bul(wb, ad: str | None):
    if ad:
        return wb[ad]
    for ws in wb:
        if any(_metin(c.value) == BASLIK for r in ws.iter_rows(max_row=10) for c in r):
            return ws
    return wb.worksheets[0]


def _baslik_satiri(ws):
    for r in ws.iter_rows():
        degerler = [_metin(c.value) for c in r]
        if "Sıra No" in degerler and any(d.startswith("Teklif Edilen Birim Fiyat") for d in degerler):
            sutun = {}
            for c in r:
                for ad in SUTUNLAR:
                    if _metin(c.value).startswith(ad):
                        sutun.setdefault(ad, c.column)
            return r[0].row, sutun
    return None, {}


def _veri_satirlari(ws, bas: int):
    satirlar = []
    for r in range(bas + 1, ws.max_row + 1):
        ilk = " ".join(_metin(ws.cell(r, c).value) for c in range(1, 8))
        if TOPLAM in ilk or ilk.startswith("TOPLAM"):
            return satirlar, r
        if any(ws.cell(r, c).value not in (None, "") for c in range(1, ws.max_column + 1)):
            satirlar.append(r)
    return satirlar, None


def _kontrol_yapisi(ws, metinler, hatalar):
    if BASLIK not in metinler:
        hatalar.append(f"başlık '{BASLIK}' yok")
    if not any(m.startswith(IKN) for m in metinler):
        hatalar.append(f"'{IKN}:' satırı yok")
    for i, dipnot in enumerate(DIPNOTLAR, 1):
        if not any(dipnot in m for m in metinler):
            hatalar.append(f"dipnot {i} eksik ya da farklı: '{dipnot[:50]}…'")
    altlar = [ws.oddFooter.left.text, ws.oddFooter.center.text, ws.oddFooter.right.text]
    if not any(ALT_BILGI in _metin(a) for a in altlar) and not any(ALT_BILGI in m for m in metinler):
        hatalar.append(f"alt bilgi '{ALT_BILGI}' yok")


def _kontrol_satir(no, ws, r, sutun, idare_kalem, bf_hane, hatalar) -> Decimal:
    a_grubu = [ws.cell(r, sutun[s]).value for s in SUTUNLAR[:5]]
    if idare_kalem is not None:
        beklenen = [str(idare_kalem[0]), _metin(idare_kalem[1]), _metin(idare_kalem[2]), _metin(idare_kalem[3]),
                    _ondalik(idare_kalem[4])]
        gercek = [str(a_grubu[0]).removesuffix(".0"), _metin(a_grubu[1]), _metin(a_grubu[2]), _metin(a_grubu[3]),
                  _ondalik(a_grubu[4]) if a_grubu[4] not in (None, "") else None]
        for ad, b, g in zip(SUTUNLAR[:5], beklenen, gercek):
            if b != g:
                hatalar.append(f"satır {no}: A grubu '{ad}' idarenin listesinden farklı ({g!r} ≠ {b!r})")
    bf_hucre, tutar_hucre = ws.cell(r, sutun["Teklif Edilen Birim Fiyat"]), ws.cell(r, sutun["Tutarı"])
    if bf_hucre.value in (None, ""):
        hatalar.append(f"satır {no}: birim fiyat boş (teklif eksik)")
        return Decimal(0)
    bf, miktar = _ondalik(bf_hucre.value), _ondalik(a_grubu[4] or 0)
    if _hane(bf) > bf_hane:
        hatalar.append(f"satır {no}: birim fiyat {_hane(bf)} hane, seçilen {bf_hane} hane (değer yuvarlanmamış)")
    if _hane(bf) > _gosterilen_hane(bf_hucre.number_format):
        hatalar.append(f"satır {no}: birim fiyat biçimi {_gosterilen_hane(bf_hucre.number_format)} hane gösteriyor, "
                       f"değer {_hane(bf)} haneli (görünen ≠ saklanan)")
    if tutar_hucre.value in (None, ""):
        hatalar.append(f"satır {no}: tutar boş")
        return Decimal(0)
    tutar = _ondalik(tutar_hucre.value)
    if (miktar * bf).quantize(HASSASIYET).normalize() != tutar:
        hatalar.append(f"satır {no}: tutar {tutar} ≠ miktar × birim fiyat {miktar * bf} (aritmetik hata)")
    if _hane(tutar) > _gosterilen_hane(tutar_hucre.number_format):
        hatalar.append(f"satır {no}: tutar biçimi {_gosterilen_hane(tutar_hucre.number_format)} hane gösteriyor, "
                       f"değer {_hane(tutar)} haneli (görünen ≠ saklanan)")
    return tutar


def denetle(xlsx: str, idare_json: str | None = None, bf_hane: int = 2, sayfa: str | None = None) -> list[str]:
    """Hata listesi döndürür; boş liste = uygun."""
    hatalar: list[str] = []
    wb, formul = load_workbook(xlsx, data_only=True), load_workbook(xlsx)
    ws = _sayfa_bul(wb, sayfa)
    wsf = formul[ws.title]
    metinler = [_metin(c.value) for r in ws.iter_rows() for c in r if isinstance(c.value, str)]
    _kontrol_yapisi(ws, metinler, hatalar)
    bas, sutun = _baslik_satiri(ws)
    if bas is None or len(sutun) < len(SUTUNLAR):
        eksik = [s for s in SUTUNLAR if s not in sutun]
        return hatalar + [f"sütun başlıkları eksik: {', '.join(eksik) or 'başlık satırı bulunamadı'}"]
    satirlar, toplam_satiri = _veri_satirlari(ws, bas)
    for r in satirlar:
        for c in (sutun["Teklif Edilen Birim Fiyat"], sutun["Tutarı"]):
            if isinstance(wsf.cell(r, c).value, str) and wsf.cell(r, c).value.startswith("=") and ws.cell(r, c).value is None:
                return hatalar + ["kitap hesaplanmamış: önce xk.excel_ile_kaydet() ya da LibreOffice kopyasında hesaplatın"]
    idare = json.loads(open(idare_json, encoding="utf-8").read()) if idare_json else None
    if idare is not None and len(idare) != len(satirlar):
        hatalar.append(f"A grubu kalem sayısı {len(satirlar)}, idarenin listesinde {len(idare)}")
    toplam = sum((_kontrol_satir(i + 1, ws, r, sutun, idare[i] if idare and i < len(idare) else None, bf_hane, hatalar)
                  for i, r in enumerate(satirlar)), Decimal(0))
    if toplam_satiri is None:
        hatalar.append(f"'{TOPLAM}' satırı yok")
    else:
        yazilan = ws.cell(toplam_satiri, sutun["Tutarı"]).value
        if yazilan in (None, "") or _ondalik(yazilan) != yuvarla_kurus(toplam).normalize():
            hatalar.append(f"toplam {yazilan} ≠ Σ tutar {toplam} kuruşa yuvarlanmış {yuvarla_kurus(toplam)}")
        for r in range(bas + 1, toplam_satiri):
            satir_metni = " ".join(_metin(ws.cell(r, c).value) for c in range(1, ws.max_column + 1))
            if re.search(r"\bKDV\b|İskonto|Iskonto|kâr\b", satir_metni):
                hatalar.append(f"satır {r}: cetvelde KDV / iskonto / kâr satırı olmamalı")
    return hatalar


def main(argv: list[str]) -> int:
    p = argparse.ArgumentParser(description="KİK012 birim fiyat teklif cetveli denetimi")
    p.add_argument("xlsx")
    p.add_argument("--idare", help="İdarenin A grubu listesi (JSON)")
    p.add_argument("--bf-hane", type=int, default=2)
    p.add_argument("--sayfa")
    a = p.parse_args(argv)
    try:
        hatalar = denetle(a.xlsx, a.idare, a.bf_hane, a.sayfa)
    except (OSError, KeyError, ValueError) as hata:
        print(f"Kullanım hatası: {hata}")
        return 2
    for h in hatalar:
        print(f"HATA: {h}")
    print("UYGUN: KİK012 ve Tebliğ 16.5 / 39 denetiminden geçti" if not hatalar else f"{len(hatalar)} hata")
    return 1 if hatalar else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
