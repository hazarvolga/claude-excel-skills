"""kik012_denetim.py: birim fiyat teklif cetvelinin KİK012 ve Tebliğ 16.5 / 39 kurallarına uygunluğu."""
from __future__ import annotations

import json
import subprocess
import sys

import pytest
from openpyxl import Workbook

from conftest import SKILLS

BETIK = SKILLS / "excel-insaat-belgeleri" / "scripts" / "kik012_denetim.py"
sys.path.insert(0, str(BETIK.parent))

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
IDARE = [  # (sıra, iş kalemi no, ad, birim, miktar)
    [1, "ÖRN-16.050/03", "Beton santralinde üretilen beton", "m³", 120.5],
    [2, "ÖRN-21.001/01", "Ahşap kalıp", "m²", 840.25],
    [3, "ÖRN-23.014", "Nervürlü çelik", "ton", 14.125],
]


def _cetvel(tmp_path, fiyatlar=(3200.0, 515.5, 28750.75), tutar_bozuk=None, toplam=None, bicim="#,##0.00###",
            bf_bicim="#,##0.00", dipnotlar=DIPNOTLAR, alt_bilgi=True):
    """Hesaplanmış (değerli) bir cetvel kitabı yazar; denetim değerlerle çalışır."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Cetvel"
    ws["A1"] = "BİRİM FİYAT TEKLİF CETVELİ"
    ws["A2"] = "İhale Kayıt Numarası (İKN): 2026/123456"
    ws.append([])
    ws.append(["Sıra No", "İş Kalemi No", "Adı ve Kısa Açıklaması", "Birimi", "Miktarı",
               "Teklif Edilen Birim Fiyat", "Tutarı"])
    tutarlar = []
    for i, (kalem, bf) in enumerate(zip(IDARE, fiyatlar)):
        tutar = kalem[4] * bf if bf is not None else None
        if tutar_bozuk is not None and i == tutar_bozuk:
            tutar = round(tutar, 0)
        tutarlar.append(tutar or 0)
        ws.append(kalem + [bf, tutar])
        r = ws.max_row
        ws.cell(r, 6).number_format = bf_bicim
        ws.cell(r, 7).number_format = bicim
    ws.append(["TOPLAM TUTAR (KDV Hariç)", None, None, None, None, None,
               toplam if toplam is not None else round(sum(tutarlar) + 1e-9, 2)])
    ws.cell(ws.max_row, 7).number_format = "#,##0.00"
    for metin in dipnotlar:
        ws.append([metin])
    if alt_bilgi:
        ws.oddFooter.left.text = "KİK012 Birim Fiyat Teklif Cetveli (Mal/Yapım/Danışmanlık)"
    yol = tmp_path / "cetvel.xlsx"
    wb.save(yol)
    idare = tmp_path / "idare.json"
    idare.write_text(json.dumps(IDARE, ensure_ascii=False), encoding="utf-8")
    return yol, idare


def _denetle(yol, idare, bf_hane=2):
    import kik012_denetim
    return kik012_denetim.denetle(str(yol), str(idare), bf_hane=bf_hane)


def test_uygun_cetvel_hatasiz(tmp_path):
    assert _denetle(*_cetvel(tmp_path)) == []


def test_tutar_carpima_esit_degil(tmp_path):
    hatalar = _denetle(*_cetvel(tmp_path, tutar_bozuk=1))
    assert any("satır 2" in h and "tutar" in h for h in hatalar)


def test_toplam_yuvarlama_hatasi(tmp_path):
    hatalar = _denetle(*_cetvel(tmp_path, toplam=1.0))
    assert any("toplam" in h.lower() for h in hatalar)


def test_birim_fiyat_hanesi_fazla(tmp_path):
    hatalar = _denetle(*_cetvel(tmp_path, fiyatlar=(3200.123, 515.5, 28750.75)))
    assert any("birim fiyat" in h and "hane" in h for h in hatalar)


def test_gorunen_tutar_kesiliyor(tmp_path):
    # 120,5 × 3200 = 385600; 840,25 × 515,5 = 433148,875 → 3 hane; biçim 2 hane gösterirse görünen ≠ saklanan
    hatalar = _denetle(*_cetvel(tmp_path, bicim="#,##0.00"))
    assert any("biçim" in h for h in hatalar)


def test_idare_listesi_degismis(tmp_path):
    yol, idare = _cetvel(tmp_path)
    bozuk = [k[:] for k in IDARE]
    bozuk[2][4] = 14.0
    idare.write_text(json.dumps(bozuk, ensure_ascii=False), encoding="utf-8")
    assert any("A grubu" in h for h in _denetle(yol, idare))


def test_bos_birim_fiyat(tmp_path):
    hatalar = _denetle(*_cetvel(tmp_path, fiyatlar=(3200.0, None, 28750.75)))
    assert any("boş" in h for h in hatalar)


def test_dipnot_ve_alt_bilgi(tmp_path):
    hatalar = _denetle(*_cetvel(tmp_path, dipnotlar=DIPNOTLAR[:4], alt_bilgi=False))
    assert any("dipnot 5" in h for h in hatalar)
    assert any("alt bilgi" in h for h in hatalar)


def test_cli_cikis_kodu(tmp_path):
    yol, idare = _cetvel(tmp_path, tutar_bozuk=1)
    sonuc = subprocess.run([sys.executable, "-I", str(BETIK), str(yol), "--idare", str(idare)],
                           capture_output=True, text=True, timeout=60)
    assert sonuc.returncode == 1 and "HATA" in sonuc.stdout
    yol, idare = _cetvel(tmp_path)
    sonuc = subprocess.run([sys.executable, "-I", str(BETIK), str(yol), "--idare", str(idare)],
                           capture_output=True, text=True, timeout=60)
    assert sonuc.returncode == 0 and "UYGUN" in sonuc.stdout


@pytest.mark.parametrize("toplam_ham, beklenen", [(500815.414, 500815.41), (500815.41582, 500815.42),
                                                  (500815.415, 500815.42)])
def test_kurus_yuvarlama_tebligdeki_ornekler(toplam_ham, beklenen):
    import kik012_denetim
    from decimal import Decimal
    assert kik012_denetim.yuvarla_kurus(Decimal(str(toplam_ham))) == Decimal(str(beklenen))


def test_dipnot_isaretli_baslik_kabul(tmp_path):
    yol, idare = _cetvel(tmp_path)
    from openpyxl import load_workbook
    wb = load_workbook(yol)
    wb["Cetvel"]["A1"] = "BİRİM FİYAT TEKLİF CETVELİ (1)"
    wb.save(yol)
    assert _denetle(yol, idare) == []
