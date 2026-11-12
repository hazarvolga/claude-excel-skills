---
name: excel-charts
description: "Use when adding, restyling or fixing charts inside an Excel workbook with openpyxl: picking the chart type, colors, data labels, axis and number formats, placing charts so they do not overlap cells or each other, or when the user says excel grafik, grafik ekle, sütun/çizgi/pasta/halka grafik, grafikler taşıyor, etiketler bozuk, grafik profesyonel görünsün. Uses xlkit.py from excel-design-system."
metadata:
  version: "1.0.0"
  scope: global
---

# Excel Grafikleri

Önce `excel-design-system` skill'ini okuyun: tema ve `xlkit` oradadır. Bu skill grafik **seçimi**,
**tek tasarım kuralı** ve **yerleşimi** anlatır.

```python
import os, sys
sys.path.insert(0, os.path.expanduser("~/.claude/skills/excel-design-system/scripts"))
import xlkit as xk
from openpyxl.chart import BarChart, LineChart, DoughnutChart, Reference
```

## 1. Grafik türünü soruya göre seçin

| Soru | Grafik | openpyxl |
|---|---|---|
| Zaman içinde nasıl değişti? | Çizgi (≤ 4 seri) | `LineChart` |
| Kategoriler arası karşılaştırma, kısa ad | Sütun | `BarChart()` |
| Kategoriler arası karşılaştırma, uzun ad ya da > 7 kategori | Yatay çubuk, büyükten küçüğe | `BarChart()` + `yatay=True` |
| Plan ve gerçek | Kümelenmiş sütun/çubuk ya da iki çizgi | |
| Bütünün parçaları, ≤ 5 dilim | Halka | `DoughnutChart(holeSize=55)` |
| Parçaların zaman içindeki payı | Yığılmış sütun | `BarChart(grouping="stacked", overlap=100)` |
| Artı/eksi sapma | Sütun, tek renk, sıfır çizgisi | `BarChart()` + `deger="delta"` |
| İki ölçü arasındaki ilişki | Dağılım | `ScatterChart` |

Kaçının: 3B grafik, 5'ten fazla dilimli pasta, çift eksenli karışık grafik (okuyucu hangi eksenin hangi
seriye ait olduğunu karıştırır; iki ayrı grafik yapın), radar.

## 2. Tek tasarım kuralı: `xk.grafik_stili`

```python
ch = LineChart()
ch.add_data(Reference(veri, min_col=3, max_col=4, min_row=6, max_row=18), titles_from_data=True)
ch.set_categories(Reference(veri, min_col=2, min_row=7, max_row=18))
xk.grafik_stili(ch, t, "Aylık ciro: plan ve gerçek", deger="milyon", ay_ekseni=True)
xk.yerlestir(ws, ch, "B12:P29")
```

Bu çağrı şunları uygular, elle tekrar etmeyin:
- Başlık çizim alanının dışında, 11 pt kalın; başlık bir **cümle** olsun ("Aylık ciro: plan ve gerçek"),
  tek kelime değil.
- Lejant altta; tek serili grafikte lejant yok (başlık yeterli).
- Yalnız yatay açık ızgara, tick yok, ince eksen çizgisi, zemin kâğıt rengi.
- Seri renkleri tema sırasıyla: vurgu, mürekkep, gri. Negatif çubuk beyaza dönmez.
- Değer ekseni biçimi `deger` ile (`milyon` → `1,2 M`). Ay ekseni Türkçe (`ay` → `Oca 26`,
  dar grafikte `ay_kisa` → `Oca`).
- Çubukta değer etiketi; yığılmışta etiket ortada, koyu dilimde beyaz.
- Halka/pastada yüzde etiketi, dilimler arası kâğıt rengi çizgi; etiket rengi dilime göre (koyu dilimde beyaz).
- Çizim alanı zemini temaya göre: açık temada beyaz, koyu temada ve halkada kâğıt rengi.

## 3. Veri etiketi biçimi (ölçülmüş, en sık hata)

Excel veri etiketinde **kaynak hücrenin biçimini** gösterir; etikete yazılan kod yok sayılır. Türkçe
Excel'de bazı kitaplarda bu biçimi **yerel sözdizimiyle** okur: virgül ondalık, nokta binlik, `General`
harf harf kod. Aynı biçim hücrede ve eksende doğru görünürken etiket bozulur:

| Hücre biçimi | Etikette görünen |
|---|---|
| `#,##0.0,, "M"` | `5200000,00,, M` |
| `+0.0%` | `+04%` |
| `General" M"` | `5190001900r1l` |
| `+0%`, `0" B"` | `+4%`, `5200 B` (doğru) |

Kural: **etiketin okuduğu hücrelerde biçim kodu yalnız `0 % + -` ve tırnaklı metinden oluşur; ölçeği
formül yapar.** Ondalık gereken değer etikette değil eksende gösterilir (eksen biçimi doğru çizilir).

```python
# Grafik verisi bloğu: tablo tam tutarı gösterir, grafik bu bloktan beslenir
for i in range(5):
    c = veri.cell(7 + i, 15, f"=ROUND(K{7 + i}/1000,0)")
    c.number_format = xk.BICIM["etiket_bin"]            # 0" B"  -> 5200 B
sapma_hucresi.number_format = xk.BICIM["etiket_delta"]  # +0%;-0%;0%
xk.grafik_stili(ch, t, "Ürün grubu: bütçe ve harcanan (bin ₺)", deger="sayi", yatay=True)
```

Ayrıntı ve diğer ölçülmüş tuzaklar: [references/openpyxl-grafik-tuzaklari.md](references/openpyxl-grafik-tuzaklari.md)

## 4. Yerleşim: grafik hiçbir şeyin üstüne binmez

- Panoda her grafik **hücre aralığına** oturtulur: `xk.yerlestir(ws, ch, "B12:P29")`. Sütun genişliği
  değişince grafik de uyar. `add_chart(ch, "B12")` + `width/height` cm vermeyin: farklı ekranda taşar.
- Var olan bir kitapta tek hücreye bağlı grafikleri `xk.bos_alana_sigdir(ws)` boş alana yayar.
- Kaydetmeden önce `xk.cakisma_denetimi(ws)` boş liste döndürmeli.
- En küçük okunur boyut: 7 sütun × 14 satır. 12 aylık sütun grafiği en az 11 sütun ister, yoksa ay
  etiketleri iki satıra kırılır (`ay_kisa` kullanın).

## 5. Görsel denetim zorunlu

openpyxl grafiği Excel'in nasıl çizeceğini bilmez. Her grafik değişikliğinden sonra:

```python
xk.excel_ile_kaydet(yol)
for png in xk.onizle(yol, "onizleme"): print(png)   # sonra her PNG'yi açıp bakın
```

Bakılacaklar: etiketler biçimli mi ve üst üste mi, eksen sayıları tekrar ediyor mu (`1 M, 1 M` = biçim
ondalığı yetersiz), başlık kesiliyor mu, ay adları kırılıyor mu, grafik sayfaya sığıyor mu.
Görüntüye bakmadan "grafik hazır" denmez.
