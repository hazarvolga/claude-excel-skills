---
name: excel-dashboard
description: "Use when building a dashboard, report cover, KPI panel or management summary sheet inside Excel: laying out KPI cards, charts and filters on a grid, wiring dropdown filters to formulas, adding navigation between sheets, or when the user says excel dashboard, excel pano, yönetim paneli, KPI kartı, rapor sayfası, özet sayfa, maliyet paneli, proje paneli. Builds on excel-design-system (xlkit) and excel-charts; includes a runnable template."
metadata:
  version: "1.0.0"
  scope: global
---

# Excel Panosu

Önce `excel-design-system` (tema, xlkit) ve `excel-charts` (grafik kuralı) skill'lerini okuyun.

**Hazır şablon:** `~/.claude/skills/excel-dashboard/scripts/ornek_pano.py`. Veri sayfası + pano
(4 KPI kartı, filtre, 4 grafik). Yeni pano için kopyalayıp `veri_kur()` ve `pano_kur()` içini
değiştirin. Görüntüsü: [references/ornek-pano.png](references/ornek-pano.png)

```bash
uv run --no-project --with openpyxl --with pymupdf --with pillow python -I ornek_pano.py cikti.xlsx kurumsal --onizle onizleme
```

## 1. Önce soru, sonra yerleşim

Panoyu kim, hangi kararı vermek için okuyor? En önemli 3-5 sayı KPI kartı olur, en önemli soru en
büyük grafik olur. Kullanıcı söylemediyse tahmin etmeyin, bir cümleyle sorun.

## 2. Mimari: veri ayrı, pano formül

```
Veri  (Excel Tabloları, ham kayıtlar)   ← kullanıcı yalnız burayı doldurur
Hesap (isteğe bağlı: özet tablolar, SUMIFS, grafik verisi blokları)
Pano  (KPI kartları, grafikler, filtre) ← hiçbir sabit sayı yok, hepsi formül
```

- Pano üzerindeki her sayı formüldür (`=SUM(Veri!D7:D18)`). Python'da hesaplanıp yazılmaz.
- Filtre: girdi hücresinde açılır liste (`xk.acilir_liste`), formüller bu hücreye bağlı
  (`INDEX/MATCH`, `SUMIFS(...; Veri!B:B; $D$6)`). Dilimleyici (slicer) openpyxl ile yazılamaz.
- Sayıyı metinle birleştirmeyin (`"Ciro: "&TEXT(...)`): Türkçe Excel'de biçim bozulur. Etiket bir
  hücrede, değer yanındaki hücrede kendi `number_format`'ıyla.

## 3. Izgara

```python
xk.izgara(ws, 23)                 # A kenar, B..X 23 eşit sütun, Y kenar
xk.zemin(ws, t, "A1:Z60")
xk.sayfa_basligi(ws, t, "Satış ve Maliyet Panosu", "2026 · tüm değerler formülle gelir", genislik=23)
```

Standart yerleşim (satır numaraları şablondaki gibi):

| Satır | Blok | Sütunlar |
|---|---|---|
| 2-4 | Başlık, alt başlık, vurgu çizgisi | B-X |
| 6 | Filtre satırı | etiket B, girdi D-H, sonuç J-R |
| 8-10 | 4 KPI kartı × 5 sütun, arada 1 boş | B-F, H-L, N-R, T-X |
| 12-29 | Ana grafik + yan grafik | B-P (15) + R-X (7) |
| 31-48 | İki eşit grafik | B-L (11) + N-X (11) |

Üç kart için 7 sütun (B-H, J-P, R-X), iki kart için 11 sütun. Bloklar ortak kenarda biter; sağ kenar
her satırda aynı sütun olmalı (şablonda X).

## 4. KPI kartı

```python
xk.kpi_karti(ws, t, "B8", 5, "Toplam ciro", "=SUM(Veri!D7:D18)", "milyon",
             alt_formul="=SUM(Veri!D7:D18)/SUM(Veri!C7:C18)-1", alt_metin="plana göre")
```

- Etiket büyük harf 9 pt, değer 22 pt, alt satır değişim (`delta_ok`: yeşil ▲ / kırmızı ▼) ya da açıklama.
- Kart hücre stilidir, şekil değil: yazdırılır, kaymaz, formülle güncellenir.
- Değer için kısa biçim (`milyon`), tam tutar tabloda kalır.
- Alt satırda karşılaştırma yoksa ne ölçüldüğünü yazın ("harcanan / bütçe"); boş bırakmayın.

## 5. Gezinme ve kitap düzeni

- İlk sayfa pano ya da menü. Çok sayfalı kitapta menü sayfası + her sayfada `xk.gezinme(wb, t, "Menü")`.
- Sekme renkleri gruba göre: pano vurgu, veri mürekkep, yardımcı sayfalar gri (`xk.sekme`).
- Veri sayfasında başlık satırı dondurulur (`xk.dondur(ws, "A7")`).
- Pano baskısı tek sayfa: `xk.baski(ws, t, alan="A1:Y49", tek_sayfa=True)`.

## 6. Bitirme kontrolü

1. `xk.cakisma_denetimi(pano)` boş liste.
2. `xk.excel_ile_kaydet()` (hesaplanmış değerler dosyaya yazılsın).
3. `xk.onizle()` ve **her PNG'ye bakın**: KPI değeri `####` değil, kart kenarları hizalı, grafik
   etiketleri biçimli, pano tek sayfada, sağ kenar boyunca bloklar aynı çizgide.
4. Filtreyi başka bir değere çevirip bağlı hücrelerin değiştiğini doğrulayın (openpyxl ile değeri
   yazın, yeniden önizleyin).
5. Boş şablon da isteniyorsa veriyi silip tekrar önizleyin: `#SAYI/0!` yok, grafikler boş ama düzgün.
