# openpyxl grafik tuzakları

Hepsi Excel for Mac (Türkçe arayüz) ile PDF'e kaydedilip görüntüden doğrulandı (Kasım 2026).

| # | Tuzak | Belirti | Çözüm |
|---|---|---|---|
| 1 | `BarChart(type="bar")` kurucuda yok sayılır | Yatay istenen grafik dikey çizilir | `chart.type = "bar"` sonradan, ya da `grafik_stili(..., yatay=True)` |
| 2 | Yatay çubukta kategori sırası ters | En büyük altta | `x_axis.scaling.orientation = "maxMin"` + `y_axis.crosses = "max"` (değer ekseni altta kalır) |
| 3 | Veri etiketi biçim kodu yok sayılır | Etiket kaynak hücre biçimini gösterir | Etiket biçimini hücreye verin |
| 4 | Türkçe Excel etiket biçimini yerel sözdizimiyle okur | `5200000,00,, M`, `+04%`, `5190001900r1l` | Etiket kaynağında yalnız `0 % + -` ve "metin"; ölçek formülle (`=ROUND(x/1000,0)`) |
| 5 | Değer ekseni biçiminde ondalık yok | `1 M, 1 M, 1 M` tekrar eden etiketler | `#,##0.0,, "M"` (ondalıklı) |
| 6 | Negatif çubuk | Beyaz/içi boş çubuk | `series.invertIfNegative = False` |
| 7 | Negatif değerli eksende kategori etiketleri | Etiketler sıfır çizgisinde, çubukların arasında | `x_axis.tickLblPos = "low"` |
| 8 | `x_axis.delete` / `y_axis.delete` varsayılanı | Bazı sürümlerde eksen hiç çizilmez | `delete = False` açıkça |
| 9 | Tek hücreye bağlı grafik (`add_chart(ch, "B5")`) | Ekrana göre taşar, hücrelerin üstüne biner | `TwoCellAnchor` ile aralık (`xk.yerlestir`) |
| 10 | openpyxl'in yazdığı grafikli dosya | exceljs gibi okuyucular `reading 'anchors'` hatası verir | Dosyayı Excel'de açıp kaydedin (`xk.excel_ile_kaydet`) |
| 11 | Excel'de aynı adlı kitap açık | Önizleme eski kopyayı gösterir, ya da `Parameter error (-50)` | xlkit açmadan önce yalnız o kitabı adıyla kapatır, PDF tarihini denetler, -50'de bir kez yeniden dener |
| 12 | Excel sandbox | Yeni dosya/klasörde `variable wbk is not defined` | Ekrandaki izin penceresini onaylayın; aynı dosya adını yeniden kullanın |
| 13 | Lejant varsayılan sağda | Çizim alanı daralır | `legend.position = "b"` |
| 14 | `chart.style` varsayılan | Office mavisi gölge ve renkler | `chart.style = None` + seri renkleri elle |
| 15 | Çizim alanı varsayılanı beyaz | Koyu temada açık seriler ve etiketler kaybolur; açık temada halkanın arkasında beyaz kare | `plot_area.graphicalProperties` temaya göre (`grafik_stili` yapar) |
| 16 | Halka/pasta etiketi tek renk | Açık dilimde beyaz etiket okunmaz | Dilim başına `DataLabel(idx=j, txPr=...)`, renk dilimin koyuluğuna göre |

## Grafik verisi bloğu deseni

Tablo tam tutarı gösterir, grafik güvenli biçimli kısa etiket ister:

```python
xk.tablo_basligi(ws, t, 6, 15, [("Bütçe", 10), ("Harcanan", 10)])
for i in range(5):
    for j, kaynak in enumerate(("K", "L")):
        c = ws.cell(7 + i, 15 + j, f"=ROUND({kaynak}{7 + i}/1000,0)")
        c.number_format = xk.BICIM["etiket_bin"]      # 0" B"
# Grafik O6:P11'den beslenir; başlıkta birimi yazın: "... (bin ₺)".
```

Gizli sütundaki veri grafikte görünmez (Excel varsayılanı). Bloğu gizlemek gerekiyorsa
`chart.plotVisOnly = False` verin.
