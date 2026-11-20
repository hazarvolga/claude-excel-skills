# Tema tokenları

Her tema `xk.Tema` örneğidir. Alanlar ve görevleri:

| Token | Görev | kurumsal | minimal | gece | mimari | teknik | baski | sunum |
|---|---|---|---|---|---|---|---|---|
| `murekkep` | Ana metin, başlık bandı zemini | 1B2A3A | 202124 | E8EAED | 262626 | 1A1A1A | 000000 | 111827 |
| `kagit` | Sayfa ve grafik zemini | F5F7FA | FAFAFA | 1E2227 | F5F5F2 | FFFFFF | FFFFFF | FFFFFF |
| `yuzey` | KPI kartı, formül hücresi | EEF1F5 | F1F3F4 | 262B31 | ECECE7 | F2F4F5 | F2F2F2 | F3F4F6 |
| `cizgi` | İnce kenarlık, eksen | C9D1DB | DADCE0 | 3C434C | D3D1CB | 9AA3AB | 808080 | D1D5DB |
| `izgara` | Grafik yatay ızgara | E6EAF0 | ECEDEF | 30363D | E4E2DD | D9DEE2 | D0D0D0 | E5E7EB |
| `ara` | İkincil metin, açıklama | 6B7785 | 70757A | 9AA4AF | 6E6B66 | 5F6B75 | 595959 | 6B7280 |
| `vurgu` | Marka, ilk seri, bağlantı | 1F6FEB | 0F9D58 | F2A33A | 3D5A6C | 0B6E99 | 000000 | 7C3AED |
| `baslik_yazi` | Sayfa başlığı, KPI değeri | Calibri Light | Calibri | Calibri | Cambria | Calibri | Calibri | Calibri Light |
| `rakam_yazi` | KPI ve antet değeri (yoksa başlık yazısı) | | | | | Consolas | | |
| `yazi_olcegi` | Yazı ve satır yüksekliği çarpanı | 1.0 | 1.0 | 1.0 | 1.0 | 1.0 | 1.0 | 1.3 |
| `renksiz` | Siyah-beyaz baskı kuralları | | | | | | evet | |
| `seriler` | Grafik seri sırası | vurgu, mürekkep, gri, açık gri, ikinci ton | | | | | | |
| `iyi / uyari / kotu` | Durum, temadan bağımsız | 2E7D32 / C98A12 / B3261E | | | | | (mürekkep) | |
| `girdi` | Kullanıcı girişi zemini | FFF8E1 | | | | | | |

## Hangi tema hangi iş için

| Tema | Kullanım |
|---|---|
| `kurumsal` | Genel yönetim raporu, varsayılan |
| `minimal` | Sade iç rapor, yoğun tablo |
| `gece` | Yalnız ekran panosu (baskıda mürekkep harcar) |
| `mimari` | Tasarım ofisi, sunum dosyası, müşteriye giden teklif |
| `teknik` | Keşif, metraj, mühendislik çizelgesi; `antet()` ile pafta görünümü |
| `baski` | Renksiz yazıcı, fotokopi, faks; ayrım desen ve kesikle |
| `sunum` | Toplantıda ekrana yansıtma; büyük yazı ve KPI |

## Renksiz baskı (`renksiz=True`)

- Çubuk, alan ve dilimler `xk.DESENLER` sırasıyla düz mürekkep ya da mürekkep-kâğıt tarama deseniyle çizilir.
- Çizgi serileri mürekkep renginde, `xk.KESIKLER` sırasıyla düz, kesik, noktalı.
- Desenli dilimin yüzde etiketi kâğıt renginde zemin alır.
- Durum renkleri mürekkebe iner: iyi normal, uyarı italik, kötü kalın ve altı çizili; KPI değişiminde ▲ ▼ kalır,
  renk belirteci (`[Color10]`, `[Red]`) silinir. Negatif değer kalın mürekkep.
- `tema("baski", vurgu="...")` verilse bile grafik dolguları mürekkeple çizilir; vurgu yalnız KPI kenarı ve başlık
  çizgisinde görünür.

## Kendi temanı yapmak

```python
t = xk.tema("kurumsal", vurgu="C8102E")                          # yalnız marka rengi
t = xk.tema("minimal", vurgu="6A1B9A", baslik_yazi="Georgia")    # serif başlık
```

Kurallar:
- Seri renkleri ardışık iki seride ayırt edilebilir olmalı: vurgu + koyu nötr ikilisi her zaman çalışır.
- Vurgu rengi kâğıt renginde en az 4.5:1 kontrast vermeli (etiket ve bağlantı metni için). Yeni temalar bunu
  sağlar; `kurumsal` (4,3) ve `minimal` (3,4) 1.0'dan beri biraz altındadır, geriye uyum için değiştirilmedi.
  Bağlantı metni önemliyse marka rengini `vurgu=` ile koyulaştırın.
- Koyu tema (`gece`) yalnız ekran panoları için; baskıda mürekkep harcar. Yazdırılacaksa açık tema.
- Altıdan fazla kategoriyi renkle ayırmaya çalışma: ilk beşi renkli, gerisi "Diğer" olarak toplanır.

## Boşluk ve ölçü

- Pano ızgarası: A sütunu 2 genişlik kenar, sonra 23-24 eşit sütun (4.6 genişlik). 24 sütun 2, 3, 4, 6
  parçaya tam bölünür.
- Satır yükseklikleri: başlık 32, KPI kartı 20 / 36 / 18, tablo başlığı 30, gövde varsayılan (15).
- Bloklar arasında bir boş sütun ve bir boş satır. Blok kenarları ortak hizada biter.

## Yazı

| Rol | Boyut | Ağırlık | Renk |
|---|---|---|---|
| Sayfa başlığı | 20 | normal | murekkep, başlık fontu |
| KPI değeri | 22 | kalın | murekkep, başlık fontu |
| Bölüm bandı | 11 | kalın | kagit üstüne murekkep |
| Tablo başlığı | 10 | kalın | kagit (koyu bant üstünde) |
| Gövde | 10 | normal | murekkep |
| KPI etiketi | 9 | kalın, BÜYÜK HARF | ara |
| Açıklama | 9 | normal | ara |
| Grafik başlığı | 11 | kalın | murekkep |
| Eksen, lejant | 9 | normal | ara |
| Veri etiketi | 8 | normal | murekkep (yığılmış koyu çubukta beyaz) |
