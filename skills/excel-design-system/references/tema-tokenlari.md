# Tema tokenları

Her tema `xk.Tema` örneğidir. Alanlar ve görevleri:

| Token | Görev | kurumsal | minimal | gece |
|---|---|---|---|---|
| `murekkep` | Ana metin, başlık bandı zemini | 1B2A3A | 202124 | E8EAED |
| `kagit` | Sayfa ve grafik zemini | F5F7FA | FAFAFA | 1E2227 |
| `yuzey` | KPI kartı, formül hücresi | EEF1F5 | F1F3F4 | 262B31 |
| `cizgi` | İnce kenarlık, eksen | C9D1DB | DADCE0 | 3C434C |
| `izgara` | Grafik yatay ızgara | E6EAF0 | ECEDEF | 30363D |
| `ara` | İkincil metin, açıklama | 6B7785 | 70757A | 9AA4AF |
| `vurgu` | Marka, ilk seri, bağlantı | 1F6FEB | 0F9D58 | F2A33A |
| `seriler` | Grafik seri sırası | vurgu, mürekkep, gri, açık gri, ikinci ton | | |
| `iyi / uyari / kotu` | Durum, temadan bağımsız | 2E7D32 / C98A12 / B3261E | | |
| `girdi` | Kullanıcı girişi zemini | FFF8E1 | | |

## Kendi temanı yapmak

```python
t = xk.tema("kurumsal", vurgu="C8102E")                          # yalnız marka rengi
t = xk.tema("minimal", vurgu="6A1B9A", baslik_yazi="Georgia")    # serif başlık
```

Kurallar:
- Seri renkleri ardışık iki seride ayırt edilebilir olmalı: vurgu + koyu nötr ikilisi her zaman çalışır.
- Vurgu rengi beyaz zeminde en az 4.5:1 kontrast vermeli (etiket ve bağlantı metni için).
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
