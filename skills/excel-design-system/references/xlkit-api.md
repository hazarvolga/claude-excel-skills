# xlkit API

`~/.claude/skills/excel-design-system/scripts/xlkit.py`. Tüm fonksiyonlar `ws` (openpyxl Worksheet) ve
`t` (Tema) alır. Hücre adresleri A1 biçiminde.

## Tema ve biçim
| Fonksiyon | Ne yapar |
|---|---|
| `tema(ad, **degisiklik)` | Hazır tema; `vurgu=` verilirse seri listesi de güncellenir |
| `BICIM[anahtar]` | Sayı/tarih biçim kodu sözlüğü |
| `etiket_bicimi(anahtar)` | Sıfırı gizleyen üç bölümlü kod (grafik verisi hücreleri için) |
| `turkce_bugun()` | Dilden bağımsız `gg.aa.yyyy` formül parçası |

## Hücre stili
| Fonksiyon | Ne yapar |
|---|---|
| `zemin(ws, t, alan)` | Izgarayı kapatır, alanı kâğıt rengine boyar |
| `sayfa_basligi(ws, t, metin, alt, hucre, genislik)` | Başlık + alt satır + vurgu çizgisi |
| `bant(ws, t, alan, metin, koyu=True)` | Birleşik bölüm başlığı |
| `tablo_basligi(ws, t, satir, sutun1, [(metin, genişlik)])` | Koyu başlık satırı + sütun genişlikleri |
| `tablo(ws, alan, ad, stil)` | Gerçek Excel Tablosu |
| `girdi_hucresi(ws, t, alan)` | Sarı zemin + kilit açık |
| `acilir_liste(ws, alan, kaynak, baslik, mesaj)` | Veri doğrulamalı liste |
| `durum_renkleri(ws, t, alan, iyi, uyari, kotu)` | Formül tabanlı trafik ışığı yazı rengi |
| `veri_cubugu(ws, t, alan)` | Vurgu renginde veri çubuğu |
| `negatif_kirmizi(ws, t, alan)` | Negatif değer kırmızı |
| `kpi_karti(ws, t, sol_ust, sutun, etiket, formul, bicim, alt_formul, alt_bicim, alt_metin)` | 3 satırlık KPI kartı |
| `dondur(ws, hucre)` | Bölmeleri dondur |
| `sekme(ws, renk)` | Sekme rengi |
| `gezinme(wb, t, menu_sayfa)` | Her sayfaya "← Menü" bağlantısı |

## Grafik
| Fonksiyon | Ne yapar |
|---|---|
| `grafik_stili(chart, t, baslik, deger, etiket, ay_ekseni, lejant, yatay, etiket_kodu)` | Tek tasarım kuralını uygular |
| `yerlestir(ws, chart, "B12:P29")` | Grafiği tam hücre aralığına oturtur ve ekler |
| `bos_alana_sigdir(ws)` | Tek hücreye bağlı grafikleri boş alana yayar (var olan kitaplar için) |
| `cakisma_denetimi(ws)` | Grafik-grafik ve grafik-dolu hücre çakışmalarını listeler |

## Yerleşim ve baskı
| Fonksiyon | Ne yapar |
|---|---|
| `izgara(ws, sutun_sayisi=24, genislik=4.6)` | Pano için eşit sütun ızgarası |
| `hucre_araligi(sutun1, satir1, sutun_say, satir_say)` | Izgara koordinatından `"B8:M23"` |
| `baski(ws, t, ust_metin, alt_metin, yatay, alan, tek_sayfa)` | A4 baskı ayarı |

## Excel (macOS)
| Fonksiyon | Ne yapar |
|---|---|
| `excel_ile_kaydet(xlsx, temizle=True)` | Excel'de aç, hesapla, kaydet; ardından kişisel izleri siler |
| `meta_temizle(xlsx, yazar="")` | Yazar, son değiştiren, şirket ve tam dosya yolunu dosyadan siler |
| `onizle(xlsx, dizin, dpi)` | Excel → PDF → sayfa başına PNG; yolları döndürür |
