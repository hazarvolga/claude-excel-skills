---
name: excel-design-system
description: "Use when an Excel workbook (.xlsx) must look professional: applying a theme or brand colors, typography, table and header styling, KPI cards, conditional formatting, Turkish number/date/currency formats, print layout, or when the user says excel tasarımı, profesyonel excel, şık tablo, kurumsal rapor, excel teması, renk paleti, biçimlendirme. Base layer for excel-charts and excel-dashboard; holds the shared xlkit.py helper library."
metadata:
  version: "1.0.0"
  scope: global
---

# Excel Tasarım Sistemi

Bu skill Excel kitabına **tutarlı bir görsel dil** verir. Grafik kuralları `excel-charts`, pano yerleşimi
`excel-dashboard` skill'indedir; ikisi de buradaki `scripts/xlkit.py` kütüphanesini kullanır.
Dosya okuma/yazma, formül yeniden hesaplama gibi temel işler `anthropic-skills:xlsx` skill'inin alanıdır.

## Kütüphane

```python
import os, sys
sys.path.insert(0, os.path.expanduser("~/.claude/skills/excel-design-system/scripts"))
import xlkit as xk
t = xk.tema("kurumsal")                    # kurumsal | minimal | gece
t = xk.tema("kurumsal", vurgu="C8102E")    # marka rengi: seriler otomatik güncellenir
```

Çalıştırma (sistem Python'una paket kurmadan):
`uv run --no-project --with openpyxl --with pymupdf --with pillow python -I betik.py`

Fonksiyon listesi ve imzalar: [references/xlkit-api.md](references/xlkit-api.md)

## Tasarım kuralları

1. **Tek vurgu rengi.** Marka rengi yalnız vurguda: başlık çizgisi, KPI kenarı, ilk seri, bağlantı.
   Geri kalan her şey mürekkep ve gri tonları. Durum renkleri (iyi / uyarı / kötü) temayla değişmez,
   çünkü anlam taşırlar.
2. **Izgarayı kapat** (`xk.zemin`). Ayrımı boşluk ve ince alt çizgi yapar, hücre ızgarası değil.
3. **Yazı hiyerarşisi:** sayfa başlığı 20 pt başlık yazısı, bölüm bandı 11 pt kalın, gövde 10 pt,
   açıklama 9 pt ikincil renk. Başlık fontu her bilgisayarda olan bir font olmalı (Georgia, Calibri
   Light); özel font gömülemez, karşı tarafta Arial'a düşer.
4. **Sayılar sağa, metin sola, başlık ortaya.** Para ve miktar her zaman binlik ayraçlı.
5. **Girdi hücresi görünür olsun:** açık sarı zemin + kilidi açık (`xk.girdi_hucresi`). Formül hücreleri
   kilitli kalır, sayfa korumasıyla kullanıcı yalnız girdiyi değiştirir.
6. **Gerçek Excel Tablosu** kullan (`xk.tablo`): filtre, şeritli satır, yapısal başvuru ve otomatik
   genişleme bedavaya gelir. Tablo adı boşluksuz ve Türkçe karaktersiz olmalı (`AylikSatis`).
7. **Değer yerine formül.** Toplam, oran, sapma hücrede formül olarak durur; Python'da hesaplanıp
   yapıştırılmaz. Kullanıcı veriyi değiştirince pano kendiliğinden güncellenmeli.
8. **Baskıya hazır:** A4 yatay, genişliğe sığdır, üst bilgi sayfa adı, alt bilgi sayfa no (`xk.baski`).
   Pano sayfaları `tek_sayfa=True`.

Renk, yazı ve boşluk tokenlarının tamamı: [references/tema-tokenlari.md](references/tema-tokenlari.md)

## Türkçe Excel tuzakları (hepsi ölçülmüş)

| Tuzak | Belirti | Çözüm |
|---|---|---|
| `TEXT(x;"#,##0")` | Türkçe Excel'de `4200000,0` | Sayıyı metne çevirme; ayrı hücre + `number_format` |
| `TEXT(TODAY();"dd.mm.yyyy")` | `00.mm.2026` | `xk.turkce_bugun()` formül parçası |
| Ay adları İngilizce | `Jan 26` | `xk.BICIM["ay"]` = `[$-41F]mmm yy` |
| openpyxl hesaplanmış değer yazmaz | Quick Look, telefon, bazı okuyucular boş görür | Son adım `xk.excel_ile_kaydet()` |
| Grafik veri etiketi biçimi | `5200000,00,, M`, `+04%` (hücrede doğru) | Etiket kaynağında yalnız `0 % + -` ve "metin"; bkz. `excel-charts` |
| Uzun tire `—` | Klişe, bazı fontlarda kutu | Orta nokta `·` veya iki nokta |

Biçim kodları (`xk.BICIM`): `sayi`, `sayi1`, `sayi2`, `yuzde`, `yuzde1`, `tl`, `tl2`, `bin`, `milyon`,
`tarih`, `ay`, `ay_kisa`, `ay_uzun`, `delta`, `delta_ok`; grafik etiketi için `etiket_bin`, `etiket_milyon`,
`etiket_sayi`, `etiket_yuzde`, `etiket_delta`. Biçim kodları her zaman **İngilizce
sözdizimiyle** yazılır (`#,##0.0`). Excel bunları dosyada böyle saklar ve Türkçe arayüzde `1.234,5` gösterir.

## Bitti sayılmadan önce

- [ ] Kitap `xk.excel_ile_kaydet()` ile Excel'de hesaplatılıp kaydedildi (macOS'ta Excel kuruluysa).
- [ ] `xk.onizle()` ile her sayfanın görüntüsü alındı ve **açılıp bakıldı**: taşan metin, `####`,
      kırpılan grafik, yanlış biçimlenmiş sayı yok.
- [ ] Hata hücresi yok (`#DEĞER!`, `#BAŞV!`, `#SAYI/0!`). Boş şablonda bölme formülleri `IFERROR` ile korundu.
- [ ] Mühendislik, maliyet veya finans sayıları kaynağıyla karşılaştırılıp doğrulandı.
- [ ] Dosya paylaşılacaksa kişisel izler silindi: Excel kaydederken kullanıcı/şirket adını ve tam klasör
      yolunu yazar. `excel_ile_kaydet()` varsayılan olarak `meta_temizle()` çağırır.

## macOS Excel otomasyonu

- Excel sandbox'ı **her yeni dosya ve klasör için** izin penceresi açar. Pencere açık kalırsa AppleScript
  `The variable wbk is not defined` hatası verir. Kullanıcıdan izni vermesini isteyin, sonra tekrar deneyin.
  Aynı dosya adını ve klasörü yeniden kullanmak yeni izin istemez.
- Aynı adlı kitap Excel'de açık kalırsa Excel yeni dosya yerine **eski açık kopyayı** kullanır ya da
  `Parameter error (-50)` verir. xlkit açılıştan önce yalnız o kitabı adıyla kapatır, PDF'in kitaptan
  yeni olduğunu denetler ve -50'de bir kez yeniden dener. Kullanıcının diğer açık kitaplarına dokunmayın.
