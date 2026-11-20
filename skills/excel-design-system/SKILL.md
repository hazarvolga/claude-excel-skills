---
name: excel-design-system
description: "Use when an Excel workbook (.xlsx) must look professional: applying a theme or brand colors, typography, table and header styling, KPI cards, conditional formatting, Turkish number/date/currency formats, print layout, or when the user says excel tasarımı, profesyonel excel, şık tablo, kurumsal rapor, excel teması, renk paleti, biçimlendirme."
metadata:
  version: "1.0.1"
  scope: global
---

# Excel Tasarım Sistemi

Bu skill Excel kitabına **tutarlı bir görsel dil** verir. Grafik kuralları `excel-charts`, pano yerleşimi
`excel-dashboard` skill'indedir; ikisi de buradaki `scripts/xlkit.py` kütüphanesini kullanır.
Dosya okuma/yazma gibi temel işler, araçta varsa genel xlsx skill'inin alanıdır. Bu skill `excel-charts` ve
`excel-dashboard` için temel katmandır: ortak `scripts/xlkit.py` kütüphanesi buradadır.

## Kütüphane

```python
import os, sys
SKILL_DIR = "<bu SKILL.md dosyasının bulunduğu klasör>"   # araç klasörünü sabit yazmayın
sys.path.insert(0, os.path.join(SKILL_DIR, "scripts"))
import xlkit as xk
t = xk.tema("kurumsal")                    # kurumsal | minimal | gece
t = xk.tema("kurumsal", vurgu="C8102E")    # marka rengi: seriler otomatik güncellenir
```

İlk kullanımda ortamı denetleyin: `python3 SKILL_DIR/scripts/ortam.py`. Python sürümünü, paketleri,
bulunan hesaplama motorlarını ve eksikleri yazar; openpyxl yoksa çıkış kodu 1 verir.

Çalıştırma: `uv` varsa sistem Python'una paket kurmadan
`uv run --no-project --with openpyxl --with pymupdf --with pillow python -I betik.py`; yoksa bir kez
`python3 -m pip install --user openpyxl pymupdf pillow`, sonra `python3 betik.py`.

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
| Koşullu biçimde dolgu | Excel'de görünmez, LibreOffice'te görünür (ör. Gantt çubukları, satır şeritleri) | `fill=xk.dolgu(renk)` (v1.0.1'den beri `fgColor` ve `bgColor` birlikte); elle yazıyorsanız `PatternFill("solid", bgColor=renk)` |

Biçim kodları (`xk.BICIM`): `sayi`, `sayi1`, `sayi2`, `yuzde`, `yuzde1`, `tl`, `tl2`, `bin`, `milyon`,
`tarih`, `ay`, `ay_kisa`, `ay_uzun`, `delta`, `delta_ok`; grafik etiketi için `etiket_bin`, `etiket_milyon`,
`etiket_sayi`, `etiket_yuzde`, `etiket_delta`. Biçim kodları her zaman **İngilizce
sözdizimiyle** yazılır (`#,##0.0`). Excel bunları dosyada böyle saklar ve Türkçe arayüzde `1.234,5` gösterir.

## Bitti sayılmadan önce

- [ ] Kitap `xk.excel_ile_kaydet()` ile hesaplatılıp kaydedildi (Excel kuruluysa; dönüş değeri motor adıdır).
- [ ] `xk.onizle()` ile her sayfanın görüntüsü alındı ve **açılıp bakıldı**: taşan metin, `####`,
      kırpılan grafik, yanlış biçimlenmiş sayı yok.
- [ ] Hata hücresi yok (`#DEĞER!`, `#BAŞV!`, `#SAYI/0!`). Boş şablonda bölme formülleri `IFERROR` ile korundu.
- [ ] Mühendislik, maliyet veya finans sayıları kaynağıyla karşılaştırılıp doğrulandı.
- [ ] Sonuç raporunda hangi motorla denetlendiği yazıldı (`excel-mac`, `excel-win`, `libreoffice`) ya da
      motor bulunamadığı için görsel denetimin yapılamadığı açıkça söylendi.
- [ ] LibreOffice ile denetlendiyse "Excel'de grafik görünümü küçük farklar gösterebilir" uyarısı verildi.
- [ ] Dosya paylaşılacaksa kişisel izler silindi: Excel kaydederken kullanıcı/şirket adını ve tam klasör
      yolunu yazar. `excel_ile_kaydet()` varsayılan olarak `meta_temizle()` çağırır.

## Hesaplama ve önizleme motorları

`excel_ile_kaydet()` ve `onizle()` sırayla şu motorları dener; hangisinin kullanıldığını döndürür:

1. `excel-mac`: macOS + Microsoft Excel (AppleScript).
2. `excel-win`: Windows + Microsoft Excel (COM, `pywin32` gerekir). Excel görünmez açılır, yalnız açtığı kitabı kapatır.
3. `libreoffice`: her sistemde, `soffice --headless`, yalıtılmış profil (kullanıcının açık LibreOffice'ine
   takılmaz, formüller her zaman yeniden hesaplanır, Türkçe yerel ayar). **Kaydetmede otomatik kullanılmaz**:
   LibreOffice ile kaydetmek biçimi bozabilir. Gerekirse `excel_ile_kaydet(x, motor="libreoffice")` ile açıkça seçin.
   Önizlemede kullanılır; Excel'den farkları: `excel-charts` skill'inde
   `references/openpyxl-grafik-tuzaklari.md`, "LibreOffice önizleme farkları".

Bir motor hata verirse nedeni uyarı olarak yazılır ve sıradakine geçilir. Hiçbiri yoksa dosya yine üretilir,
`onizle()` boş liste, `excel_ile_kaydet()` `None` döner ve "görsel denetim yapılamadı" uyarısı verilir.
`motor="..."` ile tek bir motor zorlanabilir; o zaman hata yükseltilir.

### macOS Excel ayrıntıları

- Excel sandbox'ı **her yeni dosya ve klasör için** izin penceresi açar. Pencere açık kalırsa AppleScript
  `The variable wbk is not defined` hatası verir. Kullanıcıdan izni vermesini isteyin, sonra tekrar deneyin.
  Aynı dosya adını ve klasörü yeniden kullanmak yeni izin istemez.
- Aynı adlı kitap Excel'de açık kalırsa Excel yeni dosya yerine **eski açık kopyayı** kullanır ya da
  `Parameter error (-50)` verir. xlkit açılıştan önce yalnız o kitabı adıyla kapatır, PDF'in kitaptan
  yeni olduğunu denetler ve -50'de bir kez yeniden dener. Kullanıcının diğer açık kitaplarına dokunmayın.
