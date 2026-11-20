# Şantiye takip tabloları · yapı rehberi

Üç tablo: **günlük rapor**, **aylık puantaj**, **malzeme takip**. Bunlar tam bir şantiye yönetim sistemi değildir;
hakediş, taşeron hesabı ve nakit akışı bu rehberin dışındadır. Yasal yükümlülük (İSG, SGK bildirimi gibi) hakkında
hüküm yazma; kullanıcı istediği alanı söyler.

## Günlük rapor

Bir gün = bir sayfa (A4 dikey, `xk.baski(ws, t, yatay=False, tek_sayfa=True)`). Çok günlük kitapta her gün ayrı
sayfa, sayfa adı `gg.aa` (ör. `20.11`), ilk sayfada boş bir **Şablon** sayfası; yeni gün için kullanıcı onu kopyalar.

| Bölüm | İçerik |
|---|---|
| Üst bilgi | `xk.antet`: Proje, Rapor no, Tarih, Hazırlayan. Tarih girdi hücresi, `tarih` biçimi |
| Hava | Durum (açılır liste: Açık, Parçalı bulutlu, Yağmurlu, Karlı, Rüzgârlı), en düşük / en yüksek °C, "Çalışmaya engel mi?" (Evet/Hayır) |
| Personel | Tablo: Meslek / ekip · Firma (kendi / taşeron adı) · Kişi sayısı · Not; altta `=SUM` toplam kişi |
| Makine ve ekipman | Tablo: Makine · Adet · Çalışma saati · Not; toplam saat `=SUM` |
| Yapılan işler | Tablo: Poz / iş grubu · Yer (blok, kat, aks) · Açıklama · Miktar · Birim |
| Gelen malzeme | Tablo: Malzeme · Miktar · Birim · İrsaliye no |
| Olaylar ve notlar | Birleşik, `wrap_text` geniş metin alanı (ziyaret, denetim, aksaklık, kaza / ramak kala notu) |
| İmza | Hazırlayan · Şantiye şefi; ad soyad ve imza satırı |

Tablolar `xk.tablo_basligi` ile başlıklı, gövde satırları girdi hücresi; her tabloda 6-10 boş satır.

## Aylık puantaj

Bir ay = bir sayfa (A4 yatay, genişliğe sığdır). Parametre hücreleri: **Yıl** ve **Ay** (girdi), tanımlı adlar
`PUANTAJ_YIL`, `PUANTAJ_AY`.

| Sütun | İçerik |
|---|---|
| B · Sıra, C · Ad soyad, D · Meslek, E · Firma | girdi |
| F … AJ · 1-31 | gün hücreleri, kod girilir |
| AK … | özet sütunları (aşağıda) |

- Gün başlığı (iki satır): üstte gün numarası, altta haftanın günü. Gün numarası
  `=IF(MONTH(DATE(PUANTAJ_YIL,PUANTAJ_AY,F$5))=PUANTAJ_AY,F$5,"")` (F5'te 1…31 sabit sayı, gizli satır ya da açık
  küçük yazı). Ayın dışında kalan günler (29-31) boş görünür.
- Haftanın günü: `=IF(F6="","",CHOOSE(WEEKDAY(DATE(PUANTAJ_YIL,PUANTAJ_AY,F$5),2),"Pt","Sa","Ça","Pe","Cu","Ct","Pz"))`.
  `TEXT(…,"ddd")` kullanma (dil Excel arayüzüne göre değişir).
- Hafta sonu sütunu: koşullu biçim `=AND(F$6<>"",WEEKDAY(DATE(PUANTAJ_YIL,PUANTAJ_AY,F$5),2)>5)` ile `t.yuzey` zemin.
  Ay dışı sütun: `=F$6=""` ile koyu gri zemin. Koşullu dolguda `xk.dolgu()` kullan (Excel'de görünmesi için).
- **Kodlar** sayfanın altında açıklanır ve gün hücrelerine açılır liste olarak verilir: `X` tam gün, `Y` yarım gün,
  `İ` izinli, `R` raporlu, `M` mazeretsiz yok, boş = çalışmadı / kayıt yok. Kullanıcı başka kod isterse onları kullan.
- Özet sütunları: Tam gün `=COUNTIF(F7:AJ7,"X")`, Yarım `=COUNTIF(F7:AJ7,"Y")`, İzinli, Raporlu, Mazeretsiz,
  **Çalışılan gün** `=AK7+AL7*0.5` (yarım günün değeri kullanıcıdan farklı gelirse onu tanımlı ad yap: `YARIM_GUN`).
- Ücret hesabı istenirse günlük ücret girdi sütunu ve `=Çalışılan gün × Günlük ücret`; vergi ve kesinti hesaplama.

## Malzeme takip

Gerçek Excel Tablosu (`xk.tablo(ws, alan, "MalzemeTakip")`): filtre, şerit, otomatik genişleme.

| Sütun | İçerik |
|---|---|
| Sipariş no · Tarih · Tedarikçi · Malzeme · Birim | girdi |
| Sipariş miktarı · Teslim alınan | girdi (teslimler toplamı; ayrıntı istenirse ayrı Teslimatlar tablosu ve `SUMIFS`) |
| Kalan | `=[@[Sipariş miktarı]]-[@[Teslim alınan]]` (yapısal başvuru) ya da hücre başvurusu |
| Termin | girdi tarih |
| Durum | `=IF([@[Kalan]]<=0,"Tamamlandı",IF([@Termin]<TODAY(),"Gecikti","Bekliyor"))` |

- Durum sütununa `xk.durum_renkleri`: Tamamlandı iyi, Bekliyor uyarı, Gecikti kötü (renksiz temada yazı biçimiyle).
- Fazla teslim (Kalan < 0) ayrıca işaretlenir: `xk.negatif_kirmizi` Kalan sütununa.
- Üstte küçük özet: toplam sipariş satırı, geciken satır sayısı `=COUNTIF(<Durum sütunu>,"Gecikti")`.

## Denetim

- Puantaj: 30 günlük bir ay için 31. sütun boş görünüyor; hafta sonları doğru günlerde; özet sayıları elle sayılan
  kodlarla aynı.
- Malzeme takip: Kalan ve Durum formülleri her satırda; tablo yeni satırda formülü taşıyor.
- Günlük rapor: toplam kişi ve makine saati `SUM` ile; boş şablon hatasız.
