---
name: excel-insaat-belgeleri
description: "Use when preparing construction documents in Excel: teklif dosyası, keşif cetveli, birim fiyat teklif cetveli, icmal, teklif mektubu, villa / apartman / büyük proje teklifi, ihale teklif şablonu (özel sektör), metraj cetveli, ölçü ve minha satırları, şantiye günlük raporu, puantaj, malzeme takip tablosu."
metadata:
  version: "2.1.0"
  scope: global
---

# İnşaat belgeleri (Excel)

Bu skill müşteriye giden **teklif dosyasını** ve **keşif / birim fiyat teklif cetvelini** Excel'de kurmayı anlatır.
Hazır şablon yoktur: kitabı her istekte kullanıcının verisiyle, aşağıdaki kurallara göre kendin kurarsın.
Görsel dil için `excel-design-system` skill'indeki `xlkit.py` kütüphanesini kullan (tema, bant, tablo başlığı,
girdi hücresi, baskı, önizleme); önce o skill'i oku. Sayfa sayfa yapı, sütunlar ve formül desenleri:
[references/teklif-kesif.md](references/teklif-kesif.md).

Diğer belgelerin yapı rehberleri (aynı değişmez kurallar geçerli):

| Belge | Rehber |
|---|---|
| Metraj cetveli (ölçü satırları, minha, keşfe bağlama) | [references/metraj-cetveli.md](references/metraj-cetveli.md) |
| Şantiye günlük raporu, aylık puantaj, malzeme takip | [references/santiye-takip.md](references/santiye-takip.md) |
| Proje portföy panosu | `excel-dashboard` skill'i, `references/portfoy-panosu.md` |

Kapsam dışı: kamu ihalesi (EKAP) resmî birim fiyat teklif cetveli biçimi, hakediş, donatı (demir) metrajı, Azı /
Çoğu sütunlu klasik metraj biçimi, boşluk düşüm eşikleri (kullanıcının kuralı uygulanır).

## 1. Önce sor (bilmiyorsan)

Tek mesajda, en fazla şu dört şey:

1. **Belge:** yalnız keşif cetveli mi, tam teklif dosyası mı (kapak + mektup + icmal + keşif + koşullar)? Ya da
   metraj cetveli, şantiye takip tablosu mu (ilgili rehberi oku)?
2. **Ölçek:** villa (tek seviye iş grubu), apartman (blok bazında miktar), büyük proje (bölüm > iş grubu > kalem)?
3. **Veri:** kalemler, miktarlar ve birim fiyatlar nereden gelecek: kullanıcının listesi, kendi Excel'i, yoksa boş şablon mu?
4. **Tema:** belirtilmediyse müşteriye giden belgede `mimari`, keşif ağırlıklı çalışmada `teknik`, renksiz yazıcı için `baski`.

## 2. Değişmez kurallar

1. **Fiyat uydurma.** Kullanıcı birim fiyat vermediyse birim fiyat hücresini boş girdi hücresi bırak; tahmin yazma.
   Gösterim için örnek veri istenirse poz numaralarını açıkça kurgusal yaz (`ÖRN-01.001`) ve kitaba
   "Örnek veri uydurmadır; gerçek proje için kendi pozlarınızı ve fiyatlarınızı girin." notunu koy. Resmî bir
   kurumun poz numarasını veya fiyatını kullanıcı vermeden yazma.
2. **Her toplam formül.** Tutar, ara toplam, icmal ve genel toplam hücrede formül olarak durur; Python'da
   hesaplanıp yapıştırılmaz.
3. **Oranlar Parametreler sayfasında.** KDV, iskonto, genel gider ve kâr oranı yalnız Parametreler'deki sarı girdi
   hücrelerinde; formüller onlara **tanımlı adla** başvurur (`KDV_ORANI`, `ISKONTO_ORANI`, `GGK_ORANI`, `GGK_AYRI`).
   Formüle `*0.2` ya da `20%` yazma. Girdi hücresindeki değer yalnız başlangıç değeridir.
4. **Ara toplam eklenen satırı kapsar.** İş grubu ara toplamı, bölüm toplamı ve genel toplam
   `=SUBTOTAL(9,H7:INDEX(H:H,ROW()-1))` desenindedir (7 = grup başlık satırı, H = Tutar). `SUBTOTAL` iç ara
   toplamları iki kez saymaz; `INDEX(…,ROW()-1)` bitişi, ara toplamın hemen üstüne eklenen satırı da kapsar.
   Sabit `SUM(H8:H10)` yazma.
5. **İcmal zinciri sabittir:** iş grubu tutarları → ara toplam → (+) genel gider ve kâr (yalnız `GGK_AYRI="Evet"`
   ise; değilse birim fiyatlara dahil, satır 0) → teklif tutarı → (−) iskonto → KDV matrahı → (+) KDV → genel toplam.
   İskonto kâr dahil teklif tutarına uygulanır.
6. **Sayıyı metne çevirme.** Mektupta tutar cümle içine `TEXT()` ile gömülmez (Türkçe Excel'de bozulur); tutar ayrı
   hücrede `xk.BICIM["tl2"]` biçimiyle durur. Metin birleştirme yalnız ad ve metin parametreleriyle (`="Sayın "&MUSTERI`).
7. **Hata hücresi yok.** Bölme içeren formüller (m² başına maliyet gibi) `IFERROR(…,0)` ile korunur; boş şablon
   hesaplandığında da hata çıkmaz.
8. **Formül dili İngilizce** (`SUBTOTAL`, `SUMPRODUCT`, `IF`); Excel Türkçe arayüzde kendisi çevirir.
   Büyük harf etiketlerde Python `str.upper()` yerine `xk.buyuk_harf()` kullan (`TEKLİF NO`, `TEKLIF NO` değil).
9. Mühendislik ve maliyet sayıları (miktar, birim fiyat, oran) kullanıcıdan ya da kaynağından gelir; değiştirme,
   yuvarlama, "düzeltme" yapma. Şüpheli bir değer görürsen kullanıcıya sor.
10. **Ad da uydurma.** İş grubu, bölüm, firma, proje adı kullanıcıdan gelir. Kullanıcı iş grubu vermediyse sor;
    soramıyorsan tüm kalemleri tek grupta ("İşler") topla ve bunu söyle. Kendi sınıflandırmanı ("Kaba yapı" gibi)
    kullanıcının adı gibi yazma.
11. **Eksik girdiyi görünür kıl.** Fiyatı ya da miktarı verilmeyen kalem varsa icmalin altına "Şu kalemlerin birim
    fiyatı girilmedi, tutarları 0 sayıldı: …" notu koy ve kullanıcıya da söyle.

## 3. Kurma sırası

1. `xlkit`'i yükle (yol deseni `excel-design-system` SKILL.md'de) ve temayı seç.
2. **Parametreler** sayfası: firma, müşteri, proje, teklif no, revizyon, tarih, para birimi, oranlar, geçerlilik,
   ödeme ve teslim; apartmanda daire tipleri tablosu. Her değer girdi hücresi + tanımlı ad.
3. **Keşif** sayfası: ölçeğe göre sütunlar, grup başlıkları (`xk.bant`), kalemler, ara toplamlar, genel toplam;
   başlık satırı dondurulur ve baskıda her sayfada tekrar eder.
4. **İcmal** sayfası: iş grubu satırları keşifteki ara toplamlara başvurur, ardından zincir.
5. Tam teklif dosyasında **Kapak**, **Teklif Mektubu**, **Koşullar**; değerler tanımlı ad formülleri.
6. Sayfa sırası: (Kapak, Teklif Mektubu,) İcmal, Keşif, (Koşullar,) Parametreler. Baskı: A4 dikey, genişliğe sığdır.
7. Kaydet, hesaplat, denetle (aşağıdaki liste), önizle.

## 4. Bitti sayılmadan önce

- [ ] Kitap `xk.excel_ile_kaydet()` ile hesaplatıldı (Excel yoksa bir kopyada `motor="libreoffice"`) ve
      hesaplanmış değerlerde **hata hücresi yok**.
- [ ] Bağımsız kontrol tuttu: Σ(miktar × birim fiyat) Python'da ayrıca hesaplanıp icmal ara toplamıyla
      karşılaştırıldı; KDV ve genel toplam zincire göre doğrulandı.
- [ ] Formüllerde sayısal oran yok (`re.search(r"[*/]\s*0?\.\d|\d+%", formül)` hiçbir icmal formülünde eşleşmez).
- [ ] Her ara toplam `SUBTOTAL(9,…:INDEX(…,ROW()-1))` deseninde.
- [ ] Örnek veri kullanıldıysa not ve kurgusal poz numaraları var; kullanıcının verdiği değerler değiştirilmedi.
- [ ] `xk.onizle()` ile sayfa görüntüleri alındı ve açılıp bakıldı: `####`, taşan metin, A4'e sığmayan keşif yok.
- [ ] Kullanıcıya hangi motorla denetlendiği ve hangi girdilerin boş bırakıldığı söylendi.
