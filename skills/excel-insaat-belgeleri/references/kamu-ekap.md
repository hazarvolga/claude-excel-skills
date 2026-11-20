# Kamu ihalesi · birim fiyat teklif cetveli (KİK012) ve hazırlık kitabı

4734 sayılı Kanun kapsamındaki **birim fiyat** (ya da karmanın birim fiyatlı kısmı) yapım ihaleleri için. Excel iki iş
görür: (1) **Hazırlık**: istekli birim fiyatlarını hesaplar ve aritmetik denetimden geçirir, sonra EKAP'a girer;
(2) **Basılı cetvel**: KİK012 e-formunun düzeniyle, fiziki ortamda teklif alınan işler ve teklif dosyası için.

Mevzuat değişir. Aşağıdaki kurallar **20.11.2026** itibarıyla birincil kaynaklardan doğrulandı (sonda liste). Kullanıcı
daha yeni bir değişiklik bildirirse onu esas al; kendi hafızandaki eski form kodlarını (KİK015.x/Y) kullanma.

## Değişmez kurallar

1. **A grubu idarenindir, değiştirme.** Sıra No, İş Kalemi No, Adı ve Kısa Açıklaması, Birimi, Miktarı idarece
   hazırlanır (KİK012 dipnot 2). Kullanıcıdan idarenin cetvelini / iş kalemleri listesini iste ve **aynen** al: sıra,
   metin, birim ve miktar değişmez; kalem eklenmez, silinmez, birleştirilmez. Liste yoksa cetveli kurma, iste.
2. **B grubu isteklinindir.** Teklif Edilen Birim Fiyat ve Tutarı istekli doldurur (dipnot 3). Fiyatı kullanıcı
   (ya da onun maliyet analizi) verir; uydurma. Boş birim fiyat varsa cetvel eksiktir: açıkça uyar.
3. **Tutar = Miktar × Birim fiyat** (dipnot 5). Satır tutarını **yuvarlama**: Tebliğ 39.2 birim fiyat ve tutarın iki
   ondalıktan fazla yazılmasına izin verir; yalnız **toplam teklif tutarı** en yakın iki ondalığa yuvarlanır (yarım
   kuruş ve üstü bir kuruşa tamamlanır).
4. **Görünen = saklanan.** Aritmetik hata çarpım ve toplamlarda yazılan rakamlarla aranır (Tebliğ 16.5.1, 39.3;
   Yönetmelik md. 57(7)) ve teklifi değerlendirme dışı bırakır. Bu yüzden:
   - Birim fiyat hücresinin **değeri** seçilen hane sayısına yuvarlanmış olur: `=ROUND(Hazırlık!H8,BF_HANE)`; biçimi
     aynı hane sayısını gösterir. Arkada fazla ondalık tutup az göstermek yasaktır.
   - Tutar hücresi `=E8*F8`; biçimi tutarın **tüm** ondalıklarını gösterir. Miktar en çok `m` haneli, birim fiyat
     `BF_HANE` haneli ise tutar en çok `m + BF_HANE` hanelidir; biçimi buna göre seç (ör. `#,##0.00###` en az 2, en
     çok 5 hane gösterir). Gereken hane 5'i aşarsa `BF_HANE`'yi düşürmeyi kullanıcıya öner.
   - Toplam: `=ROUND(SUM(G8:INDEX(G:G,ROW()-1)),2)`, biçim iki hane.
   - Tebliğ 16.5.3'teki tolerans (tablolama programının yuvarlama farkı toplamın binde birini aşmıyorsa komisyon
     düzeltir) bir **istisnadır**, güvenlik payı değildir; cetveli farksız kur.
5. **KDV hariç, TL.** Toplam satırı "TOPLAM TUTAR (KDV Hariç)"; para birimi Türk Lirası (e-ihale tip idari şartname
   11.2). KDV, iskonto, kâr satırı cetvelde yer almaz; bunlar hazırlık sayfasında kalır.
6. **Kısımlar.** Kısmi teklife açık ihalede her kısım ayrı cetvel ya da ayrı blok ve ayrı toplam; kısım toplamı da iki
   haneye yuvarlanır (Tebliğ 16.5.2). Kısımları idare düzenler (dipnot 1).
7. **Teklif mektubu Excel'de kurulmaz.** E-ihalede mektup KİK009 e-formuyla EKAP'ta hazırlanıp e-imzayla gönderilir;
   fiziki teklifte KİK217. Mektuptaki bedel cetvelin toplamıdır ve **rakam ve yazı ile** uyumlu yazılır (Yönetmelik
   md. 54(2)(c)). İstenirse raporunda toplamın Türkçe yazısını ver ("yazı ile: …") ve kullanıcıya kontrol ettir; bunu
   hücreye formül olarak koyma.

## Kitap yapısı

Sayfa sırası: **Cetvel**, **Hazırlık**, **Kontrol**, **Parametreler**.

### Parametreler

Girdi hücreleri ve tanımlı adlar: `IKN` (İhale Kayıt Numarası), `IS_ADI`, `IDARE`, `ISTEKLI`, `BF_HANE` (birim fiyat
ondalık hanesi, başlangıç 2, kullanıcı seçer), `GGK_ORANI` (hazırlıkta maliyet üstüne eklenen genel gider ve kâr,
yalnız hazırlık içindir).

### Hazırlık (istekli içi, sunulmaz)

| Sütun | İçerik |
|---|---|
| Sıra No · İş Kalemi No · Adı ve Kısa Açıklaması · Birimi · Miktarı | idarenin listesi, aynen (A grubu kopyası) |
| Maliyet birim fiyatı | girdi: kullanıcının analizi / fiyatı |
| Hedef birim fiyat | `=Maliyet*(1+GGK_ORANI)` (kullanıcı başka yöntem verirse onu uygula) |
| Teklif birim fiyatı | `=ROUND(Hedef,BF_HANE)` (cetvele giden değer) |
| Tutar | `=Miktar*Teklif birim fiyatı` |

Altta toplam `=ROUND(SUM(…),2)`. Bu sayfa iç çalışmadır; başlığında "Sunulmaz · istekli iç hesabı" yazsın.
EKAP'ta yalnız birim fiyat girilir (KİK e-teklif kılavuzu, 15.02.2024): kullanıcı her kalemin **Teklif birim fiyatı**
değerini aynı hane sayısıyla girer.

### Cetvel (KİK012 düzeni, basılı)

```
                       BİRİM FİYAT TEKLİF CETVELİ (1)
İhale Kayıt Numarası (İKN): <=IKN>
| A (2)                                                                   | B (3)                               |
| Sıra No | İş Kalemi No (4) | Adı ve Kısa Açıklaması | Birimi | Miktarı | Teklif Edilen Birim Fiyat | Tutarı (5) |
| 1       | ...              | ...                    | m³     | 120,000 | =ROUND(Hazırlık!H8,BF_HANE) | =E8*F8   |
...
| TOPLAM TUTAR (KDV Hariç)                    (ilk altı sütun birleşik)       | =ROUND(SUM(...),2)          |
```

- Başlık, İKN satırı, A/B grup başlıkları (A ilk beş, B son iki sütunun üstünde, birleşik), yedi sütun başlığı, toplam
  satırı ve aşağıdaki **beş dipnot aynen** (KİK012, 18.07.2024 kararı):
  1. "a) Kısmi teklif verilmesine izin verilen ihalelerde kısımlar ihale dokümanına ve ikincil ihale mevzuatına uygun
     olarak düzenlenecektir. b) Konsorsiyumların ihaleye teklif verebileceklerinin öngörülmesi halinde, bu cetvel işin
     uzmanlık gerektiren bölümleri esas alınarak idarece ayrı ayrı düzenlenecektir."
  2. "Bu sütun idarece hazırlanacaktır."
  3. "Bu sütun isteklilerce doldurulacaktır."
  4. "Bu sütuna sadece yapım işleri ihalelerinde yer verilecektir."
  5. "Hesaplama yapılırken her bir iş kaleminin miktar sütunundaki rakam ile teklif edilen birim fiyat çarpılarak o iş
     kalemine ait teklif tutarı bulunacaktır."
- Alt bilgi: "KİK012 Birim Fiyat Teklif Cetveli (Mal/Yapım/Danışmanlık)".
- Güncel formda imza alanı yoktur (imza mektupta, e-imzayla). Fiziki teklifte idare doküman ayrıca imza isterse
  kullanıcının söylediğini ekle.
- Görünüm resmî ve renksiz: `baski` teması ya da düz siyah-beyaz; **girdi sarısı, vurgu rengi, KPI yok**. Cetvelde
  hiçbir hücre girdi değildir; tüm B grubu Hazırlık'tan formülle gelir. Miktar idarenin yazdığı hane sayısıyla
  gösterilir. A4 dikey (çok kalemde yatay), genişliğe sığdır, başlık satırları her sayfada tekrar (`print_title_rows`),
  sayfa numarası alt bilgide.

### Kontrol (sunulmaz)

Formüllerle her satır için: `=IF(Cetvel!G8=Cetvel!E8*Cetvel!F8,"✓","HATA")`, birim fiyat hane kontrolü
`=IF(Cetvel!F8=ROUND(Cetvel!F8,BF_HANE),"✓","HATA")`, boş birim fiyat `=IF(Cetvel!F8="","EKSİK","✓")`; altta
`=COUNTIF(…,"HATA")` ve `=COUNTIF(…,"EKSİK")` özetleri, toplamın `ROUND(SUM,2)` olduğu ve Hazırlık toplamıyla eşitliği.

## Bitti sayılmadan önce

- [ ] A grubu idarenin listesiyle birebir: kalem sayısı, sıra, metin, birim, miktar (Python'da karşılaştır).
- [ ] Hesaplanmış kopyada her satır için `tutar == miktar × birim fiyat` **ondalık hassasiyetle** (Python `Decimal`
      ile, hücre değerlerinden) ve `toplam == ROUND(Σ tutar, 2)`, yarım kuruş yukarı.
- [ ] Her birim fiyat değeri `BF_HANE` hanede; tutar biçimi tutarın tüm hanelerini gösteriyor (görüntüde kesilen hane yok).
- [ ] Kontrol sayfasında `HATA` ve `EKSİK` sayısı 0 (ya da eksikler kullanıcıya bildirildi).
- [ ] Cetvelde KDV, iskonto, kâr satırı yok; toplam satırı "TOPLAM TUTAR (KDV Hariç)"; beş dipnot ve alt bilgi aynen.
- [ ] Önizlemede cetvel okunur; sarı ya da renkli zemin yok.
- [ ] Kullanıcıya söylendi: EKAP'ta girilecek birim fiyatlar Hazırlık'taki **Teklif birim fiyatı** sütunudur; mektupta
      rakam ve yazıyla aynı toplam yazılır; mevzuat tarihi 20.11.2026 itibarıyla doğrulandı.

## Doğrulanmamış (kullanıcıya böyle söyle)

- 01.08.2025 sonrası yeni EKAP'ta **yapım ihalesi** e-teklifinde Excel şablonu indirme / yükleme olup olmadığı ve birim
  fiyata izin verilen azami hane sayısı doğrulanamadı. (Excel şablonu yalnız doğrudan temin kılavuzunda anlatılıyor;
  orada şablona satır veya sütun eklemek yüklemeyi engelliyor.) EKAP'ta böyle bir şablon görülürse ona satır, sütun
  ekleme; değerleri Hazırlık'tan aynı haneyle taşı.
- Aşırı düşük teklif sınır değeri, teklif türünün seçimi (Yönetmelik md. 5) ve yaklaşık maliyet bu rehberin dışında.

## Kaynaklar (20.11.2026 itibarıyla)

- KİK e-formları KİK009 (Birim Fiyat Teklif Mektubu), KİK012 (Birim Fiyat Teklif Cetveli); KİK kararı 18.07.2024,
  2024/DK.D-133; e-Formlar: https://www.kik.gov.tr/Mevzuat.aspx · KİK012:
  https://www.kik.gov.tr/DokumanDownload.aspx?docTur=eform&DokumanID=418
- Eski KİK015.x/Y standart formları kaldırıldı: RG 18.05.2024 / 32550, yürürlük 01.08.2025 (Yönetmelik md. 54(1):
  teklif mektupları e-formlar esas alınarak hazırlanır).
- Kamu İhale Genel Tebliği 16.5.1-16.5.3 ve 39.1-39.3 (işlenmiş metin, son değişiklik RG 16.10.2026 / 33402):
  https://www.kik.gov.tr/DokumanDownload.aspx?DokumanID=1571 · 16.5.3 asıl: RG 20.11.2018 / 30601
- Yapım İşleri İhaleleri Uygulama Yönetmeliği md. 54, 57(7) (işlenmiş metin):
  https://www.kik.gov.tr/DokumanDownload.aspx?DokumanID=1565
- EKAP Yönetmeliği ve e-ihale tip idari şartname (md. 11.2 TL): https://www.kik.gov.tr/DokumanDownload.aspx?DokumanID=1495
- KİK E-Teklif Verme Kullanım Kılavuzu, Sürüm 2.0, 15.02.2024 (yeni EKAP öncesi olabilir).
- Danıştay 13. Daire E:2018/670, K:2019/1593 (iki haneye yuvarlama zorunluluğunun kaldırılması, aritmetik hata kontrolü).
