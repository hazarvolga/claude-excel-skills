# Proje portföy panosu · yapı rehberi

Birden çok projeyi (ya da şantiyeyi) tek ekranda gösteren yönetim panosu. Genel pano kuralları SKILL.md'de; bu
rehber portföye özgü veri yapısını, KPI'ları ve grafikleri anlatır. Sayılar kullanıcıdan gelir; örnek istenirse
proje adlarını kurgusal yaz ve "Örnek veri uydurmadır" notunu koy.

## Veri sayfası

Gerçek Excel Tablosu `Projeler` (`xk.tablo`), bir satır = bir proje:

| Sütun | İçerik | Tür |
|---|---|---|
| Proje · Yer · Proje müdürü | metin | girdi |
| Başlangıç · Planlanan bitiş · Tahmini bitiş | tarih | girdi |
| Sözleşme bedeli · Bütçe (maliyet) · Gerçekleşen maliyet | para | girdi |
| Fiziksel ilerleme % | 0-1 | girdi (ölçülen; tahmin değil) |
| Planlanan ilerleme % | `=IFERROR(MIN(1,MAX(0,(TODAY()-[@Başlangıç])/([@[Planlanan bitiş]]-[@Başlangıç]))),0)` (doğrusal; kullanıcının iş programı varsa ondan) | formül |
| Kazanılmış değer | `=[@Bütçe]*[@[Fiziksel ilerleme %]]` | formül |
| Maliyet sapması | `=[@[Kazanılmış değer]]-[@[Gerçekleşen maliyet]]` (eksi = aşım) | formül |
| CPI | `=IFERROR([@[Kazanılmış değer]]/[@[Gerçekleşen maliyet]],0)` | formül |
| SPI | `=IF([@[Planlanan ilerleme %]]=0,1,[@[Fiziksel ilerleme %]]/[@[Planlanan ilerleme %]])` (henüz başlamamış proje gecikmiş sayılmaz) | formül |
| Gecikme (gün) | `=MAX(0,[@[Tahmini bitiş]]-[@[Planlanan bitiş]])` | formül |
| Durum | `=IF(OR([@CPI]<ESIK_KOTU,[@SPI]<ESIK_KOTU),"Kritik",IF(OR([@CPI]<ESIK_UYARI,[@SPI]<ESIK_UYARI),"Dikkat","Yolunda"))` | formül |

- Eşikler (`ESIK_UYARI`, `ESIK_KOTU`) Parametreler hücresinde, tanımlı adla; başlangıç değerini kullanıcıya sor.
  Bilmiyorsa 0,95 / 0,85 öner ve öneri olduğunu söyle; koda gömme.
- Kazanılmış değer, CPI ve SPI tanımları standart proje kontrol tanımlarıdır; planlanan ilerlemenin doğrusal
  varsayıldığını panoda küçük notla belirt.

## Pano yerleşimi (24 sütunlu ızgara)

```
Başlık + filtre (Proje müdürü / Yer, açılır liste, "Tümü" seçeneği)
KPI ×4: Toplam sözleşme · Toplam gerçekleşen maliyet · Portföy CPI · Geciken proje
Sol (2/3): Bütçe ve gerçekleşen maliyet, proje başına yatay çubuk (kümelenmiş)
Sağ (1/3): Durum dağılımı halka (Yolunda / Dikkat / Kritik)
Alt: Proje tablosu özeti (Proje, ilerleme veri çubuğu, CPI, SPI, gecikme, durum renkleri)
```

- Portföy CPI = `=IFERROR(SUM(Kazanılmış değer)/SUM(Gerçekleşen maliyet),0)`; proje CPI'larının ortalaması değil.
- Geciken proje = `=COUNTIF(Projeler[Gecikme (gün)],">0")`.
- Filtre listesi Parametreler'de: "Tümü" + kişiler + birkaç boş girdi hücresi (yeni kişi elle eklenir); açılır liste
  bu aralığa bağlanır.
- Filtre "Tümü" iken toplamlar tüm tablo, aksi halde `SUMIFS(…,Projeler[Proje müdürü],$D$6)`;
  iki durumu `IF($D$6="Tümü",SUM(…),SUMIFS(…))` ile ayır.
- Yatay çubuk grafik verisi bloğu **proje sayısı kadar yuva, en çok 8**: 8'i geçerse en büyük 8 proje `LARGE`/`INDEX`
  ile seçilir, gerisi tabloda kalır. Boş yuva bırakma (grafiğin yarısı boş görünür).
- Durum halkasının dilimleri durum renkleriyle (Yolunda `t.iyi`, Dikkat `t.uyari`, Kritik `t.kotu`; renksiz temada
  desen); sıfır dilimin etiketi `0%;;` biçimiyle gizli.
- Durum renkleri `xk.durum_renkleri` (Yolunda iyi, Dikkat uyarı, Kritik kötü); ilerleme `xk.veri_cubugu`.
- Grafik verisi bloğu ve etiket biçimleri için `excel-charts` kurallarına uy (etiket hücresinde ölçek formülle).

## Denetim

- Toplam sözleşme KPI'ı = tablo sütun toplamı (filtre "Tümü").
- Portföy CPI elle: Σ kazanılmış değer / Σ gerçekleşen maliyet.
- Durum sayıları halka grafiğiyle ve tablodaki Durum sütunuyla aynı.
- Bölmeli formüllerde hata yok (gerçekleşen maliyet 0 olan yeni proje dahil).
