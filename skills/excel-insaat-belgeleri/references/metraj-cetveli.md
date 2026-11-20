# Metraj cetveli · yapı rehberi

Metraj cetveli bir pozun miktarını ölçü satırlarından hesaplar ve keşif cetveline besler. Hücre adresleri örnektir.

## Kapsam ve sınırlar

- Düzen: kurum metraj örneklerindeki **ölçü satırı** düzeni (Adet × En × Boy × Yükseklik). Minha (boşluk düşümü)
  aynı poz bloğunda **eksi satırdır**.
- **Azı / Çoğu sütunlu klasik hakediş biçimi bu rehberde yok.** Kaynaklar bu sütunların anlamında çelişiyor;
  kullanıcı bu biçimi isterse sütun anlamını ondan yazılı olarak al, tahmin etme.
- **Boşluk düşüm kuralı (hangi boşluk ne zaman düşülür) skill'de yoktur.** Kural işe, şartnameye ve sözleşmeye göre
  değişir. Kullanıcının verdiği kuralı uygula; vermediyse her minha satırını kullanıcının girdisi olarak bırak ve
  hangi kuralı uyguladığını sor. Eşik (ör. "x m²'den küçük boşluk düşülmez") uydurma.
- Donatı (demir) metrajı, Ø sütunlu ayrı bir cetveldir; bu rehberin kapsamı dışında.

## Sütunlar

| Sütun | İçerik | Tür |
|---|---|---|
| B · No | Sıra | metin |
| C · Poz No | Poz bloğunun ilk satırında | girdi |
| D · Tanım | Poz adı (blok başı) ya da ölçü açıklaması (ör. "Zemin kat aks A-B duvarı") | girdi |
| E · ± | `+` ölçü, `-` minha (açılır liste `"+,-"`, varsayılan `+`) | girdi |
| F · Adet (benzer) | Tekrar sayısı | girdi |
| G · En | m | girdi |
| H · Boy | m | girdi |
| I · Yükseklik | m | girdi |
| J · Miktar | `=IF(E9="-",-1,1)*PRODUCT(F9:I9)` | formül |
| K · Birim | Poz bloğunun ilk satırında (m, m², m³, adet) | girdi |

- `PRODUCT` boş hücreleri yok sayar: alan ölçüsünde Yükseklik boş bırakılır, uzunlukta yalnız Boy dolu olur. Boş
  ölçüyü **0 yazma** (0 yazılırsa miktar 0 olur); ölçü yoksa hücre boş kalır.
- Hepsi boş satır `PRODUCT` ile 0 verir; ekstra koruma gerekmez.
- Minha satırının açıklaması neyin düşüldüğünü söyler ("Minha: pencere 1,20 × 1,50").
- Sayı biçimi: ölçüler `xk.BICIM["sayi2"]`, miktar `"#,##0.000"` (metrajda üç ondalık yaygındır; kullanıcı başka
  isterse onu uygula).

## Poz bloğu ve toplam

```
7   [poz başı] 1  ÖRN-02.001  Hazır beton           m³          ← C, D, K dolu; F:J boş
8   +  Temel  2 × 12,00 × 0,60 × 0,80                 =IF(E8="-",-1,1)*PRODUCT(F8:I8)
9   +  ...
10  -  Minha: asansör boşluğu ...                     (eksi)
11  [poz toplamı] ÖRN-02.001 toplamı  m³              =SUBTOTAL(9,J7:INDEX(J:J,ROW()-1))
```

- Poz toplamı keşifteki ara toplamla aynı desendedir: blok başından başlar, `INDEX(J:J,ROW()-1)` ile kendi satırının
  üstünde biter; bloğa eklenen satırı kapsar.
- Poz toplamı satırı kalın, `t.yuzey` zemin; birim sütununda pozun birimi.
- Poz blokları arasında bir boş satır.

## Keşif cetveline bağlama

Keşif cetvelindeki Miktar hücresi metrajdaki poz toplamına başvurur: `='Metraj'!J11`. Böylece miktar tek yerde
girilir. Metrajı olmayan kalemde keşif miktarı girdi hücresi kalır. Bağlı miktar hücresini girdi rengiyle
boyama; `t.yuzey` zemin ver ki kullanıcı onun formül olduğunu görsün.

İsteğe bağlı **Metraj icmali** sayfası: Poz No · Tanım · Birim · Toplam miktar (`='Metraj'!J11` …) listesi;
keşif miktarları buradan da beslenebilir.

## Denetim

- Hesaplanmış kopyada her poz toplamı = o bloktaki ölçü satırlarının işaretli çarpımlarının toplamı (Python'da
  ayrıca hesapla).
- Eksi satır sayısı = `±` sütununda `-` sayısı; minha satırları gerçekten eksi.
- Birim ile dolu ölçü sayısı tutarlı mı: m³ satırında üç ölçü, m² satırında iki, m satırında bir (Adet hariç).
  Tutarsız satırı kullanıcıya göster, kendiliğinden düzeltme.
- Keşifteki bağlı miktarlar metraj poz toplamlarıyla aynı.
