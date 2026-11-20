# Teklif dosyası ve keşif cetveli · yapı rehberi

Hücre adresleri örnektir; yapıyı kendi verine göre kur, desenleri koru. Formüller İngilizce yazılır.

## Parametreler sayfası

B sütunu etiket, C sütunu girdi hücresi (`xk.girdi_hucresi`, sarı). Her değer workbook düzeyinde tanımlı ad alır.

| Tanımlı ad | Etiket | Başlangıç değeri | Biçim |
|---|---|---|---|
| `FIRMA`, `FIRMA_ADRES` | Teklif veren firma, adres | boş | metin |
| `MUSTERI` | Müşteri (işveren) | boş | metin |
| `PROJE`, `PROJE_YERI` | Proje adı, yeri | boş | metin |
| `TEKLIF_NO`, `REVIZYON`, `TEKLIF_TARIHI` | Teklif no, revizyon, tarih | boş, "0", boş | metin, metin, `tarih` |
| `PARA_BIRIMI` | Para birimi | TL | metin |
| `KDV_ORANI` | KDV oranı | 0,20 (kullanıcı değiştirir) | `yuzde` |
| `ISKONTO_ORANI` | İskonto oranı | 0 | `yuzde1` |
| `GGK_AYRI` | Genel gider ve kâr icmalde ayrı satır | Hayır (açılır liste Evet/Hayır) | metin |
| `GGK_ORANI` | Genel gider ve kâr oranı | 0 | `yuzde1` |
| `GECERLILIK_GUN` | Teklif geçerlilik süresi (gün) | 30 | `sayi` |
| `ODEME_KOSULU`, `TESLIM_SURESI` | Ödeme koşulu, teslim süresi | boş | metin |

Apartmanda altına **daire tipleri** tablosu: Tip, Adet, Net m², Brüt m² (5 girdi satırı); `DAIRE_ADET` ve
`DAIRE_BRUT` sütun aralıklarına tanımlı ad.

```python
from openpyxl.workbook.defined_name import DefinedName

def ad_ekle(wb, ad, hedef):                    # hedef: "$C$16" ya da "$E$30:$E$34"
    wb.defined_names[ad] = DefinedName(ad, attr_text=f"'Parametreler'!{hedef}")

xk.girdi_hucresi(ws, t, "C16:C16")              # tek hücre de aralık biçiminde verilir
xk.acilir_liste(ws, "C18", '"Evet,Hayır"', "Genel gider ve kâr", "Evet: icmalde ayrı satır")
```

## Keşif sayfası

### Sütunlar (ölçeğe göre)

| Ölçek | Sütunlar |
|---|---|
| villa | Sıra · Poz No · Tanım · Birim · Miktar · Birim fiyat · Tutar |
| apartman | Sıra · Poz No · Tanım · Birim · Blok A · Blok B · … · Toplam miktar · Birim fiyat · Tutar |
| büyük proje | Sıra · Poz No · Tanım · Birim · Miktar · Taşeron paketi · Rev. · Birim fiyat · Tutar |

Miktar (blok sütunları) ve birim fiyat girdi hücresidir; Tutar formüldür. Sayılar `xk.BICIM["sayi2"]`, tutarlar
`xk.BICIM["tl2"]`. Başlık `xk.tablo_basligi`, `xk.dondur(ws, "A<başlık+1>")`, baskıda başlık tekrarı
`ws.print_title_rows = "6:6"`, `xk.baski(ws, t, yatay=False)`.

### Satır düzeni ve formüller (villa, Tutar = G, Miktar = E, Birim fiyat = F)

```
6   başlık satırı
7   [bant] 01 Hafriyat ve dolgu                      ← grup başlık satırı (Tutar boş)
8   1  P-01  Makine ile kazı   m³  450   120   =E8*F8
9   2  ...                                         =E9*F9
10  [ara toplam] 01 toplamı                        =SUBTOTAL(9,G7:INDEX(G:G,ROW()-1))
11  [bant] 02 Betonarme
...
40  GENEL TOPLAM (KDV hariç)                       =SUBTOTAL(9,G7:INDEX(G:G,ROW()-1))
```

- Ara toplam aralığı **grup başlık satırından** başlar, `INDEX(G:G,ROW()-1)` ile kendi satırının bir üstünde biter.
  Kullanıcı grubun içine (ara toplamın hemen üstü dahil) satır eklerse kapsanır; yeni satırın Tutar formülünü üst
  satırdan kopyalaması yeterlidir (bunu kullanıcıya söyle).
- Genel toplam aynı desenle ilk grup başlığından başlar; `SUBTOTAL` aradaki ara toplamları saymaz.
- Ara toplam satırı kalın, üst kenarı ince çizgi; genel toplam satırı `t.yuzey` zemin, kalın.

### Apartman farkları

- `Toplam miktar = SUM(E8:G8)` (blok sütunları), `Tutar = Toplam miktar × Birim fiyat`.
- Ara toplam satırında **her blok sütununa** o grubun blok tutarı:
  `=SUMPRODUCT(E7:INDEX(E:E,ROW()-1),$I7:INDEX($I:$I,ROW()-1))` (E = Blok A, I = Birim fiyat). Başlık satırının
  boş ya da metin hücreleri SUMPRODUCT'ta 0 sayılır. Bu hücrelere `tl2` biçimi ver; miktar sütununda tutar
  durduğu için zemini `t.yuzey` yap ve ara toplam satırı olduğunu belli et.

### Büyük proje farkları

- Üç seviye: **bölüm** bandı (`xk.bant(..., koyu=True)`), **iş grubu** bandı (`koyu=False`), kalemler.
- İş grubu ara toplamı yukarıdaki desen; **bölüm toplamı** bölüm başlık satırından başlayan aynı desen
  (`SUBTOTAL` iç ara toplamları saymaz).
- Satır gruplama: kalem satırları `ws.row_dimensions[r].outline_level = 2`, iş grubu ara toplamları `1`;
  `ws.sheet_properties.outlinePr.summaryBelow = True`. Kullanıcı Excel'de 1-2-3 düğmeleriyle özet görünüme iner.

## İcmal sayfası

Tek sütunlu (villa, büyük) ya da blok sütunlu (apartman) özet. İş grubu satırları keşifteki ara toplam hücrelerine
başvurur (`='Keşif'!G10`); büyük projede bölüm satırları da.

| Satır | Formül (C = tutar sütunu) |
|---|---|
| İş grubu satırları | `='Keşif'!G10` … |
| Ara toplam | `=SUM(C7:C16)` (iş grubu satırları) |
| Genel gider ve kâr | `=IF(GGK_AYRI="Evet",C17*GGK_ORANI,0)`; "Hayır" iken etiket "Genel gider ve kâr (birim fiyatlara dahil)" ve satır `t.ara` renginde |
| Teklif tutarı | `=C17+C18` |
| İskonto | `=-C19*ISKONTO_ORANI` |
| KDV matrahı | `=C19+C20` |
| KDV | `=C21*KDV_ORANI` |
| **Genel toplam** | `=C21+C22` (kalın, `t.yuzey` zemin, büyük yazı) |
| m² başına maliyet (yalnız apartman, KDV hariç) | `=IFERROR(C21/SUMPRODUCT(DAIRE_ADET,DAIRE_BRUT),0)` |

Apartmanda iş grubu satırlarında blok sütunları keşifteki ara toplam satırının blok hücrelerine başvurur; blok
sütunlarının toplamı tutar sütunuyla aynı olmalıdır (denetimde karşılaştır).

## Kapak (tam teklif dosyası)

A4 dikey, tek sayfa (`xk.baski(ws, t, yatay=False, tek_sayfa=True)`). Üstte firma (`=FIRMA`), ortada büyük başlık
"Fiyat Teklifi", proje (`=PROJE`), müşteri (`=MUSTERI`); altta `xk.antet` ile Teklif no, Rev., Tarih
(`=TEKLIF_NO`, `=REVIZYON`, `=TEKLIF_TARIHI` + `tarih` biçimi).

## Teklif Mektubu

- Tarih sağ üstte (`=TEKLIF_TARIHI`), hitap `="Sayın "&MUSTERI&","`.
- Giriş: `="Konu: "&PROJE&" yapım işi fiyat teklifi"` ve bir paragraf (birleşik hücre, `wrap_text`, satır yüksekliği
  metne göre).
- Tutar satırı: etiket "Teklif bedeli (KDV dahil)" + **ayrı hücrede** `='İcmal'!C23` (`tl2`). `TEXT()` kullanma.
- `="Bu teklif "&GECERLILIK_GUN&" gün geçerlidir."` (sayı birleştirme gün gibi tam sayıda sorun çıkarmaz).
- İmza: `=FIRMA`, altında ad soyad ve unvan için boş girdi hücreleri.

## Koşullar

Numaralı maddeler: ödeme (`=ODEME_KOSULU`), teslim (`=TESLIM_SURESI`), geçerlilik, teminat ("Teminatın oranı ve
şekli sözleşmede belirlenir."), fiyat farkı, teklife dahil olmayan işler (girdi hücreleri), imza alanı. Oran veya
yasal hüküm uydurma; kullanıcı vermediyse girdi hücresi bırak.

## Denetim parçacığı

```python
import re
from openpyxl import load_workbook

ORAN = re.compile(r"[*/]\s*0?\.\d|\d+%")
wb = load_workbook(xlsx)                                   # formüller
for c in (c for row in wb["İcmal"].iter_rows() for c in row):
    assert not (isinstance(c.value, str) and c.value.startswith("=") and ORAN.search(c.value)), c.coordinate

hesap = load_workbook(xlsx, data_only=True)                # excel_ile_kaydet sonrası değerler
hatalar = [f"{ws.title}!{c.coordinate}" for ws in hesap for row in ws.iter_rows() for c in row
           if isinstance(c.value, str) and c.value.startswith("#")]
assert not hatalar, hatalar
```
