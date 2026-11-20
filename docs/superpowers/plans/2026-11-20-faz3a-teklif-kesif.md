# Excel skill seti · Faz 3a (teklif dosyası ve keşif cetveli) Uygulama Planı

> **GEÇERSİZ (2026-11-20):** Kullanıcı üretici betik yerine kural ağırlıklı skill'i seçti. Bkz. tasarım belgesi
> 5.6 "Karar". Bu plandaki icmal zinciri, SUBTOTAL/INDEX deseni ve ölçek kararları kural belgesine taşındı.

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Dördüncü skill `excel-insaat-belgeleri`: özel sektör için keşif / birim fiyat teklif cetveli ve teklif dosyası üreticileri; villa, apartman ve büyük proje ölçeklerinde, her temada.

**Architecture:** Üretici betikler `excel-design-system/scripts/xlkit.py`'yi kardeş klasörden bulur (Faz 1 deseni). Ortak parçalar küçük modüllere ayrılır: parametre sayfası ve tanımlı adlar, örnek veri, keşif + icmal kurucu. İki CLI (`kesif_cetveli.py`, `teklif_dosyasi.py`) bu parçaları birleştirir. Her toplam formüldür; oranlar Parametreler sayfasındaki girdi hücrelerinde, formüller onlara tanımlı adla başvurur.

**Tech Stack:** Python 3.10+, openpyxl (DefinedName, satır gruplama), pytest; formül doğrulaması LibreOffice ile yeniden hesaplatıp değer karşılaştırarak; görsel denetim Mac Excel.

**Spec:** `docs/superpowers/specs/2026-11-20-excel-skills-v2-design.md` (5.6, 5.9 Faz 3). Bu plan Faz 3'ün ilk yarısıdır; metraj cetveli, portföy panosu ve şantiye takip tabloları Faz 3b planında.

## Global Constraints

- Örnek veriler uydurmadır: poz numaraları `ÖRN-<iki hane>.<üç hane>` biçiminde, kitapta "Örnek veri uydurmadır; gerçek proje için kendi pozlarınızı ve fiyatlarınızı girin." notu. Resmî kurum adı, poz numarası, birim fiyatı yok.
- Oranlar (KDV, iskonto, genel gider ve kâr) koda ve formüle sayı olarak gömülmez; Parametreler'deki girdi hücreleri ve tanımlı adlar (`KDV_ORANI`, `ISKONTO_ORANI`, `GGK_ORANI`, `GGK_AYRI`) kullanılır. Girdi hücresindeki varsayılan değer yalnız örnektir (KDV %20, iskonto %0, GGK %0 ve "Hayır").
- Formül metinleri İngilizce fonksiyon adlarıyla yazılır (openpyxl), sayıyı metne çeviren `TEXT()` kullanılmaz (Türkçe Excel tuzağı); tutar metinde değil ayrı hücrede biçimle gösterilir.
- Hata hücresi yok: boş şablonda bölme içeren formüller `IFERROR` ile korunur.
- Depoda belirli bir şirket veya ürün markası geçmez.
- Dosya başına ≤ 400 satır.
- Testler: `uv run --no-project --with openpyxl --with pymupdf --with pillow --with pytest python -m pytest tests -q -W ignore::DeprecationWarning` ve `bash kur.test.sh`. LibreOffice gerektiren testler `skipif` ile korunur.
- Commit'ler `v2-tasarim` dalında, push yok.

## Kararlar (plan yazarının; spec'in açık bıraktığı yerler)

- **İcmal zinciri:** iş grubu tutarları → **Ara toplam** → (+) **Genel gider ve kâr** (`GGK_AYRI="Evet"` ise `ara × GGK_ORANI`, değilse 0 ve satır "birim fiyatlara dahil" notuyla soluk) → **Teklif tutarı** → (−) **İskonto** (`teklif × ISKONTO_ORANI`) → **KDV matrahı** → **KDV** (`matrah × KDV_ORANI`) → **Genel toplam**. İskonto kâr dahil teklif tutarına uygulanır.
- **SUBTOTAL:** iş grubu ara toplamları ve genel toplam `SUBTOTAL(9, …)`; böylece aynı sütundaki iç ara toplamlar iki kez sayılmaz.
- **Apartman:** her blok ayrı Miktar sütunu + `Toplam miktar = SUM(blok sütunları)`; Tutar = toplam miktar × birim fiyat; icmalde blok başına tutar `SUMPRODUCT(blok miktarları, birim fiyatlar)` iş grubu aralığında. Parametreler'de daire tipleri tablosu (tip, adet, net m², brüt m²); icmal "m² başına maliyet (KDV hariç)" = `IFERROR(KDV matrahı / SUMPRODUCT(adet, brüt m²), 0)`.
- **Büyük proje:** Bölüm > İş grubu > Kalem; kalem satırları outline seviye 2, iş grubu ara toplamları seviye 1; ek sütunlar "Taşeron paketi" ve "Rev."; bölüm toplamı `SUBTOTAL(9, bölüm aralığı)`; icmal bölüm ve iş grubu satırlarıyla.
- **Boş şablon** (`--ornek` yok): aynı iş grubu başlıkları, her grupta 5 boş kalem satırı; formüller yerinde, toplamlar 0.

## Review Focus

1. **Kullanıcı satır ekler** (iş grubu içine yeni kalem, ara toplamın hemen üstü dahil): ara toplam ve genel toplam yeni satırı içermeli → `INDEX(…,ROW()-1)` bitişli aralık; Task 2 testi deseni doğrular. Kullanıcı yeni satırın Tutar formülünü üst satırdan kopyalamalı; SKILL.md bunu söyler.
2. **KDV oranı değişir** (Parametreler'de %10): tüm zincir güncellenmeli → Task 2 LibreOffice testi.
3. **Boş şablon** hesaplanınca hata hücresi yok (m² bölümü dahil) → Task 2 testi.
4. **Renksiz tema** (`baski`) ve **sunum** ölçeğinde geniş keşif sayfası A4 genişliğine sığmalı → Task 4 görsel denetim.
5. **Türkçe karakterli müşteri / proje adı** kapak ve mektupta bozulmadan görünmeli → Task 3 testi.

---

### Task 1: Skill iskeleti, parametre sayfası, örnek veri

**Files:**
- Create: `skills/excel-insaat-belgeleri/SKILL.md` (ilk sürüm: tetikleyici açıklama + kısa kullanım; Task 5'te tamamlanır)
- Create: `skills/excel-insaat-belgeleri/scripts/belge_ortak.py`, `skills/excel-insaat-belgeleri/scripts/ornek_veri.py`
- Create: `tests/test_belgeler.py`

**Interfaces:**
- `belge_ortak.py`:
  - modül başında Faz 1 bulma deseni (kardeş `excel-design-system/scripts`), sonra `import xlkit as xk` ve `xk`'yi dışa verir.
  - `OLCEKLER = ("villa", "apartman", "buyuk")`
  - `PARAMETRELER: tuple[tuple[str, str, object, str], ...]` — (tanımlı ad, etiket, varsayılan, biçim anahtarı). Adlar: `FIRMA`, `FIRMA_ADRES`, `MUSTERI`, `PROJE`, `PROJE_YERI`, `TEKLIF_NO`, `REVIZYON`, `TEKLIF_TARIHI`, `PARA_BIRIMI`, `KDV_ORANI` (0.20, yuzde), `ISKONTO_ORANI` (0, yuzde1), `GGK_AYRI` ("Hayır"), `GGK_ORANI` (0, yuzde1), `GECERLILIK_GUN` (30, sayi), `ODEME_KOSULU`, `TESLIM_SURESI`.
  - `parametre_sayfasi(wb, t, olcek: str, ornek: bool) -> Worksheet` — "Parametreler" sayfası; her parametre etiket + girdi hücresi (`xk.girdi_hucresi`), workbook düzeyinde `DefinedName`; `GGK_AYRI` için Evet/Hayır açılır liste; `olcek == "apartman"` ise daire tipleri tablosu ve `DAIRE_ADET`, `DAIRE_BRUT` tanımlı adları (sütun aralıkları). `ornek=False` iken metin parametreleri boş, oranlar varsayılan.
  - `argumanlar(argv, aciklama) -> argparse.Namespace` — `cikti`, `--tema` (xk.TEMALAR, varsayılan "kurumsal"), `--olcek` (OLCEKLER, varsayılan "villa"), `--ornek`, `--onizle DIZIN`.
  - `ORNEK_NOTU = "Örnek veri uydurmadır; gerçek proje için kendi pozlarınızı ve fiyatlarınızı girin."`
- `ornek_veri.py`:
  - `@dataclass(frozen=True) class Kalem: poz: str; tanim: str; birim: str; miktar: tuple[float, ...]; birim_fiyat: float` (miktar: villa/büyük tek eleman, apartman blok başına)
  - `@dataclass(frozen=True) class IsGrubu: kod: str; ad: str; kalemler: tuple[Kalem, ...]; bolum: str = ""`
  - `IS_GRUPLARI: tuple[tuple[str, str], ...]` — genel adlar (kod, ad): Hafriyat ve dolgu, Betonarme, Duvar, Yalıtım, Çatı, Sıva ve şap, Doğrama, Kaplama, Boya, Sıhhi tesisat, Elektrik tesisatı, Çevre düzenlemesi.
  - `BLOKLAR = ("A", "B", "C")`
  - `veri(olcek: str, ornek: bool) -> tuple[IsGrubu, ...]` — villa: ilk 10 grup, grup başına 3 kalem; apartman: 10 grup × 3 kalem, 3 blok miktarı; büyük: 3 bölüm ("Altyapı", "Kaba inşaat", "İnce işler") × 4 grup × 5 kalem. `ornek=False`: aynı gruplar, her grupta 5 boş kalem (poz/tanım "", miktar 0'lar, fiyat 0). Sayılar sabit tohumlu (`random.Random(2026)`) üretilir, yuvarlak, gerçekçi büyüklükte; poz `ÖRN-{grup:02d}.{sıra:03d}`.
  - `daire_tipleri(ornek: bool) -> tuple[tuple[str, int, float, float], ...]` — (tip, adet, net m², brüt m²); örnek: 2+1, 3+1, 4+1.

- [ ] **Step 1: Failing testler** (`tests/test_belgeler.py`; `sys.path`'e `skills/excel-insaat-belgeleri/scripts` eklenir)
  - `test_olcek_boyutlari`: villa 10 grup × 3, apartman kalem miktarı uzunluğu 3, büyük 12 grup ve 3 farklı `bolum`.
  - `test_ornek_poz_kurgusal`: tüm poz `re.fullmatch(r"ÖRN-\d{2}\.\d{3}", poz)`.
  - `test_ornek_veri_tekrarlanabilir`: `veri("villa", True) == veri("villa", True)`.
  - `test_bos_sablon_sifir`: `veri("villa", False)` her grupta 5 kalem, tüm miktar ve fiyat 0.
  - `test_parametre_tanimli_adlar`: boş Workbook + `parametre_sayfasi(wb, tema, "apartman", True)` → `wb.defined_names` içinde `KDV_ORANI`, `GGK_AYRI`, `DAIRE_BRUT`; `KDV_ORANI` hücresinin değeri 0.2 ve girdi zemini `FFF8E1`.
- [ ] **Step 2:** FAIL doğrula. **Step 3:** Uygula. **Step 4:** PASS. **Step 5: Commit** `feat: excel-insaat-belgeleri iskeleti, parametre sayfası, örnek veri`

### Task 2: Keşif ve icmal kurucu, `kesif_cetveli.py`

**Files:**
- Create: `skills/excel-insaat-belgeleri/scripts/kesif.py`, `skills/excel-insaat-belgeleri/scripts/kesif_cetveli.py`
- Test: `tests/test_belgeler.py`

**Interfaces:**
- `kesif.py`:
  - `@dataclass(frozen=True) class KesifSonucu: sayfa: str; tutar_sutunu: str; genel_toplam: str; grup_toplamlari: tuple[tuple[str, str], ...]; blok_sutunlari: tuple[str, ...]` (hücre adresleri `"'Keşif'!K42"` biçiminde; icmal bunları kullanır)
  - `kesif_sayfasi(wb, t, olcek: str, gruplar: tuple[IsGrubu, ...]) -> KesifSonucu` — sütunlar: Sıra, Poz No, Tanım, Birim, [villa/büyük: Miktar | apartman: Blok A, Blok B, Blok C, Toplam miktar], [büyük: Taşeron paketi, Rev.], Birim fiyat, Tutar. Grup başlığı `xk.bant`, kalemler, ara toplam satırı (`SUBTOTAL(9,…)`), en altta genel toplam. Miktar ve birim fiyat girdi hücresi; tutar formül. Büyük: bölüm başlığı ve bölüm toplamı, outline. `xk.dondur` başlık altına, `xk.baski(..., yatay=False)` genişliğe sığdır, tekrar eden başlık satırı (`print_title_rows`).
  - `icmal_sayfasi(wb, t, olcek: str, kesif: KesifSonucu, gruplar) -> str` — Kararlar bölümündeki zincir; apartman blok sütunları ve m² satırı; döndürür: genel toplam hücre adresi. Her oran tanımlı adla (`KDV_ORANI` vb.).
- `kesif_cetveli.py`: `main(argv: list[str]) -> int` — Parametreler, Keşif, İcmal sayfalarını kurar; `--ornek` ise not satırı; kaydeder; `--onizle` ise `xk.onizle`. Çağrı: `python3 scripts/kesif_cetveli.py CIKTI.xlsx --tema teknik --olcek apartman --ornek`.

- [ ] **Step 1: Failing testler**
  - `test_kesif_her_olcek_ve_tema` (parametrize: 3 ölçek × `kurumsal`, `baski`, `sunum`; ayrıca villa × 7 tema): çıkış kodu 0, dosya açılır, sayfa adları `["Parametreler", "Keşif", "İcmal"]` (sıra: İcmal, Keşif, Parametreler — kullanıcı önce icmali görür).
  - `test_icmal_oran_gomulu_degil`: İcmal formüllerinde `re.search(r"[*/]\s*0?\.\d|\d+%", formül)` eşleşmez; `KDV_ORANI` geçer.
  - `test_ornek_notu_ve_poz`: `--ornek` dosyasında not metni var; boş şablonda yok.
  - `@skipif(LibreOffice yok)` `test_icmal_degerleri_dogru` (3 ölçek): `xk.excel_ile_kaydet(x, motor="libreoffice")`, `data_only` ile oku; beklenen ara toplam = Σ(toplam miktar × birim fiyat) (Python'da `veri()` ile hesaplanır), KDV = ara × 0.20, genel toplam = ara × 1.20 (GGK Hayır, iskonto 0); apartman blok tutarları toplamı = ara toplam; m² satırı = matrah / Σ(adet × brüt).
  - `@skipif` `test_kdv_ve_ggk_degisince_zincir`: Parametreler'de KDV 0.10, GGK_AYRI "Evet", GGK 0.15, iskonto 0.05 yazılıp kaydedilir, LibreOffice'le hesaplatılır: genel toplam = ara × 1.15 × 0.95 × 1.10 (±0.01).
  - `test_ara_toplam_eklenen_satiri_kapsar`: her ara toplam formülü `=SUBTOTAL(9,<sütun><grup başlık satırı>:INDEX(<sütun>:<sütun>,ROW()-1))` biçiminde (regex). Gerekçe: Excel, aralığın son satırının hemen altına (ara toplamın üstüne) eklenen satırı sabit aralığa katmaz; `INDEX(…,ROW()-1)` bitişi ara toplam satırıyla birlikte kayar, grup içine nereye eklenirse eklensin kapsar. Genel toplam ve bölüm toplamı aynı desenle.
  - `@skipif` `test_bos_sablon_hatasiz` (3 ölçek): hesaplatılmış dosyada hiçbir hücre `#` ile başlayan hata değeri değil.
- [ ] **Step 2:** FAIL doğrula. **Step 3:** Uygula. **Step 4:** PASS + bir villa ve bir apartman dosyası Mac Excel önizlemesiyle açılıp bakılır.
- [ ] **Step 5: Commit** `feat: keşif ve icmal kurucu, kesif_cetveli.py (villa, apartman, büyük)`

### Task 3: `teklif_dosyasi.py` (kapak, mektup, koşullar)

**Files:**
- Create: `skills/excel-insaat-belgeleri/scripts/teklif.py` (kapak, mektup, koşullar kurucuları), `skills/excel-insaat-belgeleri/scripts/teklif_dosyasi.py` (CLI)
- Test: `tests/test_belgeler.py`

**Interfaces:**
- `teklif.py`:
  - `kapak_sayfasi(wb, t) -> None` — proje, müşteri, teklif no, revizyon, tarih, firma; değerler `=PROJE` gibi tanımlı ad formülleri; `xk.antet` alt blokta (Teklif no, Rev., Tarih, Sayfa); A4 dikey tek sayfa.
  - `mektup_sayfasi(wb, t, genel_toplam: str) -> None` — "Sayın" + müşteri (`="Sayın "&MUSTERI&","`), giriş paragrafı (proje adı birleştirme), tutar satırı: etiket + ayrı hücre `=<genel_toplam>` `xk.BICIM["tl2"]`, geçerlilik cümlesi (`="Bu teklif "&GECERLILIK_GUN&" gün geçerlidir."`), imza alanı (firma, ad soyad / unvan boş girdi). Uzun metin hücreleri birleşik ve `wrap_text`.
  - `kosullar_sayfasi(wb, t) -> None` — başlıklı madde listesi: ödeme koşulu (`=ODEME_KOSULU`), teslim süresi, geçerlilik, teminat ("Teminatın oranı ve şekli sözleşmede belirlenir."), fiyat farkı, hariç tutulan işler (girdi hücreleri), imza alanı.
- `teklif_dosyasi.py`: `main(argv) -> int` — sayfa sırası Kapak, Teklif Mektubu, İcmal, Keşif, Koşullar, Parametreler; keşif/icmal Task 2'den.

- [ ] **Step 1: Failing testler**
  - `test_teklif_sayfa_sirasi` (3 ölçek): sayfa adları sırayla.
  - `test_mektup_metin_toplam_ayri_hucre`: mektupta `TEXT(` geçmez; genel toplama başvuran hücrenin biçimi `xk.BICIM["tl2"]`.
  - `test_turkce_ad_bozulmaz`: `--ornek` ile müşteri "Şükrü Öztürk İnşaat Ltd. Şti." yazılır (Parametreler `MUSTERI` hücresi test içinde değiştirilip LibreOffice ile hesaplatılır), kapaktaki müşteri hücresinin hesaplanmış değeri birebir aynı.
  - `@skipif` `test_teklif_hatasiz` (3 ölçek, örnek ve boş): hata hücresi yok.
- [ ] **Step 2:** FAIL. **Step 3:** Uygula. **Step 4:** PASS + Mac Excel önizlemesinde kapak ve mektup açılıp bakılır (taşma, kırpılma).
- [ ] **Step 5: Commit** `feat: teklif_dosyasi.py (kapak, mektup, koşullar)`

### Task 4: Görsel denetim ve örnek dosyalar

**Files:**
- Create: `examples/teklif-villa-mimari.xlsx`, `examples/teklif-apartman-teknik.xlsx`, `examples/teklif-buyuk-kurumsal.xlsx` ve her biri için icmal sayfası PNG'si (`examples/teklif-*-icmal.png`)
- Modify: üreticilerde görsel denetimde çıkan sorunlar (sütun genişliği, sayfa sığdırma)

- [ ] **Step 1:** Yedi tema × villa ve `baski`/`sunum` × büyük için `teklif_dosyasi.py --onizle` (Mac Excel); her sayfanın PNG'si açılıp bakılır: `####`, taşan metin, kırpılan sütun, A4'e sığmayan keşif (Review Focus 4).
- [ ] **Step 2:** Bulunan her sorun için önce bir test (ör. sütun genişliği ≥ değer), sonra düzeltme.
- [ ] **Step 3:** Üç örnek dosya `--ornek` ile üretilir, `xk.excel_ile_kaydet()` (Mac Excel) ile hesaplatılır ve kişisel izler temizlenir; icmal PNG'leri `xk.onizle` ile.
- [ ] **Step 4:** Tüm testler.
- [ ] **Step 5: Commit** `feat: örnek teklif dosyaları ve görsel düzeltmeler`

### Task 5: SKILL.md, belgeler, kurulum, sürüm 2.0.0

**Files:**
- Modify: `skills/excel-insaat-belgeleri/SKILL.md` (tam: ne zaman, iki üretici, ölçek seçimi, parametreler, icmal zinciri, "Bitti sayılmadan önce" listesi: hata hücresi yok, örnek notu, rakamlar kaynakla doğrulandı)
- Create: `skills/excel-insaat-belgeleri/references/belge-yapilari.md` (sayfa sayfa yapı, sütunlar, ölçek farkları, tanımlı adlar, icmal zinciri formülleri)
- Modify: `README.md` (dört skill tablosu, örnek dosyalar tablosu, sürüm notu 2.0.0), dört SKILL.md `metadata.version: "2.0.0"`, `tests/test_yol.py` açıklama testi dördüncü skill'i de kapsar
- [ ] **Step 1: Failing test:** `test_aciklamalar_use_when` listesine `excel-insaat-belgeleri`; `kur.test.sh`'ın `SKILLLER` dizisine eklenir → FAIL (açıklama henüz tam değilse) ya da PASS ise testin dördüncü skill'i taradığı çıktıdan doğrulanır.
- [ ] **Step 2:** Belgeleri yaz. Açıklama: "Use when preparing a construction offer or bill of quantities in Excel: teklif dosyası, keşif cetveli, birim fiyat teklif cetveli, icmal, teklif mektubu, villa / apartman / büyük proje teklifi, ihale teklif şablonu (özel sektör)."
- [ ] **Step 3:** Tüm testler, `bash kur.test.sh`, marka taraması (`grep -ri` belirli şirket/ürün adları yok).
- [ ] **Step 4: Commit** `docs: excel-insaat-belgeleri belgeleri, örnekler; sürüm 2.0.0`
