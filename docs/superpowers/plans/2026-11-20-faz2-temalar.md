# Excel skill seti · Faz 2 (temalar) Uygulama Planı

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Dört yeni tema (`mimari`, `teknik`, `baski`, `sunum`) ve her temada kullanılabilen `antet()` yardımcısı; temaların kimliğini renk dışında yazı ölçeği, rakam yazısı ve renksiz (desenli) baskı verir.

**Architecture:** `Tema` veri sınıfına geriye uyumlu üç alan eklenir (`yazi_olcegi`, `rakam_yazi`, `renksiz`); mevcut yardımcılar (`yazi`, `kpi_karti`, `sayfa_basligi`, `tablo_basligi`, `grafik_stili`, `durum_renkleri`, `negatif_kirmizi`) bu alanlara bakar. Varsayılan değerlerle mevcut üç temanın çıktısı değişmez.

**Tech Stack:** Python 3.10+, openpyxl (`PatternFillProperties`, `LineProperties.prstDash`), pytest; görsel denetim Mac Excel + LibreOffice.

**Spec:** `docs/superpowers/specs/2026-11-20-excel-skills-v2-design.md` (5.5, 5.9 Faz 2)

## Global Constraints

- 1.1.0 API'si geriye uyumlu; mevcut üç temanın token değerleri ve çıktısı değişmez.
- Vurgu rengi kâğıt renginde ≥ 4.5:1 kontrast (açık temalar beyaz/kâğıt, `gece` kendi kâğıdı).
- Başlık ve rakam yazıları Windows ve Mac Office'te bulunan fontlar: Cambria, Calibri, Calibri Light, Consolas.
- `baski` temasında grafikler renk taşımaz: seri ayrımı tarama deseni (çubuk/dilim) ve çizgi kesiği (çizgi grafiği); durum renkleri mürekkep tonlarına iner, anlam ▲ ▼ ✓ ⚠ ✗ simgeleriyle kalır.
- Depoda belirli bir şirket veya ürün markası, onun renkleri ya da adı geçmez.
- Testler: `uv run --no-project --with openpyxl --with pymupdf --with pillow --with pytest python -m pytest tests -q -W ignore::DeprecationWarning` ve `bash kur.test.sh`.
- Commit'ler `v2-tasarim` dalında, push yok; mesaj `<tür>: <açıklama>` + `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.

## Review Focus

1. **Kullanıcının kendi teması** (`xk.tema("baski", vurgu="C8102E")`): renksiz temaya renkli vurgu verilirse seriler desenli kalmalı, vurgu yalnız çizgi/kenar rengini değiştirmeli → Task 3 testi.
2. **Altıdan fazla seri / dilim** renksiz temada: desen listesi bittiğinde baştan dönmeli, hata vermemeli → Task 3 testi.
3. **Sunum ölçeğinde uzun KPI değeri** (`1.017.500 ₺`): `shrink_to_fit` korunmalı, kart taşmamalı → Task 2 görsel denetim.
4. **`antet()` dar alana** (alan sütun sayısı < alan sayısı): hata mesajıyla durmalı, sessizce üst üste yazmamalı → Task 4 testi.
5. **Gece teması ile yeni alanlar**: varsayılan değerlerde gece çıktısı birebir aynı kalmalı → Task 1 geriye uyum testi.

---

### Task 1: Tema alanları ve dört yeni tema

**Files:**
- Modify: `skills/excel-design-system/scripts/xlkit.py:44-81` (`Tema`, `TEMALAR`)
- Create: `tests/test_tema.py`

**Interfaces:**
- Produces: `Tema` yeni alanları: `yazi_olcegi: float = 1.0`, `rakam_yazi: str | None = None`, `renksiz: bool = False`. `TEMALAR` anahtarları: `kurumsal, minimal, gece, mimari, teknik, baski, sunum`.
- Değerler (murekkep, kagit, yuzey, cizgi, izgara, ara, vurgu; seriler; yazılar):

| Tema | Tokenlar | seriler | Yazı ve diğer |
|---|---|---|---|
| mimari | 262626, F5F5F2, ECECE7, D3D1CB, E4E2DD, 6E6B66, 3D5A6C | 3D5A6C, 262626, 9A9790, D3D1CB, 8AA6B5 | baslik Cambria |
| teknik | 1A1A1A, FFFFFF, F2F4F5, 9AA3AB, D9DEE2, 5F6B75, 0B6E99 | 0B6E99, 1A1A1A, 8C969E, C5CCD2, 4FA3C7 | baslik Calibri, rakam_yazi Consolas |
| baski | 000000, FFFFFF, F2F2F2, 808080, D0D0D0, 595959, 000000 | 000000, 7F7F7F, BFBFBF, 404040, D9D9D9 | baslik Calibri, renksiz True |
| sunum | 111827, FFFFFF, F3F4F6, D1D5DB, E5E7EB, 6B7280, 7C3AED | 7C3AED, 111827, 9CA3AF, D1D5DB, A78BFA | baslik Calibri Light, yazi_olcegi 1.3 |

- [ ] **Step 1: Failing testler** (`tests/test_tema.py`)
  - `test_yedi_tema_var`: `set(xk.TEMALAR) == {...7 ad}`.
  - `test_eski_temalar_degismedi`: kurumsal/minimal/gece için `dataclasses.astuple` 1.1.0 değerleriyle (dosyadaki mevcut değerler test içine sabit yazılır) ve yeni alanlar varsayılan.
  - `test_vurgu_kontrasti`: her tema için WCAG göreli parlaklık oranı `kontrast(vurgu, kagit) >= 4.5` (yardımcı test dosyasında).
  - `test_yeni_tema_alanlari`: `teknik.rakam_yazi == "Consolas"`, `baski.renksiz is True`, `sunum.yazi_olcegi == 1.3`, `mimari.baslik_yazi == "Cambria"`.
  - `test_ozel_vurgu_yeni_temada`: `xk.tema("mimari", vurgu="C8102E").seriler[0] == "C8102E"`.
- [ ] **Step 2:** FAIL doğrula.
- [ ] **Step 3:** Alanları ve dört temayı ekle.
- [ ] **Step 4:** PASS doğrula (tüm suite).
- [ ] **Step 5: Commit** `feat: dört yeni tema ve Tema alanları (yazi_olcegi, rakam_yazi, renksiz)`

### Task 2: Yazı ölçeği ve rakam yazısı

**Files:**
- Modify: `xlkit.py` → `yazi`, `sayfa_basligi`, `tablo_basligi`, `kpi_karti`, `_metin`, `_baslik`
- Test: `tests/test_tema.py`

**Interfaces:**
- Consumes: Task 1 alanları.
- Produces: `yazi(t, boyut=10, kalin=False, renk=None, baslik=False, italik=False, rakam=False) -> Font` — boyut `round(boyut * t.yazi_olcegi * 2) / 2`; `rakam=True` ve `t.rakam_yazi` varsa o font. `_olcek(t, deger: float) -> float` satır yüksekliği için `round(deger * t.yazi_olcegi, 1)`. Grafik metinleri (`_metin`, `_baslik`) aynı ölçekle (yüzde-punto birimi, tam sayıya yuvarlanır).
- `kpi_karti` değer hücresi `yazi(t, 22, kalin=True, baslik=True, rakam=True)`; `rakam` önceliklidir (teknik'te Consolas).

- [ ] **Step 1: Failing testler**
  - `test_sunum_olcegi`: sunum ile `kpi_karti` → değer hücresi `font.sz == 28.5`, değer satırı yüksekliği `46.8`; kurumsal ile `22` ve `36` (değişmedi).
  - `test_teknik_rakam_yazisi`: teknik KPI değer fontu `Consolas`, etiket fontu `Calibri`.
  - `test_grafik_metni_olcekli`: sunum ile `grafik_stili` sonrası grafik başlığı `sz == 1430`; kurumsal `1100`.
- [ ] **Step 2:** FAIL doğrula.
- [ ] **Step 3:** Uygula.
- [ ] **Step 4:** PASS + görsel denetim: `ornek_pano.py ... sunum` ve `teknik` önizlemesi (Mac Excel) açılıp bakılır; KPI değerleri kartta taşmaz (Review Focus 3).
- [ ] **Step 5: Commit** `feat: tema yazı ölçeği ve rakam yazısı`

### Task 3: Renksiz baskı (desen, çizgi kesiği, mürekkep durum)

**Files:**
- Modify: `xlkit.py` → `grafik_stili` (seri ve dilim dolguları), `durum_renkleri`, `negatif_kirmizi`, `kpi_karti` (alt biçim)
- Test: `tests/test_tema.py`

**Interfaces:**
- Produces: sabitler `DESENLER = ("solid", "ltUpDiag", "dkDnDiag", "smGrid", "pct50", "dashHorz")` ve `KESIKLER = ("solid", "dash", "sysDot", "lgDash", "dashDot")`; yardımcı `_renksiz_bicim(kod: str) -> str` (`[Color10]`, `[Red]` gibi renk belirteçlerini siler).
- Renksiz davranış (`t.renksiz`):
  - Çubuk/alan serisi i: `GraphicalProperties(pattFill=PatternFillProperties(prst=DESENLER[i % 6], fgClr=murekkep, bgClr=kagit), ln=solidFill murekkep w=6350)`; i=0 düz `murekkep`.
  - Çizgi serisi i: `ln.solidFill = murekkep`, `ln.prstDash = KESIKLER[i % 5]`.
  - Halka/pasta dilimi j: desen `DESENLER[j % 6]`, kenar `kagit`; dilim etiketi her zaman `murekkep` üstüne `kagit` olmayan okunur renk (düz siyah dilimde beyaz, desenlide siyah).
  - `durum_renkleri`: iyi → `murekkep` normal, uyarı → `murekkep` italik, kötü → `murekkep` kalın + altı çizili.
  - `negatif_kirmizi`: `murekkep` kalın.
  - `kpi_karti`: `alt_bicim` biçimi `_renksiz_bicim` ile renk belirteçsiz.
- Vurgu renkli verilse bile (`tema("baski", vurgu=...)`) dolgu desenleri `murekkep` ile çizilir; vurgu yalnız KPI kenarı ve başlık çizgisinde kalır.

- [ ] **Step 1: Failing testler**
  - `test_renksiz_cubuk_desenli`: baski + 2 serili `BarChart` → seri 0 `solidFill == "000000"`, seri 1 `pattFill.prst == "ltUpDiag"`; hiçbir seride `seriler[1..]` gri tonu solidFill yok.
  - `test_renksiz_cizgi_kesikli`: baski + 2 serili `LineChart` → seri 1 `ln.prstDash == "dash"`, iki seri de `ln.solidFill == "000000"`.
  - `test_renksiz_yedi_dilim_doner`: 7 dilimli halka → hata yok, dilim 6 deseni `DESENLER[0]`.
  - `test_renksiz_ozel_vurgu`: `tema("baski", vurgu="C8102E")` ile seri 0 dolgusu yine `000000`.
  - `test_renksiz_kpi_alt_bicim_renksiz`: baski KPI alt hücre `number_format` içinde `[Color` ve `[Red]` yok, `▲` var.
  - `test_renksiz_durum_mürekkep`: baski `durum_renkleri` kurallarının font renkleri `000000`; kurumsal'da `2E7D32` (değişmedi).
- [ ] **Step 2:** FAIL doğrula.
- [ ] **Step 3:** Uygula.
- [ ] **Step 4:** PASS + görsel denetim: `ornek_pano.py ... baski` önizlemesi Mac Excel ve LibreOffice ile; desenler ve kesikler görünür, gri tonlu yazıcıda seriler ayrışır. LibreOffice farkı varsa tuzaklar belgesine eklenir.
- [ ] **Step 5: Commit** `feat: renksiz baskı teması için desen, kesik ve mürekkep durum biçimleri`

### Task 4: `antet()` ve örnek panoda teknik antet

**Files:**
- Modify: `xlkit.py` (yeni fonksiyon, "yerleşim ve baskı" bölümü), `skills/excel-dashboard/scripts/ornek_pano.py` (teknik temada sağ üstte antet)
- Test: `tests/test_tema.py`

**Interfaces:**
- Produces: `antet(ws, t: Tema, alan: str, alanlar: list[tuple[str, str]]) -> None` — `alan` (ör. `"Q2:X4"`) eşit genişlikte `len(alanlar)` hücre grubuna bölünür; her grupta üst satır etiket (8 pt, `ara`, BÜYÜK HARF), alt satırlar birleşik değer (10 pt, `murekkep`, `rakam` yazısı); dış kenar `murekkep` orta kalınlık, iç ayraçlar `cizgi` ince. Alan sütun sayısı `len(alanlar)`'a bölünmezse ya da küçükse `ValueError("antet alanı N alan için dar: ...")`. Değerler formül olabilir (`"=Veri!B2"`).
- Örnek pano: `t.ad == "teknik"` iken başlığın sağına `antet(ws, t, <başlık satırları sağ blok>, [("Proje", "Örnek"), ("Belge no", "PNO-001"), ("Rev.", "A"), ("Tarih", "=" + turkce_bugun())])`. Diğer temalarda pano değişmez.

- [ ] **Step 1: Failing testler**
  - `test_antet_hucreleri`: 4 alan, `"B2:I4"` → `B2.value == "PROJE"`, `B3` değeri, `C2`/`D2` birleşik grup sınırları, dış sol kenar `medium`.
  - `test_antet_dar_alan_hata`: `"B2:C4"` ve 4 alan → `ValueError` mesajı `dar`.
  - `test_ornek_pano_teknik_antet`: teknik panoda bir hücre değeri `"BELGE NO"`; kurumsal panoda yok.
- [ ] **Step 2:** FAIL doğrula.
- [ ] **Step 3:** Uygula.
- [ ] **Step 4:** PASS + teknik pano önizlemesi açılıp bakılır.
- [ ] **Step 5: Commit** `feat: antet() yardımcısı, teknik temada örnek panoya antet`

### Task 5: Belgeler, yedi tema görseli, sürüm 1.2.0

**Files:**
- Modify: `skills/excel-design-system/references/tema-tokenlari.md` (7 sütunlu tablo, yeni alanlar, renksiz kuralları), `skills/excel-design-system/references/xlkit-api.md` (`antet`, `yazi(rakam=)`), `skills/excel-design-system/SKILL.md` (tema listesi ve seçim rehberi: hangi tema hangi iş için), `README.md` (tema bölümü, sürüm notu 1.2.0), `docs/temalar.png` (yedi tema), üç SKILL.md `metadata.version: "1.2.0"`
- Create: `tools/tema_gorseli.py` (README görselini üreten geliştirici betiği; skill'lerin parçası değil)

- [ ] **Step 1:** `tools/tema_gorseli.py`: yedi temada örnek panoyu üretir, `onizle()` ile ilk sayfayı alır, 4+3 ızgarada etiketli tek PNG (`docs/temalar.png`) yazar. Motor yoksa açık hata.
- [ ] **Step 2:** Görseli üret, aç, yedi panoda taşma/kırpılma olmadığını gör.
- [ ] **Step 3:** Belgeleri ve sürümleri güncelle; SKILL.md'deki tema satırı `kurumsal | minimal | gece | mimari | teknik | baski | sunum`.
- [ ] **Step 4:** Tüm testler ve `bash kur.test.sh`.
- [ ] **Step 5: Commit** `docs: yedi tema belgesi ve görseli; sürüm 1.2.0`
