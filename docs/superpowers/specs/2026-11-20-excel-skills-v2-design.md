# Excel skill seti 2.0 · tasarım belgesi

Tarih: 2026-11-20 · Durum: onaylandı (2026-11-20) · Kapsam: bu depo (`skills/` altındaki skill'ler ve kurulum)

## 1. Amaç ve neden

Skill seti (1.0.1) bugün işe yarıyor ama yalnız Claude Code'da ve yalnız macOS + Microsoft Excel ile tam
çalışıyor. Yeni sürümün üç amacı var:

1. **Her AI aracında ve her bilgisayarda çalışsın.** Claude Code / Claude Desktop, Codex, Antigravity (Gemini),
   Qoder, OpenCode, MLX-Serve; macOS, Windows ve (Excel'siz) Linux.
2. **Daha fazla görsel seçenek.** Mevcut üç temanın (kurumsal, minimal, gece) yanına dört yeni tema.
3. **İnşaat sektörüne yönelik şık belge düzenleri.** Müşteriye giden teklif ve keşif cetveli, metraj cetveli,
   yönetim/portföy panosu, şantiye takip tabloları; villa, apartman ve büyük proje ölçeklerinde.

Bugünkü sorunlar (incelemede bulundu):

- SKILL.md ve betikler kütüphaneyi `~/.claude/skills/excel-design-system/scripts` sabit yolundan çağırıyor.
  `~/.agents/skills` altındaki elle kopyalanmış sürümde yol `~/.Codex/...` olarak değiştirilmiş ve orada
  skill olmadığı için kütüphane bulunamıyor.
- Codex, Gemini, MLX-Serve, OpenCode, Qoder klasörlerinde skill hiç yok; Claude ve `.agents` kopyaları
  birbirinden bağımsız ve kayıyor.
- `excel_ile_kaydet()` ve `onizle()` yalnız macOS AppleScript ile çalışıyor; Windows'ta bu adımlar yok.

## 2. Hedefler, hedef dışı, başarısızlık sayılanlar

**Hedefler**
- Tek kaynak: bu depo. Kurulum betiği yedi araç klasörüne bağlantı (varsayılan) ya da kopya kurar.
- Windows'ta Excel varsa hesaplatma ve önizleme Excel ile; Excel yoksa LibreOffice ile; o da yoksa dosya yine
  üretilir ve görsel denetimin yapılamadığı açıkça söylenir.
- 1.0.1 ile yazılmış kullanıcı betikleri değişmeden çalışmaya devam eder (API geriye uyumlu).
- Dört yeni tema, dördüncü skill `excel-insaat-belgeleri` ve üç ölçek ayarı.

**Hedef dışı (bu sürümde yok)**
- PDF tasarımı (sıfırdan tasarlanan PDF belgeleri). Ayrı proje olarak ele alınacak. Önizleme için Excel → PDF →
  PNG yolu bugünkü gibi kalır; bu bir denetim adımıdır, ürün değildir.
- Kamu ihalesi (EKAP) birim fiyat teklif cetveli resmî biçimi. Faz 4; ayrıca tasarlanacak.
- Tam bir şantiye yönetim kitabı (hakediş, taşeron, nakit akışı zinciri). Skill yalnız tablo düzenleri verir.
- Pivot, dilimleyici, makro (openpyxl yazamaz; bugünkü sınır sürer).

**Başarısızlık sayılanlar (teknik olarak çalışsa bile)**
- Depoya belirli bir şirketin veya ürünün markası, rengi, örneği ya da göndermesi girmesi. Temalar ve
  örnekler genel ve markasızdır.
- Gerçek ya da resmî kurum birim fiyatı, poz numarası veya kataloğundan satır girmesi. Örnek veriler uydurmadır,
  "Örnek veri uydurmadır" notu taşır; poz numaraları `ÖRN-01.001` gibi açıkça kurgusal biçimdedir.
- Oranların (KDV, kâr, genel gider, teminat) koda gömülmesi. Hepsi parametre sayfasındaki girdi hücresidir.
- Bir aracın skill'i bulup kütüphaneyi bulamaması ya da sessizce yanlış yoldan eski kopyayı yüklemesi.
- Windows'ta veya LibreOffice'te bozuk görünen ama Mac Excel'de denetlendiği için "bitti" sayılan çıktı.

## 3. Kısıtlar

- Depo herkese açık, MIT lisanslı. Belgeler Türkçe öncelikli, README'de kısa İngilizce özet sürer.
- Python 3.10+. Bağımlılıklar: `openpyxl` (zorunlu), `pymupdf` + `pillow` (önizleme için). Windows Excel yolu
  için `pywin32` (yalnız Windows'ta, isteğe bağlı).
- Skill açıklamaları (`description`) yalnız tetikleyici koşulu söyler ("Use when ..."); skill'in içeriğini veya
  bağımlılığını anlatmaz. Bu bilgiler gövdeye taşınır.
- Betikler göreli yolla çağrılır: `python3 scripts/x.py` (skill klasöründen) ya da tam yol. `~/.claude` gibi
  bir araç klasörü hiçbir dosyada sabit yazılmaz.
- Frontmatter yalnız `name`, `description` ve `metadata` içerir (araçlar arası ortak alt küme).
- Biçim kodları İngilizce sözdizimiyle yazılır; Türkçe Excel tuzakları tablosu geçerliliğini korur.

## 4. Değerlendirilen yaklaşımlar

| Yaklaşım | Özeti | Bedeli |
|---|---|---|
| **A · Tek kaynak, ortak kütüphane, set halinde kurulum (seçildi)** | `xlkit.py` yalnız `excel-design-system/scripts/` içinde; diğer skill'ler kardeş klasörden bulur, bulamazsa açık hata verir. `kur.sh` / `kur.ps1` dört skill'i birlikte kurar | Skill tek başına kopyalanırsa çalışmaz; kurulum hep set halinde |
| B · Her skill kendi kütüphane kopyası | Her skill bağımsız çalışır | Dört kopya zamanla ayrışır (bugünkü `.agents` kaymasının aynısı); eşitleme testi gerekir |
| C · Kütüphane ayrı Python paketi | En temiz sürümleme | Paket yayın yükü; ilk kullanımda internet; çevrimdışı araçlarda sorun |

Kullanıcı A'yı seçti: denenmiş kurulum desenini izliyor, tek kaynak tutuyor, yayın yükü getirmiyor.

## 5. Kararlaştırılan tasarım

### 5.1 Depo yapısı (hedef)

```
skills/
  excel-design-system/   SKILL.md, references/, scripts/xlkit.py, scripts/ortam.py
  excel-charts/          SKILL.md, references/
  excel-dashboard/       SKILL.md, references/, scripts/ornek_pano.py, scripts/portfoy_pano.py
  excel-insaat-belgeleri/ SKILL.md, references/, scripts/ (belge üreticileri)
kur.sh, kur.ps1, kur.test.sh
tests/                   pytest testleri
examples/, docs/
```

### 5.2 Kütüphaneyi bulma

- Her betik kütüphaneyi kendi konumundan bulur: `Path(__file__).resolve().parent` → kardeş skill
  `../../excel-design-system/scripts`. Bağlantıyla kurulumda `resolve()` depo içindeki gerçek yolu verir;
  kopyayla kurulumda kardeş klasör aynı araç klasöründedir. İkisi de yoksa mesaj:
  "excel-design-system bulunamadı; skill setini kur.sh / kur.ps1 ile birlikte kurun."
- SKILL.md'de kullanıcı betiği için yol şöyle anlatılır: "Bu SKILL.md'nin bulunduğu klasörü `SKILL_DIR`
  say; `sys.path.insert(0, f"{SKILL_DIR}/scripts")`." Araç adı geçmez.

### 5.3 Platform katmanı (hesaplat, kaydet, önizle)

`excel_ile_kaydet()` ve `onizle()` imzaları değişmez; içeride sırayla dener:

1. **macOS + Microsoft Excel:** bugünkü AppleScript yolu (bilinen tuzaklarla birlikte).
2. **Windows + Microsoft Excel:** COM otomasyonu (`win32com`, `Workbooks.Open`, `CalculateFull`, `Save`,
   `ExportAsFixedFormat`). Excel görünmez açılır, yalnız açtığı kitabı kapatır, kullanıcının diğer kitaplarına
   dokunmaz.
3. **LibreOffice (her sistemde):** `soffice --headless` ile hesaplatma ve PDF. Sonuç raporunda "LibreOffice ile
   denetlendi; Excel'de grafik görünümü küçük farklar gösterebilir" uyarısı yer alır.
4. **Hiçbiri yok:** dosya üretilir, işlev açık bir uyarıyla `None` / boş liste döndürür; skill "görsel denetim
   yapılamadı" der. Sessiz geçiş yoktur.

İsteğe bağlı `motor="excel" | "libreoffice"` parametresi seçimi zorlar. Hangi motorun kullanıldığı dönüş
değerinde ya da günlükte görünür.

`scripts/ortam.py` (yeni): Python sürümü, kurulu paketler, bulunan motorlar ve eksikler için kurulum önerisini
yazar. Skill ilk kullanımda bunu çalıştırır.

Bağımlılık çalıştırma: `uv` varsa `uv run --no-project --with ... python -I`; yoksa `python3 -m pip install
--user openpyxl pymupdf pillow` önerisi. SKILL.md ikisini de anlatır.

### 5.4 Kurulum

- `kur.sh` (macOS/Linux) ve `kur.ps1` (Windows), örnek alınan desenle aynı davranır: varsayılan bağlantı,
  `--kopya`, `--kontrol`, `--kaldir`; yalnız var olan araç klasörlerine kurar; aynı adda yabancı içerik varsa
  dokunmaz ve bildirir; kopya kurulumu işaret dosyasıyla tanınır.
- Hedef klasörler: `~/.claude/skills`, `~/.agents/skills`, `~/.codex/skills`, `~/.gemini/skills`,
  `~/.mlx-serve/skills`, `~/.config/opencode/skills`, `~/.qoder/skills`. Windows'ta karşılıkları
  `%USERPROFILE%` altında.
- Windows'ta sembolik bağlantı yönetici hakkı ya da Geliştirici Modu ister; olmazsa betik dizin bağlantısı
  (junction) dener, o da olmazsa kopyaya düşer ve bunu söyler.
- Bugün elle kopyalanmış eski klasörler (işaretsiz; `~/.claude/skills` ve `~/.agents/skills` altında) "yabancı"
  sayılır; `--kontrol` bunları listeler. `--eskiyi-degistir` bayrağı (varsayılan değil) eski klasörü silmez,
  `~/.excel-skills-yedek/<tarih>/<araç>/` altına taşır, sonra bağlantıyı kurar. Faz 1'in sonunda bu bayrak bu
  bilgisayarda kullanıcı onayıyla bir kez çalıştırılır; "bitti" tanımı `--kontrol` çıktısında yedi klasörün
  "bağlı" görünmesidir.
- Geriye uyumluluk API düzeyindedir. 1.0.1 betiklerindeki sabit `~/.claude/skills/...` yolu Claude'da bağlantı
  sayesinde çalışmaya devam eder; diğer araçlar için README'de kısa bir göç notu (5.2'deki yol deseni) verilir.

### 5.5 Yeni temalar

Mevcut `Tema` veri sınıfına geriye uyumlu iki alan eklenir: `yazi_olcegi: float = 1.0` (sunum teması için) ve
`renksiz: bool = False` (siyah-beyaz baskı teması için). Renk değerleri öneridir; render edilip kontrast
denetiminden geçince kesinleşir.

| Tema | Kimlik | Öneri tokenlar (murekkep / kagit / vurgu / başlık yazısı) |
|---|---|---|
| `mimari` | Serif başlık, taş-gri kâğıt, ince çizgi, geniş boşluk | 262626 / F7F5F0 / 3D5A6C / Cambria |
| `teknik` | Antet bloğu, belirgin ızgara, eşit genişlikte rakam | 1A1A1A / FFFFFF / 0B6E99 / Calibri, rakamlar Consolas |
| `baski` | Renksiz yazıcıda okunur; seri ayrımı tarama deseni ve çizgi kalınlığıyla, durum ✓ ⚠ ✗ simgeleriyle | 000000 / FFFFFF / 000000 / Calibri |
| `sunum` | Ekrana yansıtma; büyük KPI, `yazi_olcegi=1.3`, yüksek satırlar | 111827 / FFFFFF / 7C3AED / Calibri Light |

Kurallar: vurgu beyaz zeminde ≥ 4.5:1 kontrast; başlık fontları Windows ve Mac Office'te bulunan fontlar;
`baski` temasında grafik dolguları desenli (`PatternFillProperties`), renk taşımaz. Yeni yardımcı:
`antet(ws, t, alan, alanlar)` (teknik paftadaki gibi proje / belge no / revizyon / tarih / hazırlayan bloğu),
her temada kullanılabilir.

### 5.6 Yeni skill: `excel-insaat-belgeleri`

**Karar (2026-11-20, kullanıcı): kural ağırlıklı skill.** İlk planda her belge için hazır üretici betik vardı
(`python3 scripts/<belge>.py ... --olcek ...`). Uygulamaya başlanınca bunun skill değil şablon ürünü olduğu
görüldü; kullanıcı kural ağırlıklı yolu seçti. `excel-insaat-belgeleri` üretici betik içermez: SKILL.md ve
`references/` altındaki yapı rehberi yapay zekâya belgeyi nasıl kuracağını anlatır, yapay zekâ her istekte kitabı
kullanıcının verisiyle `xlkit` üzerinden kendisi kurar. Kurallar: tüm toplamlar formül; oranlar Parametreler
sayfasındaki girdi hücrelerinde ve tanımlı adlarla (`KDV_ORANI` gibi); ara toplamlar `SUBTOTAL(9, …)` ve
eklenen satırı kapsayan `INDEX(…,ROW()-1)` bitişli aralık; `TEXT()` ile sayıyı metne çevirme yok; hata hücresi
yok. Bedeli: çıktı her istekte birebir aynı olmaz; doğruluk, skill'in "bitti" listesindeki denetimlerle
(LibreOffice/Excel ile hesaplatıp toplamları karşılaştırma) sağlanır. Aşağıdaki tablolar yapı rehberinin
özetidir. Faz 3a uygulama planı (`plans/2026-11-20-faz3a-teklif-kesif.md`) bu kararla geçersizdir.

| Belge | Sayfalar ve içerik |
|---|---|
| **Teklif dosyası** | Kapak, Teklif Mektubu (metin alanları parametreden), İcmal, Keşif, Koşullar ve imza alanı. İcmal zinciri: iş grubu tutarları → ara toplam → (isteğe bağlı) iskonto → KDV matrahı → KDV → genel toplam. Varsayılan: genel gider ve kâr birim fiyatlara dahil; Parametreler'deki bir anahtarla icmalde ayrı "Genel gider ve kâr" satırı açılabilir. Teminat icmalde yer almaz, Koşullar sayfasında metin olarak durur |
| **Keşif / birim fiyat teklif cetveli** | Sıra, Poz No, Tanım, Birim, Miktar, Birim Fiyat, Tutar; iş grubu başlık satırı + ara toplam (`SUBTOTAL(9, …)` ile genel toplam ara toplamları iki kez saymaz); icmal sayfası |
| **Metraj cetveli** | Kurum örneğindeki düzen: No, Poz No, Tanım, Birim, Adet, En, Boy, Yükseklik, Miktar (= çarpım), Toplam; minha (boşluk düşümü) poz bloğu içinde "Minha" etiketli eksi satırdır, poz toplamı satırların toplamıdır. Azı / Çoğu sütunlu klasik hakediş biçimi bu sürümde yok: kaynaklar sütun anlamında çelişiyor, iki bağımsız kaynakla netleşmeden eklenmez |
| **Portföy panosu** (`excel-dashboard` içinde) | Proje başına satır; bütçe, gerçekleşen, sapma, ilerleme; KPI kartları ve grafikler |
| **Şantiye takip tabloları** | Günlük rapor (tarih, hava, meslek bazında personel, makine, yapılan işler, notlar), aylık puantaj (gün ızgarası, hafta sonu biçimi), malzeme takip (sipariş, teslim, kalan) |

**Ölçek** ayrı şablon değil, aynı yapının ayarıdır:

| Ölçek | Keşif yapısı | Tipik boyut |
|---|---|---|
| `villa` | Tek seviye iş grubu, tek sayfa keşif | 8–12 iş grubu |
| `apartman` | Keşifte her blok ayrı Miktar sütunu (Blok A, Blok B, …) + Toplam Miktar sütunu; tutar toplam miktardan; icmalde blok başına tutar sütunu. Daire tipleri (tip, adet, net / brüt alan) Parametreler sayfasında bir tablo; icmal m² başına maliyeti buradan hesaplar | Blok başına icmal |
| `buyuk` | Üç seviye (Bölüm > İş grubu > Kalem), Excel satır gruplama (outline), taşeron paketi sütunu, revizyon sütunu, çok sayfalı icmal | Yüzlerce kalem |

Mühendislik veya maliyet anlamı taşıyan her hesap (metraj formülü, KDV tabanı, ara toplam mantığı) uygulama
sırasında kaynakla karşılaştırılıp doğrulanır. Yapıcının elindeki gerçek metraj ve hakediş örnekleri yalnız
salt okunur başvurudur: düzen ve formül mantığı onlardan doğrulanır, içerikleri (ad, sayı, kurum) depoya
kopyalanmaz.

### 5.7 SKILL.md düzenlemeleri

- Dört açıklama "Use when ..." ile başlar ve yalnız tetikleyici koşulu sayar; bağımlılık cümleleri gövdeye taşınır.
- Sabit yollar kaldırılır (5.2).
- "Bitti sayılmadan önce" listesine: hangi motorla denetlendiği yazıldı; LibreOffice ile denetlendiyse uyarı verildi.

### 5.8 Doğrulama

- `tests/` (pytest): tema kontrast denetimi; her üretici her tema × ölçek için hatasız dosya üretir; üretilen
  dosyada hata hücresi ve kırık başvuru yok; geriye uyumluluk (1.0.1 örnek panosu aynı API ile çalışır);
  kütüphane bulma (bağlantı ve kopya kurulumu).
- `kur.test.sh`: geçici bir HOME altında bağla / kopya / kontrol / kaldır ve yabancı klasöre dokunmama.
- Görsel denetim: her tema ve her belge için önizleme PNG'si alınır ve açılıp bakılır. Mac Excel ve
  (bu bilgisayara kurulan) LibreOffice ile yapılır.
- Windows: geliştirme bilgisayarında Windows ortamı yok. Windows yolu (COM, `kur.ps1`) birim testleri ve kod
  incelemesiyle yazılır, README'de "Windows: dış test bekliyor" diye işaretlenir. `docs/windows-test.md` adım
  adım bir test listesi verir (kurulum, `--kontrol`, örnek pano üretme, `excel_ile_kaydet`, `onizle`, beklenen
  çıktı ve ekran görüntüsü isteği); Windows kullanan biri bu listeyi çalıştırıp sonucu iletince işaret kalkar.
  Faz 1 bu dış test olmadan kapanabilir.
- Testler yerelde koşulur (GitHub Actions kullanılmıyor).

### 5.9 Sürüm ve aşamalar

Sürüm 2.0.0 (yeni skill ve kurulum değişikliği). Her faz ayrı commit dizisi ve kendi doğrulaması ile biter:

1. **Faz 1 · Altyapı:** yol düzeltmeleri, platform katmanı, `ortam.py`, `kur.sh` / `kur.ps1`, açıklama
   düzenlemesi, testler.
2. **Faz 2 · Temalar:** dört tema, `antet()`, örnek panonun yedi temadaki görüntüsü.
3. **Faz 3 · Özel sektör belgeleri:** `excel-insaat-belgeleri` kural ağırlıklı (3a: teklif dosyası ve keşif;
   3b: metraj, portföy panosu, şantiye takip), üç ölçek.
4. **Faz 4 · Kamu formatı:** (2026-11-20) kaynak taraması (KİK e-formları, Tebliğ 16.5 / 39, Yönetmelik md. 54, 57)
   sonrası kullanıcı "hazırlık + basılı cetvel" seçti. Kural ağırlıklı rehber `references/kamu-ekap.md` ve
   kurulmuş cetveli denetleyen `scripts/kik012_denetim.py` (şablon değil, denetim). Teklif mektubu EKAP e-formunda
   kalır. Doğrulanamayan: yeni EKAP'ta yapım e-teklifinde Excel içe/dışa aktarma ve azami hane.

README, kurulum ve "Bilinen sınırlar" her fazda güncellenir. Push ve GitHub işlemleri kullanıcı onayıyla.

## 6. Yapıcıya bırakılanlar

- Tema renklerinin son değerleri (kontrast ve render denetiminden sonra).
- LibreOffice'te tam yeniden hesaplatmayı garanti etmenin yolu (headless dönüşümde formüller varsayılan olarak
  yeniden hesaplanmayabilir; makro, yapılandırma ya da UNO ile çözülür).
- Windows'ta `kur.ps1` için junction / kopya geri düşüş ayrıntısı.
- Üretici betiklerin iç modül bölümlemesi (dosya başına ≤ 400 satır hedefi).
- Örnek verinin içeriği (kurgusal, gerçekçi büyüklükte).
