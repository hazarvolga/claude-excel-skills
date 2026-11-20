# Excel skill seti · Faz 1 (altyapı) Uygulama Planı

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Skill setini her AI aracında ve macOS / Windows / Linux'ta çalışır hale getirmek: sabit yolları kaldırmak, Excel'e bağlı adımlara Windows Excel ve LibreOffice motorlarını eklemek, tek kaynaktan yedi araç klasörüne kurulum.

**Architecture:** `xlkit.py` tek ortak kütüphane olarak `excel-design-system/scripts/` içinde kalır. Excel otomasyonu yeni `motor.py` modülüne taşınır; `excel_ile_kaydet()` ve `onizle()` imzaları korunur, içeride motor seçer. `kur.sh` / `kur.ps1` `skills/` altındaki her skill'i araç klasörlerine bağlar.

**Tech Stack:** Python 3.10+, openpyxl, pymupdf, pillow, pytest; pywin32 (yalnız Windows); LibreOffice `soffice --headless`; bash; PowerShell 5.1+.

**Spec:** `docs/superpowers/specs/2026-11-20-excel-skills-v2-design.md` (bölüm 5.1–5.4, 5.7, 5.8, 5.9 Faz 1)

## Global Constraints

- Hiçbir dosyada araç klasörü (`~/.claude`, `~/.Codex` vb.) sabit yazılmaz; yollar `Path(__file__).resolve()` ile bulunur.
- Skill `description` alanı "Use when ..." ile başlar ve yalnız tetikleyici koşulu sayar.
- Frontmatter yalnız `name`, `description`, `metadata`.
- Betikler `python3 scripts/x.py` ya da tam yolla çağrılır; `uv` varsa `uv run --no-project --with openpyxl --with pymupdf --with pillow python -I`.
- 1.0.1 API'si geriye uyumlu: mevcut fonksiyon adları ve parametreleri değişmez, yalnız isteğe bağlı parametre eklenir.
- Depoda belirli bir şirket veya ürün markası geçmez.
- Motor sırası: macOS Excel → Windows Excel → LibreOffice → yok. Sessiz geçiş yok; kullanılan motor döndürülür, motor yoksa uyarı.
- Testler yerelde: `uv run --no-project --with openpyxl --with pymupdf --with pillow --with pytest python -m pytest tests -q` ve `bash kur.test.sh`.
- Push yok. Commit'ler `v2-tasarim` dalında, mesajlar `<tür>: <açıklama>` + `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.

## Review Focus

1. **Türkçe karakterli ve boşluklu yol** (`Masaüstü/Teklif 2026 ş.xlsx`): her motor dosyayı açabilmeli → Task 2 testi.
2. **Excel kurulu ama otomasyon engelli** (Codex sandbox'ı, izin penceresi, `osascript` hatası): çökme değil, sonraki motora geçip hangi motorun kullanıldığını söylemeli → Task 2 testi.
3. **LibreOffice zaten açık** (kullanıcının açık bir belgesi var): headless çağrı kullanıcının oturumuna takılmamalı → yalıtılmış profil, Task 2 testi profil argümanını doğrular.
4. **LibreOffice ile kaydetmenin tasarımı bozması:** otomatik modda özgün dosyanın üzerine LibreOffice ile yazılmamalı → Task 2 testi.
5. **Kurulumun iki kez, başka bir klasörden ya da boşluklu depo yolundan çalıştırılması:** sonuç aynı olmalı → Task 4 testi.

---

### Task 1: Kütüphaneyi konumdan bulma ve SKILL.md düzeltmeleri

**Files:**
- Modify: `skills/excel-design-system/scripts/xlkit.py:1-12` (modül belgesi: sabit yol örneği kaldırılır)
- Modify: `skills/excel-dashboard/scripts/ornek_pano.py:14` (sabit `sys.path.insert` yerine konumdan bulma)
- Modify: `skills/excel-design-system/SKILL.md`, `skills/excel-charts/SKILL.md`, `skills/excel-dashboard/SKILL.md`
- Create: `tests/test_yol.py`, `tests/conftest.py`

**Interfaces:**
- Produces: betiklerde kullanılacak bulma deseni (her betiğe satır içi, ortak modül yok çünkü kütüphane yüklenmeden önce çalışmalı):
  ```python
  _BURA = Path(__file__).resolve().parent
  _ADAYLAR = (_BURA.parent.parent / "excel-design-system" / "scripts", _BURA)
  ```
  İlk `xlkit.py` içeren aday `sys.path`'e eklenir; hiçbiri yoksa `SystemExit("excel-design-system bulunamadı; skill setini kur.sh / kur.ps1 ile birlikte kurun.")`.
- Produces: `tests/conftest.py` → fixture `skill_kopyasi(tmp_path, kip: "bagla"|"kopya") -> Path` (skills/ ağacını tmp altına bağlantı ya da kopya olarak kurar, araç klasörü kökünü döndürür).

- [ ] **Step 1: Failing testler** (`tests/test_yol.py`)
  - `test_sabit_arac_yolu_yok`: `skills/**` altındaki tüm `.md` ve `.py` dosyalarında `~/.claude`, `~/.Codex`, `~/.codex`, `.agents/skills` geçmez.
  - `test_aciklamalar_use_when`: üç SKILL.md'nin `description` değeri `"Use when"` ile başlar; `xlkit`, `Base layer`, `Builds on`, `Uses ` sözcüklerini içermez.
  - `test_ornek_pano_kopya_kurulumda_calisir` ve `..._bagla_kurulumda_calisir`: fixture ile kurulan ağaçta `python -I <kok>/excel-dashboard/scripts/ornek_pano.py <tmp>/p.xlsx minimal` çıkış kodu 0, dosya oluşur.
  - `test_kardes_yoksa_acik_hata`: `excel-design-system` silinmiş kopya ağaçta aynı komut çıkış ≠ 0 ve stderr `excel-design-system bulunamadı` içerir.
- [ ] **Step 2:** `pytest tests/test_yol.py -q` → 5 FAIL.
- [ ] **Step 3:** `ornek_pano.py` bulma desenini uygula; xlkit modül belgesini desenle güncelle.
- [ ] **Step 4:** SKILL.md'ler: kod örneği "Bu SKILL.md'nin bulunduğu klasör `SKILL_DIR`" diliyle; charts / dashboard için `SKILL_DIR/../excel-design-system/scripts`. Açıklamalardan bağımlılık cümleleri gövdeye taşınır ("Bu skill excel-design-system içindeki xlkit.py'yi kullanır"). Açıklamalardaki Türkçe tetikleyici ifadeler korunur.
- [ ] **Step 5:** `pytest tests/test_yol.py -q` → PASS.
- [ ] **Step 6: Commit** `fix: sabit skill yolları kaldırıldı, kütüphane konumdan bulunur`

### Task 2: Motor katmanı (macOS Excel, Windows Excel, LibreOffice)

**Files:**
- Create: `skills/excel-design-system/scripts/motor.py`
- Modify: `skills/excel-design-system/scripts/xlkit.py:512-604` (`_applescript`, `_excel_betigi` motor.py'ye taşınır; `excel_ile_kaydet`, `onizle` motoru çağırır)
- Create: `tests/test_motor.py`

**Interfaces:**
- Produces (`motor.py`):
  - `MOTORLAR = ("excel-mac", "excel-win", "libreoffice")`
  - `bulunan_motorlar() -> list[str]` (sırayla; `excel-mac`: darwin + `/Applications/Microsoft Excel.app`; `excel-win`: win32 + `win32com` içe aktarılabiliyor + `Excel.Application` kayıtlı; `libreoffice`: `soffice` PATH'te ya da `/Applications/LibreOffice.app/Contents/MacOS/soffice` ya da `C:\Program Files\LibreOffice\program\soffice.exe`)
  - `hesapla_kaydet(xlsx: str, motor: str | None = None) -> str | None` → kullanılan motor
  - `pdf_al(xlsx: str, pdf: str, motor: str | None = None) -> str | None` → kullanılan motor
  - `class MotorHatasi(RuntimeError)`
- Produces (`xlkit.py`, geriye uyumlu):
  - `excel_ile_kaydet(xlsx: str, temizle: bool = True, motor: str | None = None) -> str | None`
  - `onizle(xlsx: str, cikti_dizin: str, dpi: int = 150, max_genislik: int = 2000, motor: str | None = None) -> list[str]`

Kararlar:
- `motor=None` (otomatik): sırayla dener; bir motor `MotorHatasi`, `subprocess.CalledProcessError` ya da `OSError` verirse `warnings.warn` ile nedenini yazıp sıradakine geçer. Hepsi başarısızsa / hiç yoksa `warnings.warn("Hiçbir hesaplama motoru bulunamadı: görsel denetim yapılamadı")`, `None` / `[]` döner.
- `motor="..."` açıkça verilirse yalnız o denenir, hata yükseltilir.
- **LibreOffice ve kaydetme:** otomatik modda `hesapla_kaydet` LibreOffice'i atlar (özgün dosyanın biçimini bozabilir) ve uyarır; yalnız `motor="libreoffice"` açıkça verilirse özgün dosyayı LibreOffice çıktısıyla değiştirir. `pdf_al` otomatik modda LibreOffice'i kullanır (özgün dosyaya dokunmaz).
- **LibreOffice yeniden hesaplama:** her çağrı geçici bir profil kullanır: `-env:UserInstallation=file:///<tmp>/lo-profil`; profilin `user/registrymodifications.xcu` dosyasına `/org.openoffice.Office.Calc/Formula/Load` altında `OOXMLRecalcMode = 0` (her zaman yeniden hesapla) yazılır. Komut: `soffice -env:... --headless --norestore --convert-to pdf --outdir <dizin> <xlsx>` (kaydetme için `--convert-to xlsx:"Calc MS Excel 2007 XML"`), zaman aşımı 180 sn.
- **Windows Excel:** `_excel_win(xlsx, pdf=None, dispatch=None)`; `dispatch` varsayılanı `win32com.client.DispatchEx`. `Visible=False`, `DisplayAlerts=False`, `Workbooks.Open(abs)`, `Application.CalculateFull()`, `pdf` varsa `ExportAsFixedFormat(0, pdf)` yoksa `Save()`, `finally` içinde yalnız açılan kitap `Close(False)` ve kendi açtığı uygulama `Quit()`.
- **macOS Excel:** bugünkü AppleScript davranışı aynen (aynı adlı kitabı kapatma, -50'de bir yeniden deneme, PDF tarih denetimi).

- [ ] **Step 1: Failing testler** (`tests/test_motor.py`)
  - `test_otomatik_sira_ve_gecis`: `bulunan_motorlar` → `["excel-mac", "libreoffice"]` olacak şekilde monkeypatch; `excel-mac` `CalledProcessError` verir; `pdf_al(..., motor=None) == "libreoffice"` ve bir `UserWarning` içinde `excel-mac` geçer.
  - `test_motor_yoksa_uyari_ve_bos`: `bulunan_motorlar → []`; `xlkit.onizle(x, d) == []`, uyarı metni `görsel denetim yapılamadı`; `xlkit.excel_ile_kaydet(x) is None`.
  - `test_libreoffice_otomatik_kaydetmede_atlanir`: yalnız `libreoffice` var; `hesapla_kaydet(x)` `None` döner, dosyanın baytları değişmez, uyarı `LibreOffice` içerir.
  - `test_libreoffice_komutu_yalitilmis_profil`: `subprocess.run` yakalanır; argümanlarda `-env:UserInstallation=file://` ve `--headless` var; profil xcu dosyasında `OOXMLRecalcMode` ve `<value>0</value>`.
  - `test_excel_win_cagri_sirasi`: sahte COM nesnesiyle `_excel_win(x, pdf=p, dispatch=sahte)` → çağrılar sırasıyla `Open`, `CalculateFull`, `ExportAsFixedFormat(0, p)`, `Close(False)`, `Quit`; `Open` hata verse de `Quit` çağrılır.
  - `@skipif(soffice yok)` `test_libreoffice_gercek_hesaplama_turkce_yol`: `tmp/"Teklif 2026 ş"/"çıktı.xlsx"` içinde `A1=2, A2="=A1*21"`; `hesapla_kaydet(x, motor="libreoffice")` sonra `openpyxl.load_workbook(x, data_only=True)["Sheet"]["A2"].value == 42`; `pdf_al` ile PDF oluşur ve boş değildir.
  - `@skipif(excel-mac yok)` `test_excel_mac_onizle_turkce_yol`: aynı yolla `onizle` en az bir PNG döndürür.
- [ ] **Step 2:** `pytest tests/test_motor.py -q` → FAIL (motor modülü yok).
- [ ] **Step 3:** `motor.py`'yi yukarıdaki kararlarla yaz; xlkit'teki iki fonksiyonu ona bağla. `onizle` PDF'i `motor.pdf_al` ile alır, PNG çizim kısmı aynen kalır.
- [ ] **Step 4:** `pytest tests -q` → PASS (Mac'te LibreOffice ve Excel testleri de koşar; ikisi de kuruluysa atlanan test olmamalı).
- [ ] **Step 5:** Elle görsel denetim: `ornek_pano.py` → `onizle(motor="excel-mac")` ve `onizle(motor="libreoffice")` PNG'leri açılıp karşılaştırılır; LibreOffice'teki farklar `skills/excel-charts/references/openpyxl-grafik-tuzaklari.md` sonuna "LibreOffice önizleme farkları" başlığıyla yazılır.
- [ ] **Step 6: Commit** `feat: motor katmanı (Windows Excel, LibreOffice) ve otomatik geçiş`

### Task 3: Ortam raporu ve SKILL.md motor anlatımı

**Files:**
- Create: `skills/excel-design-system/scripts/ortam.py`
- Modify: `skills/excel-design-system/SKILL.md` ("macOS Excel otomasyonu" bölümü → "Hesaplama ve önizleme motorları"; "Bitti sayılmadan önce" listesine iki madde)
- Test: `tests/test_ortam.py`

**Interfaces:**
- Consumes: `motor.bulunan_motorlar()`
- Produces: `rapor() -> dict` anahtarları `python` (str), `paketler` (dict ad→sürüm|None: openpyxl, pymupdf, PIL, win32com), `motorlar` (list[str]), `eksikler` (list[str], insan okur önerileri). CLI: `python3 scripts/ortam.py` okunur Türkçe rapor yazar, `--json` JSON yazar; zorunlu `openpyxl` yoksa çıkış 1.

- [ ] **Step 1: Failing test:** `test_rapor_anahtarlari` (anahtarlar ve türler), `test_openpyxl_yoksa_cikis_1` (`python -I -c` ile `sys.modules["openpyxl"]=None` yapan sarmalayıcı üzerinden), `test_json_cikti_gecerli`.
- [ ] **Step 2:** FAIL doğrula.
- [ ] **Step 3:** `ortam.py` yaz. `eksikler` örnekleri: `"Görsel denetim için LibreOffice kurun: https://www.libreoffice.org"`, `"pip: python3 -m pip install --user openpyxl pymupdf pillow"`.
- [ ] **Step 4:** SKILL.md: ilk kullanımda `python3 SKILL_DIR/scripts/ortam.py` çalıştır; motor sırası ve LibreOffice'in kaydetmede yalnız açıkça seçilince kullanıldığı; Bitti listesine "Hangi motorla denetlendiği yazıldı" ve "LibreOffice ile denetlendiyse 'Excel'de grafik görünümü küçük farklar gösterebilir' uyarısı verildi".
- [ ] **Step 5:** PASS doğrula.
- [ ] **Step 6: Commit** `feat: ortam.py raporu, SKILL.md motor anlatımı`

### Task 4: Kurulum betikleri (`kur.sh`, `kur.ps1`)

**Files:**
- Create: `kur.sh`, `kur.test.sh`, `kur.ps1`

**Interfaces:**
- Produces: `bash kur.sh [--kopya | --kontrol | --kaldir | --eskiyi-degistir]`; `pwsh -File kur.ps1 [-Kopya | -Kontrol | -Kaldir | -EskiyiDegistir]`. Skill listesi `skills/*/SKILL.md` olan her klasör (Faz 3'te dördüncü skill kendiliğinden girer). İşaret dosyası `.kur-excel-skills`. Yedek kökü `~/.excel-skills-yedek/<YYYYMMDD-HHMMSS>/<araç-klasörü, / → _>/<skill>`.

Davranış, örnek alınan başka bir skill deposundaki `kur.sh` ile aynıdır (yalnız var olan araç klasörleri, yabancıya dokunmama, çıkış 1), farklar: çoklu skill; `--eskiyi-degistir` yalnız işaretsiz düz klasörleri (bağlantı değil) yedeğe taşıyıp bağlar, yabancı bağlantılara yine dokunmaz. `kur.ps1`: hedefler `$env:USERPROFILE` altında aynı göreli yollar; önce `New-Item -ItemType SymbolicLink`, olmazsa `-ItemType Junction`, o da olmazsa `Copy-Item -Recurse` + işaret dosyası; hangisinin kullanıldığını yazar.

- [ ] **Step 1: Failing test** `kur.test.sh` (örnek testin tüm durumları, her skill için) artı:
  - `eskiyi-degistir: işaretsiz eski klasör yedeğe taşınır ve bağlanır` (yedek kökünde SKILL.md var, hedef bağlantı)
  - `eskiyi-degistir: yabancı bağlantıya dokunmaz`
  - `bağla: dört değil, skills/ altındaki her skill kurulur` (geçici bir `skills/deneme-skill/SKILL.md` ile, test sonunda silinir)
  - `bağla: boşluklu depo yolundan ve başka cwd'den çalışır` (depo `"$FAKE/depo yolu"` altına kopyalanıp oradan çağrılır, iki kez)
- [ ] **Step 2:** `bash kur.test.sh` → FAIL.
- [ ] **Step 3:** `kur.sh` yaz.
- [ ] **Step 4:** `bash kur.test.sh` → `Sonuç: N geçti, 0 kaldı`.
- [ ] **Step 5:** `kur.ps1` yaz; bu bilgisayarda Windows yok: `pwsh` kuruluysa `pwsh -NoProfile -Command "[ScriptBlock]::Create((Get-Content kur.ps1 -Raw)) | Out-Null"` ile yalnız sözdizimi denetlenir, değilse atlanır ve not edilir.
- [ ] **Step 6: Commit** `feat: kur.sh ve kur.ps1 ile yedi araç klasörüne kurulum`

### Task 5: Belgeler, Windows test listesi, sürüm ve yerel kurulum

**Files:**
- Create: `docs/windows-test.md`
- Modify: `README.md` (Kurulum: `kur.sh` / `kur.ps1`; Gereksinimler: motorlar; desteklenen araçlar tablosu; Bilinen sınırlar: "Windows: dış test bekliyor"; 1.0.1'den geçiş notu; Sürüm notları 1.1.0), üç SKILL.md `metadata.version: "1.1.0"`

Sürüm kararı: Faz 1 geriye uyumlu ve eklemeli olduğu için 1.1.0; spec'teki 2.0.0, dördüncü skill eklenince (Faz 3) verilir.

- [ ] **Step 1:** `docs/windows-test.md`: adımlar (depoyu indir, `pwsh -File kur.ps1 -Kontrol`, `kur.ps1`, `python scripts\ortam.py`, örnek pano üret, `excel_ile_kaydet`, `onizle`), her adımda beklenen çıktı, geri gönderilecekler (ekran görüntüsü, `ortam.py --json` çıktısı, `onizleme\sayfa-01.png`).
- [ ] **Step 2:** README ve sürüm alanları güncellenir; `test_sabit_arac_yolu_yok` README'yi de tarayacak şekilde genişletilir (README'deki kurulum örnekleri yalnız `kur.sh` / `kur.ps1` gösterir).
- [ ] **Step 3:** Tüm testler: `pytest tests -q` ve `bash kur.test.sh` → hepsi geçer.
- [ ] **Step 4: Commit** `docs: kurulum, motorlar, Windows test listesi; sürüm 1.1.0`
- [ ] **Step 5 (kullanıcı onayıyla):** bu bilgisayarda `bash kur.sh --kontrol` çıktısı gösterilir, onay alınınca `bash kur.sh --eskiyi-degistir`; ardından `--kontrol` yedi klasörde de üç skill için `bağlı` göstermelidir. Eski kopyalar `~/.excel-skills-yedek/` altında durur.
