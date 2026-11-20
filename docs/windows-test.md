# Windows test listesi

Bu skill seti macOS'ta geliştirildi. Windows Excel yolu ve `kur.ps1` birim testleriyle ve macOS PowerShell ile
denendi ama gerçek bir Windows bilgisayarında henüz denenmedi. Bu liste Windows kullanan birinin yaklaşık
15 dakikada çalıştırması içindir. Sonuçlar gelince README'deki "Windows: dış test bekliyor" işareti kalkar.

## Gerekenler

- Windows 10 ya da 11, Microsoft Excel (Microsoft 365 ya da 2019+).
- Python 3.10+ ([python.org](https://www.python.org/downloads/), kurarken "Add python.exe to PATH" işaretli).
- Git ya da depoyu ZIP olarak indirme.

Komutlar PowerShell'de çalıştırılır. Her adımın çıktısını (ya da ekran görüntüsünü) saklayın.

## Adımlar

**1. Depoyu indirin**

```powershell
git clone https://github.com/hazarvolga/claude-excel-skills.git
cd claude-excel-skills
```

Beklenen: `skills`, `kur.ps1`, `README.md` klasörde görünür.

**2. Paketleri kurun**

```powershell
python -m pip install --user openpyxl pymupdf pillow pywin32
```

Beklenen: hata yok ("Successfully installed" ya da "Requirement already satisfied").

**3. Ortam raporu**

```powershell
python skills\excel-design-system\scripts\ortam.py
python skills\excel-design-system\scripts\ortam.py --json > ortam.json
```

Beklenen: `openpyxl`, `pymupdf`, `PIL`, `win32com` satırlarında sürüm; `Motorlar` satırında `excel-win`.
LibreOffice kuruluysa `libreoffice` da görünür.

**4. Kurulum durumu (değişiklik yapmaz)**

```powershell
powershell -ExecutionPolicy Bypass -File kur.ps1 -Kontrol
```

Beklenen: var olan araç klasörleri için satırlar; olmayanlar için "araç klasörü yok, atlandı". Hata yok.

**5. Kurulum**

```powershell
powershell -ExecutionPolicy Bypass -File kur.ps1
```

Beklenen: her skill için `bağlı (bağlantı)`, `bağlı (junction)` ya da `kopya (...)`. Hangisinin yazdığını not
edin: Geliştirici Modu kapalıysa junction ya da kopya beklenir, bu normaldir.

**6. Örnek pano üretin**

```powershell
python skills\excel-dashboard\scripts\ornek_pano.py "$env:USERPROFILE\Desktop\Test Panosu ş.xlsx" kurumsal
```

Beklenen: masaüstünde `Test Panosu ş.xlsx` oluşur, Excel'de açılınca 4 KPI kartı ve 4 grafik görünür.

**7. Excel ile hesaplatıp kaydetme**

```powershell
python -c "import sys; sys.path.insert(0, r'skills\excel-design-system\scripts'); import xlkit as xk; print(xk.excel_ile_kaydet(r'$env:USERPROFILE\Desktop\Test Panosu ş.xlsx'))"
```

Beklenen: çıktı `excel-win`. Excel penceresi açılmaz (görünmez çalışır); Excel'de önceden açık kitaplarınız
kapanmaz.

**8. Önizleme**

```powershell
python -c "import sys; sys.path.insert(0, r'skills\excel-design-system\scripts'); import xlkit as xk; print(xk.onizle(r'$env:USERPROFILE\Desktop\Test Panosu ş.xlsx', 'onizleme'))"
```

Beklenen: `onizleme\sayfa-01.png` ve diğer sayfaların yolları listelenir. `sayfa-01.png` açılınca pano
taşmasız ve okunur görünür.

**9. Kaldırma**

```powershell
powershell -ExecutionPolicy Bypass -File kur.ps1 -Kaldir
```

Beklenen: 5. adımda kurulanlar kaldırılır; araç klasörlerindeki başka skill'lere dokunulmaz.

## Geri gönderilecekler

1. 3. adımdaki `ortam.json` dosyası.
2. 5. ve 9. adımların ekran görüntüsü.
3. 7. adımın çıktısı.
4. `onizleme\sayfa-01.png`.
5. Hata veren adım olduysa: adım numarası ve hata metninin tamamı.

Kişisel dosya yolu veya kullanıcı adı içeren satırları göndermeden önce silebilirsiniz.
