<#
Excel skill setini (skills\ altındaki her skill) bilgisayardaki AI araçlarının skill klasörlerine kurar.
Varsayılan: sembolik bağlantı; olmazsa dizin bağlantısı (junction); o da olmazsa işaretli kopya.
Kullanım (PowerShell 5.1+ ya da pwsh):
  powershell -ExecutionPolicy Bypass -File kur.ps1                  # bağla (yalnız var olan araç klasörlerine)
  powershell -ExecutionPolicy Bypass -File kur.ps1 -Kopya           # kopyala (güncellemede tekrar çalıştır)
  powershell -ExecutionPolicy Bypass -File kur.ps1 -Kontrol         # yalnız durumu göster
  powershell -ExecutionPolicy Bypass -File kur.ps1 -Kaldir          # yalnız bu betiğin kurduklarını kaldır
  powershell -ExecutionPolicy Bypass -File kur.ps1 -EskiyiDegistir  # elle kopyalanmış eski klasörleri yedekleyip bağla
Başka bir skill'in üzerine asla yazmaz; aynı adda yabancı içerik varsa atlar ve bildirir.
#>
param([switch]$Kopya, [switch]$Kontrol, [switch]$Kaldir, [switch]$EskiyiDegistir)
$ErrorActionPreference = 'Stop'

$Root = $PSScriptRoot
$Mark = '.kur-excel-skills'
$Base = if ($env:USERPROFILE) { $env:USERPROFILE } else { $HOME }
$Yedek = Join-Path $Base (Join-Path '.excel-skills-yedek' (Get-Date -Format 'yyyyMMdd-HHmmss'))
$Targets = @('.claude/skills', '.agents/skills', '.codex/skills', '.gemini/skills', '.mlx-serve/skills',
             '.config/opencode/skills', '.qoder/skills') | ForEach-Object { Join-Path $Base $_ }

$Skills = Get-ChildItem -Directory (Join-Path $Root 'skills') | Where-Object { Test-Path (Join-Path $_.FullName 'SKILL.md') }
if (-not $Skills) { Write-Error "skills\ altında skill bulunamadı: $Root\skills"; exit 2 }

function Get-Hedef($t) { $i = Get-Item -LiteralPath $t -Force -ErrorAction SilentlyContinue; if ($i -and $i.LinkType) { [string]$i.Target } else { $null } }
function Test-Bizim($t, $src) {
  $hedef = Get-Hedef $t
  if ($hedef) { return ($hedef.TrimEnd('\', '/') -eq $src.TrimEnd('\', '/')) }
  return ((Test-Path -LiteralPath $t -PathType Container) -and (Test-Path -LiteralPath (Join-Path $t $Mark)))
}
function Test-EskiKopya($t) {
  (-not (Get-Hedef $t)) -and (Test-Path -LiteralPath $t -PathType Container) -and
  (-not (Test-Path -LiteralPath (Join-Path $t $Mark))) -and (Test-Path -LiteralPath (Join-Path $t 'SKILL.md'))
}
function Get-Durum($t, $src) {
  $hedef = Get-Hedef $t
  if ($hedef) { if (Test-Bizim $t $src) { return 'bağlı' } else { return "YABANCI BAĞLANTI -> $hedef" } }
  if (Test-Bizim $t $src) { return 'kopya' }
  if (Test-EskiKopya $t) { return 'ESKİ ELLE KOPYA (-EskiyiDegistir ile yedekleyip bağla)' }
  if (Test-Path -LiteralPath $t) { return 'YABANCI KLASÖR (dokunulmadı)' }
  return 'yok'
}
function Remove-Hedef($t) {
  if (Get-Hedef $t) { (Get-Item -LiteralPath $t -Force).Delete() } else { Remove-Item -LiteralPath $t -Recurse -Force }
}
function New-Kurulum($t, $src) {
  if ($Kopya) { Copy-Item -LiteralPath $src -Destination $t -Recurse; New-Item -ItemType File (Join-Path $t $Mark) | Out-Null; return 'kopya' }
  try { New-Item -ItemType SymbolicLink -Path $t -Target $src | Out-Null; return 'bağlantı' } catch {}
  try { New-Item -ItemType Junction -Path $t -Target $src | Out-Null; return 'junction' } catch {}
  Copy-Item -LiteralPath $src -Destination $t -Recurse; New-Item -ItemType File (Join-Path $t $Mark) | Out-Null
  return 'kopya (bağlantı izni yok: Geliştirici Modu ya da yönetici gerekir)'
}

$Failed = 0
foreach ($dir in $Targets) {
  if (-not (Test-Path -LiteralPath $dir -PathType Container)) { Write-Host ("  {0,-45} araç klasörü yok, atlandı" -f $dir); continue }
  foreach ($s in $Skills) {
    $t = Join-Path $dir $s.Name; $src = $s.FullName; $not = ''
    if ($Kaldir) { if (Test-Bizim $t $src) { Remove-Hedef $t } }
    elseif (-not $Kontrol) {
      if (Test-Path -LiteralPath $t) {
        if (Test-Bizim $t $src) { Remove-Hedef $t }
        elseif ($EskiyiDegistir -and (Test-EskiKopya $t)) {
          $hedefKlasor = Join-Path $Yedek (($dir.Substring($Base.Length).TrimStart('\', '/')) -replace '[\\/]', '_')
          New-Item -ItemType Directory -Force $hedefKlasor | Out-Null
          Move-Item -LiteralPath $t -Destination (Join-Path $hedefKlasor $s.Name)
        }
        else { $Failed++; Write-Host ("  {0,-45} {1,-22} {2}" -f $dir, $s.Name, (Get-Durum $t $src)); continue }
      }
      $not = " ($(New-Kurulum $t $src))"
    }
    Write-Host ("  {0,-45} {1,-22} {2}{3}" -f $dir, $s.Name, (Get-Durum $t $src), $not)
  }
}
if (Test-Path -LiteralPath $Yedek) { Write-Host "Eski kopyalar yedeklendi: $Yedek" }
if ($Failed -gt 0) { Write-Host "$Failed yerde aynı adda yabancı içerik var; elle bakın (bkz. -Kontrol)."; exit 1 }
