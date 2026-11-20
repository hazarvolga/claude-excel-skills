#!/usr/bin/env bash
# kur.sh davranış testleri. Sahte bir HOME altında çalışır; gerçek araç klasörlerine dokunmaz.
# Kullanım: bash kur.test.sh
set -uo pipefail

HERE="$(cd "$(dirname "$0")" && pwd)"
FAKE="$(cd "$(mktemp -d "${TMPDIR:-/tmp}/kur-test.XXXXXX")" && pwd)"   # pwd: TMPDIR'deki // biçimini düzeltir
trap 'rm -rf "$FAKE" "$HERE/skills/deneme-skill"' EXIT
PASS=0; FAIL=0
check() { if [ "$2" = "0" ]; then PASS=$((PASS + 1)); echo "PASS  $1"; else FAIL=$((FAIL + 1)); echo "FAIL  $1"; fi; }
k() { HOME="$FAKE" bash "$HERE/kur.sh" "$@" >/dev/null 2>&1; }
ok() { "$@" && echo 0 || echo 1; }
SKILLLER=(excel-design-system excel-charts excel-dashboard excel-insaat-belgeleri)
hepsi_bagli() { local d="$1" s; for s in "${SKILLLER[@]}"; do [ "$(readlink "$d/$s")" = "$HERE/skills/$s" ] || return 1; done; }

mkdir -p "$FAKE/.claude/skills" "$FAKE/.qoder/skills" "$FAKE/.codex/skills"
mkdir -p "$FAKE/.codex/skills/excel-charts"; echo baska > "$FAKE/.codex/skills/excel-charts/SKILL.md"

k; RC=$?
check "bağla: her skill var olan klasöre bağlanır" "$(ok hepsi_bagli "$FAKE/.claude/skills")"
check "bağla: olmayan araç klasörü oluşturulmaz" "$(ok test ! -e "$FAKE/.gemini")"
check "bağla: yabancı klasörün üzerine yazılmaz" "$(ok grep -q baska "$FAKE/.codex/skills/excel-charts/SKILL.md")"
check "bağla: aynı klasördeki diğer skill'ler yine bağlanır" "$(ok test -L "$FAKE/.codex/skills/excel-dashboard")"
check "bağla: yabancı içerik varsa çıkış 1" "$(ok test "$RC" -eq 1)"

k; check "bağla: ikinci çalıştırma bağlantıyı korur" "$(ok hepsi_bagli "$FAKE/.qoder/skills")"

mkdir -p "$HERE/skills/deneme-skill"; echo "---" > "$HERE/skills/deneme-skill/SKILL.md"
k; check "bağla: skills/ altındaki her skill kurulur" "$(ok test -L "$FAKE/.qoder/skills/deneme-skill")"
k --kaldir; rm -rf "$HERE/skills/deneme-skill"

k --kopya
check "kopya: bağlantı yerine işaretli kopya" "$(ok test -d "$FAKE/.qoder/skills/excel-dashboard" -a ! -L "$FAKE/.qoder/skills/excel-dashboard" -a -f "$FAKE/.qoder/skills/excel-dashboard/.kur-excel-skills")"

OUT="$(HOME="$FAKE" bash "$HERE/kur.sh" --kontrol 2>&1)"
check "kontrol: güncel kopyayı bildirir" "$(printf '%s' "$OUT" | grep -q 'kopya (güncel)' && echo 0 || echo 1)"

k --kaldir
check "kaldır: bizim kurulumu siler" "$(ok test ! -e "$FAKE/.qoder/skills/excel-dashboard")"
check "kaldır: yabancı klasöre dokunmaz" "$(ok test -f "$FAKE/.codex/skills/excel-charts/SKILL.md")"

# eski elle kopyalanmış (işaretsiz) klasör ve yabancı bağlantı
mkdir -p "$FAKE/.claude/skills/excel-design-system"; echo eski > "$FAKE/.claude/skills/excel-design-system/SKILL.md"
mkdir -p "$FAKE/baska-yer"; ln -s "$FAKE/baska-yer" "$FAKE/.qoder/skills/excel-charts"
k --eskiyi-degistir
check "eskiyi-degistir: işaretsiz eski klasör bağlanır" "$(ok test "$(readlink "$FAKE/.claude/skills/excel-design-system")" = "$HERE/skills/excel-design-system")"
check "eskiyi-degistir: eski klasör yedeğe taşınır" "$(ok sh -c "grep -rqs eski '$FAKE/.excel-skills-yedek'")"
check "eskiyi-degistir: yabancı bağlantıya dokunmaz" "$(ok test "$(readlink "$FAKE/.qoder/skills/excel-charts")" = "$FAKE/baska-yer")"
k --kaldir; rm -f "$FAKE/.qoder/skills/excel-charts"

# boşluklu depo yolu ve başka cwd
mkdir -p "$FAKE/depo yolu"; cp -R "$HERE/kur.sh" "$HERE/skills" "$FAKE/depo yolu/"
( cd / && HOME="$FAKE" bash "$FAKE/depo yolu/kur.sh" >/dev/null 2>&1; HOME="$FAKE" bash "$FAKE/depo yolu/kur.sh" >/dev/null 2>&1 )
check "bağla: boşluklu depo yolundan ve başka cwd'den, iki kez" "$(ok test "$(readlink "$FAKE/.qoder/skills/excel-charts")" = "$FAKE/depo yolu/skills/excel-charts")"

# kırık bağlantı (depo taşındı): yabancı sayılır, kurulum yarıda kalmaz, diğer skill'ler kurulur
HOME="$FAKE" bash "$FAKE/depo yolu/kur.sh" --kaldir >/dev/null 2>&1; k --kaldir; ln -s "$FAKE/silinmis-depo/excel-charts" "$FAKE/.qoder/skills/excel-charts"
k; check "kur.sh kırık bağlantı: dokunmaz, diğerlerini kurar" "$(ok test -L "$FAKE/.qoder/skills/excel-dashboard" -a "$(readlink "$FAKE/.qoder/skills/excel-charts")" = "$FAKE/silinmis-depo/excel-charts")"
k --kaldir
if command -v pwsh >/dev/null 2>&1; then
  HOME="$FAKE" pwsh -NoProfile -File "$HERE/kur.ps1" >/dev/null 2>&1
  check "kur.ps1 kırık bağlantı: yarıda kalmaz, diğerlerini kurar" "$(ok test -L "$FAKE/.qoder/skills/excel-dashboard" -a "$(readlink "$FAKE/.qoder/skills/excel-charts")" = "$FAKE/silinmis-depo/excel-charts")"
  HOME="$FAKE" pwsh -NoProfile -File "$HERE/kur.ps1" -Kaldir >/dev/null 2>&1
fi
rm -f "$FAKE/.qoder/skills/excel-charts"

echo; echo "Sonuç: $PASS geçti, $FAIL kaldı"
[ "$FAIL" -eq 0 ]
