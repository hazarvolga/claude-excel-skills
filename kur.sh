#!/usr/bin/env bash
# Excel skill setini (skills/ altındaki her skill) bilgisayardaki AI araçlarının skill klasörlerine kurar.
# Varsayılan: sembolik bağlantı (tek kaynak, güncelleme her araca anında yansır).
# Kullanım:
#   bash kur.sh                    # bağla (yalnız var olan araç klasörlerine)
#   bash kur.sh --kopya            # bağlantı izlemeyen araçlar için kopyala (güncellemede tekrar çalıştır)
#   bash kur.sh --kontrol          # her klasörde ne durduğunu göster, hiçbir şey değiştirme
#   bash kur.sh --kaldir           # yalnız bu betiğin kurduğu bağlantı/kopyaları kaldır
#   bash kur.sh --eskiyi-degistir  # elle kopyalanmış eski klasörleri ~/.excel-skills-yedek/ altına taşıyıp bağla
# Başka bir skill'in üzerine asla yazmaz; aynı adda yabancı içerik varsa atlar ve bildirir.
set -uo pipefail

ROOT="$(cd "$(dirname "$0")" && pwd)"
MARK=".kur-excel-skills"   # kopya kurulumlarını tanımak için işaret dosyası
YEDEK="$HOME/.excel-skills-yedek/$(date +%Y%m%d-%H%M%S)"
TARGETS=(
  "$HOME/.claude/skills"          # Claude Code / Claude Desktop
  "$HOME/.agents/skills"          # Codex, Gemini CLI, Copilot CLI ortak takma ad
  "$HOME/.codex/skills"           # Codex
  "$HOME/.gemini/skills"          # Antigravity / Gemini
  "$HOME/.mlx-serve/skills"       # MLX-Serve
  "$HOME/.config/opencode/skills" # OpenCode
  "$HOME/.qoder/skills"           # Qoder
)

MODE="bagla"
case "${1:-}" in
  "") ;;
  --kopya) MODE="kopya" ;;
  --kontrol) MODE="kontrol" ;;
  --kaldir) MODE="kaldir" ;;
  --eskiyi-degistir) MODE="eski" ;;
  -h|--help) sed -n '2,10p' "$0" | sed 's/^# \{0,1\}//'; exit 0 ;;
  *) echo "Bilinmeyen argüman: $1" >&2; exit 2 ;;
esac

SKILLS=()
for d in "$ROOT"/skills/*/; do [ -f "$d/SKILL.md" ] && SKILLS+=("$(basename "$d")"); done
[ "${#SKILLS[@]}" -gt 0 ] || { echo "skills/ altında skill bulunamadı: $ROOT/skills" >&2; exit 2; }

ours() { # hedef bu skill'e ait mi: bizim bağlantımız ya da işaretli kopyamız
  local t="$1" src="$2"
  if [ -L "$t" ]; then [ "$(readlink "$t")" = "$src" ]; return; fi
  [ -d "$t" ] && [ -f "$t/$MARK" ]
}

eski_kopya() { # elle kopyalanmış, işaretsiz, SKILL.md'li düz klasör (bağlantı değil)
  local t="$1"; [ ! -L "$t" ] && [ -d "$t" ] && [ ! -f "$t/$MARK" ] && [ -f "$t/SKILL.md" ]
}

state() {
  local t="$1" src="$2"
  if [ -L "$t" ] && [ "$(readlink "$t")" = "$src" ]; then echo "bağlı"
  elif [ -L "$t" ]; then echo "YABANCI BAĞLANTI -> $(readlink "$t")"
  elif [ -d "$t" ] && [ -f "$t/$MARK" ]; then
    if diff -rq -x "$MARK" -x __pycache__ "$src" "$t" >/dev/null 2>&1; then echo "kopya (güncel)"
    else echo "kopya (ESKİ, --kopya ile yenile)"; fi
  elif eski_kopya "$t"; then echo "ESKİ ELLE KOPYA (--eskiyi-degistir ile yedekleyip bağla)"
  elif [ -e "$t" ]; then echo "YABANCI KLASÖR (dokunulmadı)"
  else echo "yok"; fi
}

yedekle() { # eski kopyayı yedek köküne taşır: <yedek>/<araç klasörü, / → _>/<skill>
  local t="$1" dir="$2" name="$3" hedef
  hedef="$YEDEK/$(printf '%s' "${dir#"$HOME"/}" | tr '/' '_')"
  mkdir -p "$hedef" && mv "$t" "$hedef/$name"
}

FAILED=0
for dir in "${TARGETS[@]}"; do
  if [ ! -d "$dir" ]; then printf '  %-45s araç klasörü yok, atlandı\n' "$dir"; continue; fi
  for name in "${SKILLS[@]}"; do
    t="$dir/$name"; src="$ROOT/skills/$name"
    case "$MODE" in
      kontrol) ;;
      kaldir) if ours "$t" "$src"; then rm -rf "$t"; fi ;;
      bagla|kopya|eski)
        if [ -e "$t" ] || [ -L "$t" ]; then
          if ours "$t" "$src"; then rm -rf "$t"
          elif [ "$MODE" = "eski" ] && eski_kopya "$t"; then yedekle "$t" "$dir" "$name" || { FAILED=$((FAILED + 1)); continue; }
          else FAILED=$((FAILED + 1)); printf '  %-45s %-22s %s\n' "$dir" "$name" "$(state "$t" "$src")"; continue
          fi
        fi
        if [ "$MODE" = "kopya" ]; then cp -R "$src" "$t" && rm -rf "$t/scripts/__pycache__" && : > "$t/$MARK"
        else ln -s "$src" "$t"; fi ;;
    esac
    printf '  %-45s %-22s %s\n' "$dir" "$name" "$(state "$t" "$src")"
  done
done

[ -d "$YEDEK" ] && echo "Eski kopyalar yedeklendi: $YEDEK"
[ "$FAILED" -eq 0 ] || { echo "$FAILED yerde aynı adda yabancı içerik var; elle bakın (bkz. --kontrol)."; exit 1; }
