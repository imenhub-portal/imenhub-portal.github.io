#!/bin/bash
#
# i-Nventori Ofis — one-shot setup for a new Mac.
#
# Run it either way:
#   • from inside an existing clone  — just configures this machine
#   • standalone (curl one-liner)    — clones first, then configures
#
# Safe to re-run. It never force-pushes, never resets, and never overwrites
# a config file that already exists.
#
set -u

REPO_URL="https://github.com/imenhub-portal/imenhub-portal.github.io.git"
REPO_DIR_DEFAULT="$HOME/Documents/imenhub-portal.github.io"
APP="i-Nventoriofis"

bold() { printf '\033[1m%s\033[0m\n' "$1"; }
ok()   { printf '  \033[32m✓\033[0m %s\n' "$1"; }
warn() { printf '  \033[33m!\033[0m %s\n' "$1"; }
bad()  { printf '  \033[31m✗\033[0m %s\n' "$1"; }
step() { printf '\n\033[1m%s\033[0m\n' "$1"; }

echo
bold "i-Nventori Ofis — setup Mac"
echo "─────────────────────────────────────────────"

# ── 1. Prerequisites ────────────────────────────────────────────────
step "1. Semak keperluan"

if ! command -v git >/dev/null 2>&1; then
  bad "git tiada. Buka Terminal dan jalankan: xcode-select --install"
  echo; read -r -p "Tekan Enter untuk tutup..." _ </dev/tty; exit 1
fi
ok "git $(git --version | awk '{print $3}')"

if command -v node >/dev/null 2>&1; then
  ok "node $(node -v)"
else
  warn "node tiada — ujian tidak boleh dijalankan (apps tetap berfungsi)."
  warn "Pasang kemudian: brew install node"
fi

# ── 2. Locate or clone the repo ─────────────────────────────────────
step "2. Cari atau clone repo"

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO=""

# Already inside a clone?
if git -C "$SCRIPT_DIR" rev-parse --show-toplevel >/dev/null 2>&1; then
  REPO="$(git -C "$SCRIPT_DIR" rev-parse --show-toplevel)"
  ok "Guna clone sedia ada: $REPO"
elif [ -d "$REPO_DIR_DEFAULT/.git" ]; then
  REPO="$REPO_DIR_DEFAULT"
  ok "Jumpa clone sedia ada: $REPO"
else
  echo "  Repo belum ada. Clone ke:"
  echo "    $REPO_DIR_DEFAULT"
  mkdir -p "$(dirname "$REPO_DIR_DEFAULT")"
  if git clone "$REPO_URL" "$REPO_DIR_DEFAULT"; then
    REPO="$REPO_DIR_DEFAULT"
    ok "Selesai clone"
  else
    bad "Clone gagal — semak sambungan internet."
    echo; read -r -p "Tekan Enter untuk tutup..." _ </dev/tty; exit 1
  fi
fi

cd "$REPO" || exit 1

# ── 3. The excludes that do NOT travel with the repo ────────────────
step "3. Betulkan peraturan abaikan tempatan"

# .git/info/exclude is per-clone and never pushed, so _local/ and _backend/
# arrive UNIGNORED on a fresh machine. _backend/ holds live Apps Script
# secrets in some sibling apps and this repo is PUBLIC, so this matters.
EXCL=".git/info/exclude"
touch "$EXCL"
for rule in "_local/" "_backend/" ".claude/"; do
  if grep -qxF "$rule" "$EXCL" 2>/dev/null; then
    ok "sudah diabaikan: $rule"
  else
    printf '%s\n' "$rule" >> "$EXCL"
    ok "ditambah: $rule"
  fi
done

# ── 4. Local dev-server config (not tracked) ────────────────────────
step "4. Sediakan pelayan setempat"

mkdir -p "$APP/.claude"
LAUNCH="$APP/.claude/launch.json"
if [ -f "$LAUNCH" ]; then
  ok "launch.json sudah ada — tidak diganti"
else
  cat > "$LAUNCH" <<'JSON'
{
  "version": "0.0.1",
  "configurations": [
    {
      "name": "i-nventori-static",
      "runtimeExecutable": "npx",
      "runtimeArgs": ["--yes", "http-server", "-p", "8125", "-c-1", "."],
      "port": 8125
    }
  ]
}
JSON
  ok "launch.json dicipta"
fi

# ── 5. Pull latest ──────────────────────────────────────────────────
step "5. Tarik versi terkini"

# --autostash so a dirty tree does not block the pull; --rebase to keep
# history linear, which this shared monorepo expects.
if git pull --rebase --autostash 2>&1 | tail -3; then
  ok "Terkini: $(git log --oneline -1)"
else
  warn "Pull tidak selesai — lihat mesej di atas."
fi

# ── 6. Prove it works ───────────────────────────────────────────────
step "6. Jalankan ujian"

if command -v node >/dev/null 2>&1; then
  cd "$APP" || exit 1
  B="$(node tests/test_backend.js 2>&1 | tail -1)"
  H="$(node tests/check_html.js 2>&1 | tail -1)"
  case "$B" in *"assertions passed"*) ok "backend — $B";; *) bad "backend — $B";; esac
  case "$H" in *"checks passed"*)     ok "frontend — $H";; *) bad "frontend — $H";; esac
  cd "$REPO" || exit 1
else
  warn "node tiada — ujian dilangkau"
fi

# ── 7. A clickable updater, created locally so Gatekeeper leaves it alone ──
step "7. Cipta fail klik untuk kemas kini"

UPD="$REPO/Kemas Kini i-Nventori.command"
cat > "$UPD" <<'UPDEOF'
#!/bin/bash
# Double-click to pull the latest and re-run the tests.
set -u
cd "$(dirname "${BASH_SOURCE[0]}")" || exit 1
printf '\033[1mKemas kini i-Nventori Ofis\033[0m\n\n'
git pull --rebase --autostash 2>&1 | tail -4
printf '\nTerkini: %s\n\n' "$(git log --oneline -1)"
if command -v node >/dev/null 2>&1; then
  cd i-Nventoriofis || exit 1
  node tests/test_backend.js 2>&1 | tail -1
  node tests/check_html.js  2>&1 | tail -1
fi
printf '\n'
read -r -p "Tekan Enter untuk tutup..." _ </dev/tty
UPDEOF
chmod +x "$UPD"
ok "Dicipta: Kemas Kini i-Nventori.command"

# ── 8. Warn about the credential trap we already hit once ───────────
step "8. Akaun git"

WHO="$(git config user.name 2>/dev/null || true)"
if [ -z "$WHO" ]; then
  warn "Identiti git belum ditetapkan. Jalankan:"
  echo "      git config --global user.name  \"Nama Anda\""
  echo "      git config --global user.email \"anda@contoh.com\""
else
  ok "Identiti git: $WHO"
fi
warn "Bila git minta log masuk untuk PUSH, guna ruxxzif89 atau imenhub-portal."
warn "Akaun photonicsmeeting TIADA akses tulis — ia pernah menyebabkan ralat 403."

# ── Done ────────────────────────────────────────────────────────────
echo
echo "─────────────────────────────────────────────"
bold "Siap."
echo
echo "  Projek     : $REPO/$APP"
echo "  Baca dulu  : $APP/CLAUDE.md"
echo "  Kemas kini : klik dua kali \"Kemas Kini i-Nventori.command\""
echo
echo "  Apps langsung : https://imenhub-portal.github.io/i-Nventoriofis/"
echo

if command -v opencode >/dev/null 2>&1; then
  read -r -p "Buka dalam opencode sekarang? [y/N] " a </dev/tty
  case "$a" in [yY]*) cd "$REPO/$APP" && opencode . ;; esac
else
  echo "  (opencode tiada dalam PATH — buka folder di atas secara manual)"
  echo
  read -r -p "Tekan Enter untuk tutup..." _ </dev/tty
fi
