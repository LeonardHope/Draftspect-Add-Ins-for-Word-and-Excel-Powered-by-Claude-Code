#!/usr/bin/env bash
# Sideload Draftspect into Windows-side Word + Excel from WSL.
#
# Mirrors the win32 branch of app/sideload.mjs, but driven from WSL so the
# daemon can keep running inside WSL while Office runs on the Windows host.
# The Electron tray + auto-sideloader refuse to run on linux, hence this
# script. Idempotent — safe to re-run; pass --uninstall to remove.

set -euo pipefail

REG_KEY='HKCU\Software\Microsoft\Office\16.0\WEF\Developer'
HOSTS=(word excel)

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
MANIFESTS_DIR="$PROJECT_ROOT/manifests"

MODE="install"
if [[ "${1:-}" == "--uninstall" ]]; then
  MODE="uninstall"
fi

if ! grep -qiE 'microsoft|wsl' /proc/version 2>/dev/null; then
  echo "error: this script is for WSL only (detected: $(uname -sr))" >&2
  exit 1
fi
for cmd in cmd.exe reg.exe wslpath; do
  command -v "$cmd" >/dev/null 2>&1 || { echo "error: '$cmd' not in PATH" >&2; exit 1; }
done

# cmd.exe emits CRLF and a UNC-path warning to stderr when launched from a
# WSL cwd — drop both.
WIN_USER="$(cmd.exe /c "echo %USERNAME%" 2>/dev/null | tr -d '\r\n')"
WIN_APPDATA_WIN="$(cmd.exe /c "echo %APPDATA%" 2>/dev/null | tr -d '\r\n')"

if [[ -z "$WIN_USER" || -z "$WIN_APPDATA_WIN" ]]; then
  echo "error: couldn't read Windows USERNAME / APPDATA via cmd.exe" >&2
  exit 1
fi

WIN_APPDATA_WSL="$(wslpath -u "$WIN_APPDATA_WIN")"
TARGET_DIR_WSL="$WIN_APPDATA_WSL/Draftspect/manifests"
TARGET_DIR_WIN="$WIN_APPDATA_WIN\\Draftspect\\manifests"

echo "Windows user:     $WIN_USER"
echo "Target (Windows): $TARGET_DIR_WIN"
echo

if [[ "$MODE" == "uninstall" ]]; then
  for host in "${HOSTS[@]}"; do
    dst_wsl="$TARGET_DIR_WSL/${host}.xml"
    dst_win="$TARGET_DIR_WIN\\${host}.xml"
    reg.exe delete "$REG_KEY" /v "$dst_win" /f >/dev/null 2>&1 || true
    rm -f "$dst_wsl"
    echo "  removed: ${host}.xml"
  done
  echo
  echo "Uninstall complete. Restart Word/Excel for the change to take effect."
  exit 0
fi

if command -v npm >/dev/null 2>&1 && [[ ! -d "$PROJECT_ROOT/node_modules" ]]; then
  echo "Installing npm dependencies…"
  ( cd "$PROJECT_ROOT" && npm install )
  echo
fi

mkdir -p "$TARGET_DIR_WSL"

for host in "${HOSTS[@]}"; do
  src="$MANIFESTS_DIR/${host}.xml"
  dst_wsl="$TARGET_DIR_WSL/${host}.xml"
  dst_win="$TARGET_DIR_WIN\\${host}.xml"

  if [[ ! -f "$src" ]]; then
    echo "warn: $src missing — skipping" >&2
    continue
  fi

  cp "$src" "$dst_wsl"
  reg.exe add "$REG_KEY" /v "$dst_win" /t REG_SZ /d "$dst_win" /f >/dev/null
  echo "  installed: ${host}.xml  ->  $dst_win"
done

cat <<'EOF'

Sideload complete.

Next:
  1. npm run dev                     # start the daemon (HTTP :47834, WS :47833)
  2. (Re)launch Word / Excel on Windows
  3. Insert -> Office Add-ins -> SHARED FOLDER  (or My Add-ins -> DEVELOPER ADD-INS)
     -> Draftspect for Word / Excel

To remove later: scripts/wsl-sideload.sh --uninstall
EOF
