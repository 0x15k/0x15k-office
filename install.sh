#!/usr/bin/env bash
set -euo pipefail

# Turnkey macOS Installer for 0x15k Office AI for Microsoft Office
# Installs runtime into ~/.local/share/office-ai (avoiding ~/Downloads TCC restrictions),
# sideloads Office WEF manifests, configures a LaunchAgent daemon, and links `office-ai`.

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
INSTALL_DIR="$HOME/.local/share/office-ai"
MANIFEST_SRC="$PROJECT_DIR/manifest.xml"
PLIST_LABEL="com.antigravity.office-ai"
PLIST_PATH="$HOME/Library/LaunchAgents/${PLIST_LABEL}.plist"
LOG_PATH="$HOME/Library/Logs/office-ai.log"
CLI_LINK="$HOME/.local/bin/office-ai"

SIDELOAD_ONLY=false
if [[ "${1:-}" == "--sideload-only" ]]; then
  SIDELOAD_ONLY=true
fi

log()  { printf "[office-ai] %s\n" "$*"; }
warn() { printf "[office-ai] WARN: %s\n" "$*" >&2; }
fail() { printf "[office-ai] ERROR: %s\n" "$*" >&2; exit 1; }

resolve_node() {
  if command -v node >/dev/null 2>&1; then
    command -v node
    return
  fi
  for candidate in /opt/homebrew/bin/node /usr/local/bin/node; do
    if [[ -x "$candidate" ]]; then
      echo "$candidate"
      return
    fi
  done
  fail "Node.js is required (v18+). Install it via: brew install node"
}

install_dependencies() {
  if [[ ! -d "$PROJECT_DIR/node_modules" ]]; then
    log "Installing Node.js dependencies..."
    (cd "$PROJECT_DIR" && npm install --no-fund --no-audit)
  fi
}

setup_certificates() {
  if [[ -f "$PROJECT_DIR/cert.pem" && -f "$PROJECT_DIR/key.pem" ]]; then
    log "HTTPS certificates already present."
    return
  fi

  if command -v mkcert >/dev/null 2>&1; then
    log "Generating trusted local HTTPS certificates via mkcert..."
    mkcert -install >/dev/null 2>&1 || true
    (cd "$PROJECT_DIR" && mkcert -key-file key.pem -cert-file cert.pem localhost 127.0.0.1 ::1)
  else
    warn "mkcert not found; server.js will auto-generate self-signed certificates on startup."
  fi
}

sync_runtime_files() {
  if [[ "$PROJECT_DIR" == "$INSTALL_DIR" ]]; then
    return
  fi
  mkdir -p "$INSTALL_DIR"
  rsync -a --delete --exclude='.git' --exclude='.source-dir' --exclude='skills/learned-*.md' "$PROJECT_DIR/" "$INSTALL_DIR/"
  printf '%s\n' "$PROJECT_DIR" > "$INSTALL_DIR/.source-dir"
  log "Runtime synced to $INSTALL_DIR"
}

sideload_office_apps() {
  [[ -f "$MANIFEST_SRC" ]] || fail "manifest.xml not found at $MANIFEST_SRC"

  local apps=(
    "com.microsoft.Word:Word"
    "com.microsoft.Excel:Excel"
    "com.microsoft.Powerpoint:PowerPoint"
  )

  for entry in "${apps[@]}"; do
    local bundle_id="${entry%%:*}"
    local app_name="${entry##*:}"
    local docs_dir="$HOME/Library/Containers/${bundle_id}/Data/Documents"
    local wef_dir="${docs_dir}/wef"

    if [[ ! -d "$docs_dir" ]]; then
      log "${app_name}: container not initialized yet (open ${app_name} once to enable sideloading)."
      continue
    fi

    if mkdir -p "$wef_dir" 2>/dev/null && cp -f "$MANIFEST_SRC" "$wef_dir/manifest.xml" 2>/dev/null; then
      log "${app_name}: manifest installed to $wef_dir/manifest.xml"
    else
      log "${app_name}: run 'office-ai install' in Terminal.app if the add-in is not yet visible in the Ribbon."
    fi

    defaults write "$bundle_id" OfficeWebAddinDeveloperExtras -bool true >/dev/null 2>&1 || true
  done
}

install_global_cli() {
  local cli_target="$INSTALL_DIR/bin/office-cli.js"
  mkdir -p "$HOME/.local/bin"
  chmod +x "$cli_target"
  ln -sf "$cli_target" "$CLI_LINK"
  log "Global CLI linked at $CLI_LINK -> $cli_target"
}

install_launch_agent() {
  local node_bin="$1"
  mkdir -p "$HOME/Library/LaunchAgents" "$HOME/Library/Logs"
  : > "$LOG_PATH"

  cat > "$PLIST_PATH" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>${PLIST_LABEL}</string>
  <key>ProgramArguments</key>
  <array>
    <string>${node_bin}</string>
    <string>${INSTALL_DIR}/server.js</string>
  </array>
  <key>WorkingDirectory</key>
  <string>${INSTALL_DIR}</string>
  <key>EnvironmentVariables</key>
  <dict>
    <key>HOME</key>
    <string>${HOME}</string>
    <key>PATH</key>
    <string>${HOME}/.local/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin</string>
  </dict>
  <key>RunAtLoad</key>
  <true/>
  <key>KeepAlive</key>
  <true/>
  <key>StandardOutPath</key>
  <string>${LOG_PATH}</string>
  <key>StandardErrorPath</key>
  <string>${LOG_PATH}</string>
</dict>
</plist>
EOF

  launchctl unload "$PLIST_PATH" >/dev/null 2>&1 || true
  launchctl load -w "$PLIST_PATH"
  log "LaunchAgent registered and started (${PLIST_LABEL})."
}

wait_for_server() {
  local tries=16
  while (( tries > 0 )); do
    if curl -k -s --max-time 2 https://localhost:3000/api/status >/dev/null 2>&1; then
      log "Server is live at https://localhost:3000"
      return 0
    fi
    sleep 0.5
    ((tries--))
  done
  warn "Server did not respond on https://localhost:3000 within 8s. Check logs: $LOG_PATH"
  return 1
}

main() {
  if [[ "$SIDELOAD_ONLY" == true ]]; then
    sideload_office_apps
    exit 0
  fi

  local node_bin
  node_bin="$(resolve_node)"
  log "Using Node.js: $node_bin"

  install_dependencies
  setup_certificates
  sync_runtime_files
  sideload_office_apps
  install_global_cli
  install_launch_agent "$node_bin"
  wait_for_server

  echo ""
  echo "Installation complete!"
  echo "  Runtime Dir        : $INSTALL_DIR"
  echo "  Background Service : $PLIST_LABEL (auto-starts at login)"
  echo "  CLI Command        : office-ai (at $CLI_LINK)"
  echo "  Try                : office-ai status"
  if [[ ":$PATH:" != *":$HOME/.local/bin:"* ]]; then
    echo "  Note               : Add ~/.local/bin to your PATH if 'office-ai' is not found:"
    echo "                       export PATH=\"\$HOME/.local/bin:\$PATH\""
  fi
}

main "$@"
