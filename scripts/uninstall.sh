#!/bin/bash
set -euo pipefail

DRY_RUN=false
RESTORE_LEGACY=false
while [[ $# -gt 0 ]]; do
  case "$1" in
    --dry-run|-n) DRY_RUN=true; shift ;;
    --restore-legacy) RESTORE_LEGACY=true; shift ;;
    *) echo "Không rõ tham số: $1"; exit 1 ;;
  esac
done

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"
CODEX_HOME="${HOME}/.codex"

BEGIN_MARKER='# BEGIN MEDMATE'
END_MARKER='# END MEDMATE'

log_plan() {
  if [[ "$DRY_RUN" == true ]]; then
    echo "[DRY-RUN] $1"
  else
    echo "$1"
  fi
}

if [[ ! -d "$CODEX_HOME" ]]; then
  echo "Lỗi: Không tìm thấy Codex home tại ${CODEX_HOME}." >&2
  exit 1
fi

echo "Trình gỡ cài đặt MedMate Scheduler (macOS)"

# 1. Remove plugin symlink
PLUGIN_LINK="${CODEX_HOME}/plugins/medmate-scheduler"
if [[ -L "$PLUGIN_LINK" ]] || [[ -e "$PLUGIN_LINK" ]]; then
  log_plan "Xóa liên kết plugin: ${PLUGIN_LINK}"
  if [[ "$DRY_RUN" != true ]]; then
    rm -rf "$PLUGIN_LINK"
  fi
else
  log_plan "Liên kết plugin không tồn tại tại ${PLUGIN_LINK}."
fi

# 2. Strip MEDMATE block + restore plugin states from config.toml
CONFIG_PATH="${CODEX_HOME}/config.toml"
if [[ -f "$CONFIG_PATH" && "$DRY_RUN" != true ]]; then
  SNAPSHOT_PATH="${CODEX_HOME}/.tmp/medmate-config-snapshot.json"

  python3 - "$CONFIG_PATH" "$SNAPSHOT_PATH" << 'PYEOF'
import sys, re, json, os
config_path = sys.argv[1]
snap_path = sys.argv[2]

with open(config_path, 'r', encoding='utf-8') as f:
    text = f.read()

# Remove MEDMATE block
text = re.sub(r'\r?\n?# BEGIN MEDMATE\s*.*?\s*# END MEDMATE\r?\n?', '\n', text, flags=re.DOTALL).strip()

# Restore plugin states from snapshot
if os.path.exists(snap_path):
    try:
        with open(snap_path, 'r', encoding='utf-8') as f:
            snapshot = json.load(f)
        for name, old_value in snapshot.items():
            pattern = r'(\[plugins\."' + re.escape(name) + r'"\]\s*\n\s*enabled\s*=\s*)(true|false)'
            if re.search(pattern, text):
                text = re.sub(pattern, r'\g<1>' + old_value, text)
        os.remove(snap_path)
    except Exception as e:
        print(f"Cảnh báo: Không khôi phục snapshot: {e}", file=sys.stderr)

with open(config_path, 'w', encoding='utf-8') as f:
    f.write(text)
    if not text.endswith('\n'):
        f.write('\n')
PYEOF
  log_plan "Xóa khối MEDMATE và khôi phục trạng thái plugin từ snapshot"
else
  log_plan "config.toml không tồn tại hoặc chạy dry-run."
fi

# 2b. Remove AGENTS.md block from ~/.codex/AGENTS.md
AGENTS_TARGET="${CODEX_HOME}/AGENTS.md"
AGENTS_BEGIN='# BEGIN MEDMATE AGENTS'
AGENTS_END='# END MEDMATE AGENTS'
if [[ -f "$AGENTS_TARGET" && "$DRY_RUN" != true ]]; then
  log_plan "Xóa MedMate rules khỏi ~/.codex/AGENTS.md"
  python3 - "$AGENTS_TARGET" "$AGENTS_BEGIN" "$AGENTS_END" << 'PYEOF'
import sys, re
path = sys.argv[1]
begin_marker = sys.argv[2]
end_marker = sys.argv[3]

with open(path, 'r', encoding='utf-8') as f:
    text = f.read()

pattern = r'\r?\n?' + re.escape(begin_marker) + r'\s*.*?\s*' + re.escape(end_marker) + r'\r?\n?'
text = re.sub(pattern, '\n', text, flags=re.DOTALL).strip()

with open(path, 'w', encoding='utf-8') as f:
    f.write(text)
    if not text.endswith('\n'):
        f.write('\n')
PYEOF
else
  log_plan "~/.codex/AGENTS.md không tồn tại hoặc chạy dry-run."
fi

# 3. Remove marketplace entry
MARKETPLACE_PATH="${HOME}/.agents/plugins/marketplace.json"
if [[ -f "$MARKETPLACE_PATH" ]]; then
  log_plan "Xóa medmate-scheduler khỏi marketplace.json"
  if [[ "$DRY_RUN" != true ]]; then
    python3 - "$MARKETPLACE_PATH" << 'PYEOF'
import json, sys
path = sys.argv[1]
try:
    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)
except (json.JSONDecodeError, FileNotFoundError):
    sys.exit(0)

if 'plugins' in data:
    before = len(data['plugins'])
    data['plugins'] = [p for p in data['plugins'] if p.get('name') != 'medmate-scheduler']
    after = len(data['plugins'])
    if after != before:
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
            f.write('\n')
PYEOF
  fi
else
  log_plan "marketplace.json không tồn tại."
fi

# 4. Remove memory file
MEMORY_TARGET="${CODEX_HOME}/memories/medmate-scheduler.md"
if [[ -f "$MEMORY_TARGET" ]]; then
  log_plan "Xóa file bộ nhớ: ${MEMORY_TARGET}"
  if [[ "$DRY_RUN" != true ]]; then
    rm -f "$MEMORY_TARGET"
  fi
else
  log_plan "File bộ nhớ không tồn tại."
fi

# 5. Optionally restore legacy excel skill
if [[ "$RESTORE_LEGACY" == true ]]; then
  TMP_DIR="${CODEX_HOME}/.tmp"
  if [[ -d "$TMP_DIR" ]]; then
    LATEST=$(find "$TMP_DIR" -maxdepth 1 -type d -name 'excel.bak.*' | sort -r | head -n 1)
    if [[ -n "$LATEST" ]]; then
      RESTORE_TARGET="${CODEX_HOME}/skills/excel"
      if [[ -e "$RESTORE_TARGET" ]]; then
        echo "Cảnh báo: skills/excel đã tồn tại; không ghi đè. Bản sao lưu mới nhất: ${LATEST}" >&2
      else
        log_plan "Khôi phục skill excel cũ từ ${LATEST}"
        if [[ "$DRY_RUN" != true ]]; then
          mkdir -p "$(dirname "$RESTORE_TARGET")"
          mv "$LATEST" "$RESTORE_TARGET"
        fi
      fi
    else
      log_plan "Không tìm thấy bản sao lưu excel.bak.* trong ${TMP_DIR}."
    fi
  else
    log_plan "${TMP_DIR} không tồn tại; không có gì để khôi phục."
  fi
fi

echo "Đã gỡ cài MedMate Scheduler hoàn tất."
