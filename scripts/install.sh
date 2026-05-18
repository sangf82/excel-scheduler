#!/bin/bash
set -euo pipefail

DRY_RUN=false
while [[ $# -gt 0 ]]; do
  case "$1" in
    --dry-run|-n) DRY_RUN=true; shift ;;
    *) echo "Không rõ tham số: $1"; exit 1 ;;
  esac
done

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Handle remote execution (curl ... | bash) — SCRIPT_DIR will be the current working directory
if [[ "${SCRIPT_DIR}" == "$(pwd)" ]] || [[ ! -f "${SCRIPT_DIR}/install.sh" ]]; then
    echo "Detected remote execution. Cloning repository first..."
    REPO_URL='https://github.com/sangf82/excel-scheduler.git'
    CLONE_DIR="${TMPDIR:-/tmp}/medmate-scheduler-$$"

    if ! command -v git &>/dev/null; then
        echo "Error: Git is required for remote install." >&2
        echo "Please install git or clone the repo manually:" >&2
        echo "  git clone ${REPO_URL}" >&2
        exit 1
    fi

    if ! git clone "${REPO_URL}" "${CLONE_DIR}" 2>/dev/null; then
        echo ""
        echo "=== LỖI: Không thể clone repository ===" >&2
        echo "Repository có thể là PRIVATE và bạn chưa cấu hình xác thực git." >&2
        echo "" >&2
        echo "Các giải pháp:" >&2
        echo "  1. Chuyển repository sang PUBLIC trên GitHub (Settings → Danger Zone → Change visibility)." >&2
        echo "  2. Hoặc clone thủ công bằng tài khoản có quyền truy cập:" >&2
        echo "       git clone ${REPO_URL}" >&2
        echo "     Sau đó chạy script cài đặt từ thư mục đã clone:" >&2
        echo "       ./scripts/install.sh" >&2
        echo "" >&2
        echo "=== ERROR: Failed to clone repository ===" >&2
        echo "The repository may be PRIVATE and your git credentials are not configured." >&2
        echo "" >&2
        echo "Solutions:" >&2
        echo "  1. Make the repository PUBLIC on GitHub." >&2
        echo "  2. Or clone manually with an authorized account:" >&2
        echo "       git clone ${REPO_URL}" >&2
        echo "     Then run the local install script:" >&2
        echo "       ./scripts/install.sh" >&2
        echo ""
        exit 1
    fi

    echo "Repository cloned to ${CLONE_DIR}"
    exec "${CLONE_DIR}/scripts/install.sh" "$@"
fi

PROJECT_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"
CODEX_HOME="${HOME}/.codex"
AGENTS_HOME="${HOME}/.agents"

BEGIN_MARKER='# BEGIN MEDMATE'
END_MARKER='# END MEDMATE'

# Project trust block only (plugin states are updated in-place to avoid duplicate TOML tables)
MEDMATE_BLOCK=$(cat <<EOF
${BEGIN_MARKER}
[projects."${PROJECT_ROOT}"]
trust_level = "trusted"
${END_MARKER}
EOF
)

log_plan() {
  if [[ "$DRY_RUN" == true ]]; then
    echo "[DRY-RUN] $1"
  else
    echo "$1"
  fi
}

ensure_dir() {
  if [[ ! -d "$1" ]]; then
    log_plan "Tạo thư mục: $1"
    if [[ "$DRY_RUN" != true ]]; then
      mkdir -p "$1"
    fi
  fi
}

check_dependencies() {
  echo "--- Kiểm tra phụ thuộc / Dependency check ---"

  # 1. Python 3.8+
  PYTHON_EXE=""
  for candidate in python3 python; do
    if command -v "$candidate" &>/dev/null; then
      ver_str=$("$candidate" --version 2>&1 || true)
      if [[ "$ver_str" =~ Python[[:space:]]+([0-9]+)\.([0-9]+) ]]; then
        major="${BASH_REMATCH[1]}"
        minor="${BASH_REMATCH[2]}"
        if [[ "$major" -gt 3 ]] || [[ "$major" -eq 3 && "$minor" -ge 8 ]]; then
          PYTHON_EXE="$candidate"
          echo "  Python OK: $ver_str ($PYTHON_EXE)"
          break
        fi
      fi
    fi
  done

  if [[ -z "$PYTHON_EXE" ]]; then
    echo ""
    echo "=== LỖI: Không tìm thấy Python 3.8+ ==="
    echo "Python chưa được cài đặt hoặc phiên bản quá cũ."
    echo ""
    echo "Hướng dẫn cài đặt macOS:"
    echo "  brew install python3"
    echo "Hoặc tải từ: https://www.python.org/downloads/macos/"
    echo ""
    echo "=== ERROR: Python 3.8+ not found ==="
    echo "Please install Python:"
    echo "  brew install python3"
    echo "  or https://www.python.org/downloads/macos/"
    exit 1
  fi

  # 2. openpyxl (auto-install)
  if ! "$PYTHON_EXE" -c "import openpyxl" 2>/dev/null; then
    echo "  openpyxl chưa có. Đang cài đặt..."
    if ! "$PYTHON_EXE" -m pip install --user openpyxl 2>/dev/null; then
      "$PYTHON_EXE" -m pip install openpyxl
    fi
    if ! "$PYTHON_EXE" -c "import openpyxl" 2>/dev/null; then
      echo "=== LỖI: Cài đặt openpyxl thất bại ==="
      echo "Chạy tay: pip3 install openpyxl"
      echo "=== ERROR: openpyxl installation failed ==="
      echo "Try manually: pip3 install openpyxl"
      exit 1
    fi
    echo "  openpyxl đã cài đặt xong"
  else
    op_ver=$("$PYTHON_EXE" -c "import openpyxl; print(openpyxl.__version__)" 2>/dev/null || true)
    echo "  openpyxl OK: $op_ver"
  fi

  # 3. Node.js 16+
  NODE_EXE=""
  if command -v node &>/dev/null; then
    ver_str=$(node --version 2>/dev/null || true)
    if [[ "$ver_str" =~ v?([0-9]+) ]]; then
      major="${BASH_REMATCH[1]}"
      if [[ "$major" -ge 16 ]]; then
        NODE_EXE="$(command -v node)"
        echo "  Node.js OK: $ver_str ($NODE_EXE)"
      fi
    fi
  fi

  if [[ -z "$NODE_EXE" ]]; then
    echo ""
    echo "=== LỖI: Không tìm thấy Node.js 16+ ==="
    echo "Node.js cần thiết để chạy excel-mcp-server qua npx."
    echo ""
    echo "Hướng dẫn cài đặt macOS:"
    echo "  brew install node"
    echo "Hoặc tải từ: https://nodejs.org/ (chọn LTS)"
    echo ""
    echo "=== ERROR: Node.js 16+ not found ==="
    echo "Node.js is required for the excel MCP server (npx)."
    echo "  brew install node"
    echo "  or https://nodejs.org/ (choose LTS)"
    exit 1
  fi

  # 4. npx
  if command -v npx &>/dev/null; then
    npx_ver=$(npx --version 2>/dev/null || true)
    echo "  npx OK: $npx_ver"
  else
    echo "=== CẢNH BÁO: Không tìm thấy npx ==="
    echo "npx thường đi kèm với Node.js. Hãy kiểm tra lại cài đặt Node.js."
    echo "=== WARNING: npx not found ==="
    echo "npx usually comes with Node.js. Please verify your Node.js installation."
  fi

  echo "--- Phụ thuộc đã sẵn sàng / Dependencies OK ---"
  echo ""

  # Export for later steps
  export PYTHON_EXE
}

if [[ ! -d "$CODEX_HOME" ]]; then
  echo "Lỗi: Không tìm thấy Codex home tại ${CODEX_HOME}. Vui lòng cài Codex Desktop trước." >&2
  exit 1
fi

echo "Trình cài đặt MedMate Scheduler (macOS)"
echo "Thư mục dự án: ${PROJECT_ROOT}"
echo "Codex home:    ${CODEX_HOME}"
if [[ "$DRY_RUN" == true ]]; then
  echo "Đang chạy ở chế độ --dry-run. Không thay đổi gì trên đĩa."
fi

# --- Dependency check ---
if [[ "$DRY_RUN" != true ]]; then
  check_dependencies
else
  log_plan "Bỏ qua kiểm tra phụ thuộc (dry-run)."
fi

# 1. Plugin symlink
PLUGIN_LINK="${CODEX_HOME}/plugins/medmate-scheduler"
ensure_dir "$(dirname "$PLUGIN_LINK")"

if [[ -L "$PLUGIN_LINK" ]] || [[ -e "$PLUGIN_LINK" ]]; then
  log_plan "Liên kết plugin đã tồn tại tại ${PLUGIN_LINK}; giữ nguyên."
else
  log_plan "Tạo symlink: ${PLUGIN_LINK} -> ${PROJECT_ROOT}"
  if [[ "$DRY_RUN" != true ]]; then
    ln -s "${PROJECT_ROOT}" "${PLUGIN_LINK}"
  fi
fi

# 2. Patch ~/.codex/config.toml
CONFIG_PATH="${CODEX_HOME}/config.toml"
if [[ ! -f "$CONFIG_PATH" ]]; then
  log_plan "Tạo mới config.toml tại ${CONFIG_PATH}"
  if [[ "$DRY_RUN" != true ]]; then
    touch "$CONFIG_PATH"
  fi
fi

if [[ -f "$CONFIG_PATH" && "$DRY_RUN" != true ]]; then
  SNAPSHOT_PATH="${CODEX_HOME}/.tmp/medmate-config-snapshot.json"

  python3 - "$CONFIG_PATH" "$SNAPSHOT_PATH" "$PROJECT_ROOT" << 'PYEOF'
import sys, re, json, os
config_path = sys.argv[1]
snap_path = sys.argv[2]

with open(config_path, 'r', encoding='utf-8') as f:
    text = f.read()

# Snapshot plugin states (only once)
plugin_names = [
    'browser@openai-bundled',
    'presentations@openai-primary-runtime',
    'documents@openai-primary-runtime',
    'spreadsheets@openai-primary-runtime'
]
snapshot = {}
for name in plugin_names:
    m = re.search(r'\[plugins\."' + re.escape(name) + r'"\]\s*\n\s*enabled\s*=\s*(true|false)', text)
    if m:
        snapshot[name] = m.group(1)

if snapshot and not os.path.exists(snap_path):
    os.makedirs(os.path.dirname(snap_path), exist_ok=True)
    with open(snap_path, 'w', encoding='utf-8') as f:
        json.dump(snapshot, f)
        f.write('\n')

# Remove old MEDMATE block
text = re.sub(r'# BEGIN MEDMATE\s*.*?\s*# END MEDMATE', '', text, flags=re.DOTALL).strip()

# Update existing plugin entries in-place; collect missing ones
replacements = {
    'browser@openai-bundled': 'false',
    'presentations@openai-primary-runtime': 'false',
    'documents@openai-primary-runtime': 'false',
    'spreadsheets@openai-primary-runtime': 'true',
}
missing = []
for name, value in replacements.items():
    pattern = r'(\[plugins\."' + re.escape(name) + r'"\]\s*\n\s*enabled\s*=\s*)(true|false)'
    if re.search(pattern, text):
        text = re.sub(pattern, r'\g<1>' + value, text)
    else:
        missing.append(f'[plugins."{name}"]\nenabled = {value}\n')

# Append missing plugins + project trust block at the end
append = ''
if missing:
    append += '\n'.join(missing) + '\n'
append += '# BEGIN MEDMATE\n[projects."' + sys.argv[3] + '"]\ntrust_level = "trusted"\n# END MEDMATE\n'

if not text.endswith('\n'):
    text += '\n'
text += '\n' + append

with open(config_path, 'w', encoding='utf-8') as f:
    f.write(text)
PYEOF
  log_plan "Cập nhật config.toml (snapshot plugin, cập nhật entries, thêm project trust)"
fi

# 2b. Install AGENTS.md into ~/.codex/AGENTS.md (global rule enforcement)
AGENTS_SOURCE="${PROJECT_ROOT}/AGENTS.md"
AGENTS_TARGET="${CODEX_HOME}/AGENTS.md"
AGENTS_BEGIN='# BEGIN MEDMATE AGENTS'
AGENTS_END='# END MEDMATE AGENTS'

if [[ -f "$AGENTS_SOURCE" && "$DRY_RUN" != true ]]; then
  log_plan "Cập nhật ~/.codex/AGENTS.md với MedMate Scheduler rules"
  "$PYTHON_EXE" - "$AGENTS_SOURCE" "$AGENTS_TARGET" "$AGENTS_BEGIN" "$AGENTS_END" << 'PYEOF'
import sys, re
source_path = sys.argv[1]
target_path = sys.argv[2]
begin_marker = sys.argv[3]
end_marker = sys.argv[4]

with open(source_path, 'r', encoding='utf-8') as f:
    source_text = f.read()

existing = ''
try:
    with open(target_path, 'r', encoding='utf-8') as f:
        existing = f.read()
except FileNotFoundError:
    pass

# Remove old block
pattern = re.escape(begin_marker) + r'\s*.*?\s*' + re.escape(end_marker)
existing = re.sub(pattern, '', existing, flags=re.DOTALL).strip()

block = f'\n{begin_marker}\n{source_text}\n{end_marker}\n'
sep = '' if existing.endswith('\n') or not existing else '\n'
new_text = existing + sep + block

with open(target_path, 'w', encoding='utf-8') as f:
    f.write(new_text)
PYEOF
else
  log_plan "Không tìm thấy AGENTS.md trong thư mục dự án; bỏ qua."
fi

# 3. Patch ~/.agents/plugins/marketplace.json
MARKETPLACE_PATH="${AGENTS_HOME}/plugins/marketplace.json"
ensure_dir "$(dirname "$MARKETPLACE_PATH")"

if [[ "$DRY_RUN" != true ]]; then
  if [[ ! -f "$MARKETPLACE_PATH" ]]; then
    echo '{"name":"medmate","interface":{"displayName":"MedMate Marketplace"},"plugins":[]}' > "$MARKETPLACE_PATH"
  fi

  # Use Python for reliable JSON manipulation (jq may not be installed on macOS)
  python3 - "$MARKETPLACE_PATH" "$PROJECT_ROOT" << 'PYEOF'
import json, sys
path = sys.argv[1]
project_root = sys.argv[2]
with open(path, 'r', encoding='utf-8') as f:
    data = json.load(f)
if 'plugins' not in data:
    data['plugins'] = []
plugins = [p for p in data['plugins'] if p.get('name') != 'medmate-scheduler']
plugins.append({
    'name': 'medmate-scheduler',
    'source': {
        'source': 'local',
        'path': project_root
    },
    'category': 'Productivity',
    'policy': {
        'installation': 'AVAILABLE',
        'authentication': 'ON_INSTALL'
    }
})
data['plugins'] = plugins
with open(path, 'w', encoding='utf-8') as f:
    json.dump(data, f, indent=2, ensure_ascii=False)
    f.write('\n')
PYEOF
else
  log_plan "Ghi marketplace.json với entry medmate-scheduler"
fi

# 3b. Setup ~/.codex/mcp.json (inject excel MCP server)
MCP_JSON_PATH="${CODEX_HOME}/mcp.json"
if [[ "$DRY_RUN" != true ]]; then
  if [[ ! -f "$MCP_JSON_PATH" ]]; then
    echo '{"mcpServers":{}}' > "$MCP_JSON_PATH"
  fi
  python3 - "$MCP_JSON_PATH" << 'PYEOF'
import json, sys
path = sys.argv[1]
try:
    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)
except Exception:
    data = {}
if 'mcpServers' not in data:
    data['mcpServers'] = {}
data['mcpServers']['excel'] = {
    'command': 'npx',
    'args': ['-y', '@negokaz/excel-mcp-server']
}
with open(path, 'w', encoding='utf-8') as f:
    json.dump(data, f, indent=2, ensure_ascii=False)
    f.write('\n')
PYEOF
  log_plan "Cập nhật mcp.json với excel MCP server"
else
  log_plan "Cập nhật mcp.json (dry-run)"
fi

# 4. Seed memory file
MEMORY_SOURCE="${PROJECT_ROOT}/memories/medmate-scheduler.seed.md"
MEMORY_TARGET="${CODEX_HOME}/memories/medmate-scheduler.md"
ensure_dir "$(dirname "$MEMORY_TARGET")"

if [[ -f "$MEMORY_TARGET" ]]; then
  log_plan "File bộ nhớ đã có tại ${MEMORY_TARGET}; giữ nguyên."
else
  log_plan "Sao chép bộ nhớ mẫu sang ${MEMORY_TARGET}"
  if [[ "$DRY_RUN" != true ]]; then
    cp "$MEMORY_SOURCE" "$MEMORY_TARGET"
  fi
fi

# 5. Backup legacy ~/.codex/skills/excel
LEGACY_SKILL="${CODEX_HOME}/skills/excel"
if [[ -e "$LEGACY_SKILL" ]]; then
  TIMESTAMP=$(date +%Y%m%d-%H%M%S)
  BACKUP_ROOT="${CODEX_HOME}/.tmp/excel.bak.${TIMESTAMP}"
  log_plan "Sao lưu skill cũ: ${LEGACY_SKILL} -> ${BACKUP_ROOT}"
  if [[ "$DRY_RUN" != true ]]; then
    ensure_dir "$(dirname "$BACKUP_ROOT")"
    mv "$LEGACY_SKILL" "$BACKUP_ROOT"
  fi
else
  log_plan "Không có ~/.codex/skills/excel cũ để sao lưu."
fi

# 6. Move project-local excel/ -> legacy/excel/
PROJECT_EXCEL="${PROJECT_ROOT}/excel"
LEGACY_DIR="${PROJECT_ROOT}/legacy"
if [[ -e "$PROJECT_EXCEL" ]]; then
  ensure_dir "$LEGACY_DIR"
  TARGET="${LEGACY_DIR}/excel"
  if [[ -e "$TARGET" ]]; then
    log_plan "legacy/excel đã tồn tại; để nguyên project excel/."
  else
    log_plan "Di chuyển ${PROJECT_EXCEL} -> ${TARGET}"
    if [[ "$DRY_RUN" != true ]]; then
      mv "$PROJECT_EXCEL" "$TARGET"
    fi
  fi
else
  log_plan "Không có thư mục excel/ cục bộ để di chuyển."
fi

# 7. Build the workbook
BUILD_SCRIPT="${PROJECT_ROOT}/scripts/build_template.py"
log_plan "Chạy build_template.py để tạo lại scheduling-template.xlsx"
if [[ "$DRY_RUN" != true ]]; then
  # PYTHON_EXE already discovered & validated in check_dependencies
  "$PYTHON_EXE" "$BUILD_SCRIPT"
  if [[ $? -ne 0 ]]; then
    echo "Lỗi: build_template.py thất bại." >&2
    exit 1
  fi
fi

echo ""
echo "Đã cài đặt MedMate Scheduler. Hãy khởi động lại Codex Desktop, mở scheduling-template.xlsx (sheet INPUT), rồi gõ: Thêm bệnh nhân BN036 tên Nguyễn Văn X loại Ghép gan"
