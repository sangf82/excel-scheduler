# MedMate Scheduler

> Auto-schedule surgeries into `scheduling-template.xlsx` — medical domain only.

---

## 🇻🇳 Tiếng Việt

### Yêu cầu trước khi cài (ngoài Excel và Codex Desktop)

- **Python 3.8+** — cần để chạy `build_template.py` và tương tác workbook.
  - Windows: https://www.python.org/downloads/ (tích "Add to PATH")
  - macOS: `brew install python3`
- **Node.js 16+** — cần để `npx` chạy `excel-mcp-server`.
  - Windows/macOS: https://nodejs.org/ (chọn LTS)
- **openpyxl** — script tự động cài qua `pip` nếu chưa có.

### Plugin này làm gì

MedMate Scheduler là một Codex plugin tự động xếp lịch phẫu thuật vào file Excel `scheduling-template.xlsx`. Workbook có 4 sheet:

- **`HƯỚNG DẪN`** — hướng dẫn sử dụng template.
- **`INPUT`** — sheet nhập dữ liệu đầu vào:
  - **INPUT 1** — Lịch phòng khám (`Ngày`, `Thứ`, `Mã BS`).
  - **INPUT 2** — Lịch họp (`Ngày`, `Thứ`, `Nội dung`).
  - **INPUT 3** — Lịch trực cấp cứu (`Ngày`, `Thứ`, `Mã BS`).
  - **INPUT 4** — Danh sách bác sĩ (`Mã BS`, `Tên bác sĩ`, `Loại PT phụ trách`).
  - **INPUT 5** — Danh sách bệnh nhân cần xếp (`Mã BN`, `Tên bệnh nhân`, `Loại PT`).
- **`TÍNH TOÁN`** — công thức kiểm tra xung đột (không sửa tay).
- **`OUTPUT`** — ngày mổ, thứ, bác sĩ, bệnh nhân, loại PT, kết luận.

Plugin chỉ trả lời các yêu cầu về xếp lịch phẫu thuật trong workbook này. Các yêu cầu khác sẽ bị từ chối.

### Cài đặt (1 dòng lệnh)

**Windows (PowerShell):**
```powershell
powershell.exe -File .\scripts\install.ps1
```

> PowerShell 7+: `pwsh -File .\scripts\install.ps1`
> Thêm `-WhatIf` để xem trước: `powershell.exe -File .\scripts\install.ps1 -WhatIf`

**macOS (Terminal):**
```bash
chmod +x ./scripts/install.sh
./scripts/install.sh
```

> Thêm `--dry-run` để xem trước: `./scripts/install.sh --dry-run`

### 3 câu lệnh chat thử sau khi cài

1. `Thêm bệnh nhân BN036 tên Nguyễn Văn X loại Ghép gan`
2. `Cập nhật lịch trực ngày 2026-05-20 cho BS02`
3. `Chạy xếp lịch tuần này và xem OUTPUT CUỐI CÙNG`

### Cập nhật dữ liệu

Mở `scheduling-template.xlsx`, chuyển sang sheet **INPUT**:

- Sửa **INPUT 1** (columns A–C, rows 5-11) để thay đổi lịch phòng khám.
- Sửa **INPUT 2** (columns E–G, rows 5-11) để cập nhật các cuộc họp khóa lịch.
- Sửa **INPUT 3** (columns I–K, rows 5-10) để cập nhật lịch trực cấp cứu.
- Sửa **INPUT 4** (columns M–O, rows 5-10) để cập nhật danh sách bác sĩ và loại PT phụ trách (`Ghép gan`, `Cắt gan`, `Tái thùy`, `Cắt gan S`).
- Thêm bệnh nhân ở **INPUT 5** (columns Q–S, rows 5-40, tối đa 36 bệnh nhân).
- Sheet **TÍNH TOÁN** và **OUTPUT** sẽ tự cập nhật theo công thức. Giới hạn 11 ca/ngày được mã hóa cứng trong công thức (`M{row}<=11`).

Chi tiết schema xem `skills/scheduler/rules.md` và thuật toán xem `skills/scheduler/algorithm.md`.

### Gỡ cài

**Windows:**
```powershell
powershell.exe -File .\scripts\uninstall.ps1
```

Thêm `-RestoreLegacy` để khôi phục skill `excel` cũ (nếu đã backup).

**macOS:**
```bash
./scripts/uninstall.sh
```

Thêm `--restore-legacy` để khôi phục skill cũ.

---

## 🇬🇧 English

### Requirements (besides Excel and Codex Desktop)

- **Python 3.8+** — required to run `build_template.py` and interact with the workbook.
  - Windows: https://www.python.org/downloads/ (check "Add to PATH")
  - macOS: `brew install python3`
- **Node.js 16+** — required for `npx` to run the `excel-mcp-server`.
  - Windows/macOS: https://nodejs.org/ (choose LTS)
- **openpyxl** — the installer auto-installs it via `pip` if missing.

### What this plugin does

MedMate Scheduler is a Codex plugin that auto-schedules surgical cases into `scheduling-template.xlsx`. The workbook contains four sheets:

- **`HƯỚNG DẪN`** — usage instructions.
- **`INPUT`** — data entry sheet containing:
  - **INPUT 1** — Clinic schedule (date, weekday, doctor code).
  - **INPUT 2** — Meeting schedule (date, weekday, topic).
  - **INPUT 3** — Emergency duty rota (date, weekday, doctor code).
  - **INPUT 4** — Doctor roster (doctor code, full name, procedure type covered).
  - **INPUT 5** — Patients waiting for surgery (patient code, name, procedure type).
- **`TÍNH TOÁN`** — formula-driven conflict-check block (do not edit by hand).
- **`OUTPUT`** — final surgery list: date, weekday, doctor, patient, procedure, verdict.

The plugin only responds to surgery-scheduling requests against this workbook. Everything else is refused.

### Install (one command)

**Windows (PowerShell):**
```powershell
powershell.exe -File .\scripts\install.ps1
```

> PowerShell 7+: `pwsh -File .\scripts\install.ps1`
> Add `-WhatIf` for a dry run.

**macOS (Terminal):**
```bash
chmod +x ./scripts/install.sh
./scripts/install.sh
```

> Add `--dry-run` for a dry run.

### Three sample chat commands after install

1. `Add patient BN036 named Nguyễn Văn X, procedure Ghép gan`
2. `Update emergency duty 2026-05-20 to BS02`
3. `Run this week's schedule and show OUTPUT CUỐI CÙNG`

### How to edit data

Open `scheduling-template.xlsx` and switch to **INPUT**:

- Edit **INPUT 1** (columns A–C, rows 5-11) for clinic blocks.
- Edit **INPUT 2** (columns E–G, rows 5-11) for meetings that block doctors.
- Edit **INPUT 3** (columns I–K, rows 5-10) for emergency duty.
- Edit **INPUT 4** (columns M–O, rows 5-10) for the doctor roster and procedure coverage (`Ghép gan`, `Cắt gan`, `Tái thùy`, `Cắt gan S`).
- Add patients in **INPUT 5** (columns Q–S, rows 5-40, max 36 patients).
- The **TÍNH TOÁN** and **OUTPUT** sheets refresh automatically through Excel formulas. The daily cap of 11 surgeries is hardcoded in the formulas (`M{row}<=11`).

Full schema in `skills/scheduler/rules.md`; algorithm in `skills/scheduler/algorithm.md`.

### Uninstall

**Windows:**
```powershell
powershell.exe -File .\scripts\uninstall.ps1
```

Pass `-RestoreLegacy` to restore the previous `excel` skill (if backed up during install).

**macOS:**
```bash
./scripts/uninstall.sh
```

Pass `--restore-legacy` to restore the previous skill.
