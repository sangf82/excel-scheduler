# MedMate Scheduler

> Auto-schedule surgeries into isolated Excel project files — medical domain only.

---

## 🇻🇳 Tiếng Việt

### Kiến trúc Hệ thống & Kỹ năng (Modular Skills)
MedMate Scheduler là một Codex plugin tự động xếp lịch phẫu thuật. Hệ thống sử dụng khái niệm **Active Workspace**:
- Template `scheduling-template.xlsx` chỉ đóng vai trò là **khuôn mẫu (Blueprint)**.
- Khi làm việc, plugin sẽ khởi tạo một **Dự án mới** tại thư mục `Documents/MedMate_Schedules/[Tên_Dự_Án]/schedule.xlsx`.

Để quản lý logic phức tạp, hệ thống được chia thành 4 kỹ năng độc lập:
1. **`project-initializer`**: Khởi tạo thư mục dự án mới, tự động build file Excel từ template và hỗ trợ import dữ liệu (như danh sách bác sĩ, lịch trực) từ file cũ.
2. **`data-editor`**: Phụ trách toàn bộ thao tác CRUD (Thêm/Sửa/Xóa) dữ liệu trên các phần INPUT 1-5 của file Excel đang active. Tự động xác minh trước và sau khi ghi qua MCP.
3. **`rule-manager`**: Cập nhật logic xếp lịch và công thức trên sheet `TÍNH TOÁN` qua MCP, đồng thời lưu trữ bộ luật bền vững vào memory toàn cục để tự động áp dụng cho các project sau.
4. **`clarifier`**: Chặn các yêu cầu ngoài luồng (off-topic) và gọi giao diện tương tác (UI form) để yêu cầu người dùng làm rõ nếu dữ liệu đầu vào không đầy đủ.

### Cấu trúc File Excel
- **`HƯỚNG DẪN`** — hướng dẫn sử dụng template.
- **`INPUT`** — sheet nhập dữ liệu đầu vào (Lịch phòng khám, Lịch họp, Lịch trực, Danh sách BS, Danh sách Bệnh nhân).
- **`TÍNH TOÁN`** — công thức kiểm tra xung đột (không sửa tay).
- **`OUTPUT`** — kết quả ngày mổ, bác sĩ, bệnh nhân cuối cùng.

### Yêu cầu trước khi cài (ngoài Excel và Codex Desktop)
- **Python 3.8+** — cần để chạy `build_template.py` và tương tác workbook.
  - Windows: https://www.python.org/downloads/ (tích "Add to PATH")
  - macOS: `brew install python3`
- **Node.js 16+** — cần để `npx` chạy `excel-mcp-server`.
  - Windows/macOS: https://nodejs.org/ (chọn LTS)

### Cài đặt (1 dòng lệnh)

**Windows (PowerShell):**
```powershell
irm https://raw.githubusercontent.com/sangf82/excel-scheduler/main/scripts/install.ps1 | iex
```

**macOS (Terminal):**
```bash
curl -fsSL https://raw.githubusercontent.com/sangf82/excel-scheduler/main/scripts/install.sh | bash
```

### Hướng dẫn sử dụng
1. Bắt đầu bằng cách tạo dự án mới: `Tạo lịch tuần 4 tháng 5`.
2. Hệ thống sẽ tạo `Documents/MedMate_Schedules/Tuan4_Thang5/schedule.xlsx` và đặt nó làm Active Project.
3. Chat các lệnh thêm dữ liệu:
   - `Thêm bệnh nhân BN036 tên Nguyễn Văn X loại Ghép gan`
   - `Cập nhật lịch trực ngày 2026-05-20 cho BS02`
4. Cập nhật luật (nếu cần): `Đổi luật ghép gan tối đa 1 ca một ngày`

### Gỡ cài đặt
**Windows:** `powershell.exe -File .\scripts\uninstall.ps1`
**macOS:** `./scripts/uninstall.sh`

---

## en English

### System Architecture & Modular Skills
MedMate Scheduler is a Codex plugin that auto-schedules surgical cases. The system utilizes an **Active Workspace** concept:
- The base `scheduling-template.xlsx` is strictly a **Blueprint**.
- Actual scheduling occurs in isolated project folders located at `Documents/MedMate_Schedules/[Project_Name]/schedule.xlsx`.

The plugin is broken down into 4 specialized skills:
1. **`project-initializer`**: Creates new scheduling projects from the template and handles historical data imports.
2. **`data-editor`**: Manages all CRUD operations on the active Excel file (INPUT 1-5). Performs verifications before and after writing.
3. **`rule-manager`**: Handles algorithmic updates directly to the `TÍNH TOÁN` sheet formulas and commits persistent rule changes to memory.
4. **`clarifier`**: Intercepts off-topic requests and renders native UI forms to clarify ambiguous or incomplete user requests.

### The Excel Workbook Structure
- **`HƯỚNG DẪN`** — usage instructions.
- **`INPUT`** — data entry sheet (Clinics, Meetings, Duties, Doctors, Patients).
- **`TÍNH TOÁN`** — formula-driven conflict-check block (do not edit by hand).
- **`OUTPUT`** — final generated schedule.

### Requirements (besides Excel and Codex Desktop)
- **Python 3.8+** — required for `build_template.py`.
- **Node.js 16+** — required for the `excel-mcp-server`.

### Installation

**Windows (PowerShell):**
```powershell
irm https://raw.githubusercontent.com/sangf82/excel-scheduler/main/scripts/install.ps1 | iex
```

**macOS (Terminal):**
```bash
curl -fsSL https://raw.githubusercontent.com/sangf82/excel-scheduler/main/scripts/install.sh | bash
```

### Usage Workflow
1. Initialize a new project: `Create a new schedule for Week 4 of May`.
2. The agent will set up `Documents/MedMate_Schedules/Week4_May/schedule.xlsx` as the active project.
3. Interact with the active schedule:
   - `Add patient BN036 named Nguyễn Văn X, procedure Ghép gan`
   - `Update emergency duty 2026-05-20 to BS02`
4. Update rules (if needed): `Change the rule so Ghép gan is limited to 1 per day`

### Uninstall
**Windows:** `powershell.exe -File .\scripts\uninstall.ps1`
**macOS:** `./scripts/uninstall.sh`
