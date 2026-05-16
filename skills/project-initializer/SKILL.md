---
name: project-initializer
description: Creates new scheduling projects in the user's Documents folder. Builds the Excel template at the target location and sets the active project context. Also handles importing historical data.
---

# Project Initializer

## Operating Scope

This skill is responsible for the lifecycle of creating a new surgical scheduling project.
Instead of operating on a single static template, this skill creates standalone projects.

**Default Directory:** `C:\Users\sluon\Documents\MedMate_Schedules\[Tên_Dự_Án]\`

## Workflow

1. **Ask for Project Name:** When the user wants to create a new schedule (e.g., "Tạo lịch tuần sau"), ask for a project/folder name (e.g., `Tuan3_Thang5`).
2. **Build Template:** Use the local script to build the fresh Excel file at the new location:
   `python c:\projects\MedMate\excel-scheduler\scripts\build_template.py --output "C:\Users\sluon\Documents\MedMate_Schedules\[Tên_Dự_Án]\schedule.xlsx"`
3. **Ask for Import:** Check if there are existing projects in `Documents\MedMate_Schedules\`. Ask the user: *"Bạn có muốn import cấu hình (Danh sách bác sĩ, Lịch họp, Lịch trực) từ file cũ sang không?"*. If yes, use the Excel MCP to read from the old file and write to the new file.
4. **Set Active Project:** Inform the user that the active project is now set to this new file, and all subsequent edits will apply to it.
