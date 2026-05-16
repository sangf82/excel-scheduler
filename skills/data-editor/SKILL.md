---
name: data-editor
description: Handles all CRUD (Create, Read, Update, Delete) operations on the active scheduling Excel file (specifically INPUT 1-5 sections).
---

# Data Editor

## Operating Scope

This skill edits the **Active Project File** (e.g., `C:\Users\sluon\Documents\MedMate_Schedules\[Tên_Dự_Án]\schedule.xlsx`). 

It is responsible for modifying:
- INPUT 1: Lịch phòng khám
- INPUT 2: Lịch họp
- INPUT 3: Lịch trực cấp cứu
- INPUT 4: Danh sách bác sĩ
- INPUT 5: Danh sách bệnh nhân

## Workflow

1. **Verify Active File:** Ensure there is an active project file set. If not, prompt the user to specify or create one using `project-initializer`.
2. **Read Current Data:** Use the `@negokaz/excel-mcp-server` to read the current state of the relevant INPUT section.
3. **Confirm (Confirm.md):** Send the confirmation message using `confirm.md` before performing any write operations. Wait for user consent.
4. **Write Data:** Write the new data to the correct cells. **Never write to TÍNH TOÁN or OUTPUT sheets.**
5. **Verify (Read-back):** Re-read the modified cells to verify the data was saved correctly and report back to the user.
