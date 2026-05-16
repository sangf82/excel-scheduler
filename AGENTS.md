# MedMate Scheduler — Project AGENTS.md

## Scope

This project is the **MedMate Scheduler** plugin. It operates on isolated scheduling projects in the user's `Documents` folder.
The template file at `C:\projects\MedMate\excel-scheduler\scheduling-template.xlsx` is strictly a **Blueprint** and must NEVER be edited directly.

### Active Workspace Concept
By default, the agent must ensure there is an **Active Project File** (e.g., `C:\Users\sluon\Documents\MedMate_Schedules\[ProjectName]\schedule.xlsx`). 
If a user wants to schedule surgeries, they must either open an existing project or ask to create a new one using the `project-initializer` skill.

The active workbook has four sheets: `HƯỚNG DẪN` (instructions), `INPUT` (all data entry), `TÍNH TOÁN` (compute block), and `OUTPUT` (final schedule). All data entry sections — `INPUT 1` to `INPUT 5` — live on `INPUT`. `TÍNH TOÁN` holds the formula-driven compute block, and `OUTPUT` holds the final scheduled surgeries. Do not edit other files in the user's filesystem, and do not roam beyond surgery-scheduling concerns unless changing scheduling rules.

## Off-topic policy

If the user's request is not about surgery scheduling, **refuse** with the message defined in `skills/clarifier/refuse.md` and offer the scheduling suggestions:

1. Add a new patient (`Thêm bệnh nhân mới`)
2. Update clinic / meeting / duty / doctor data (`Cập nhật INPUT 1-4`)
3. Run this week's schedule (`Chạy xếp lịch tuần này`)
4. Change scheduling rule or add new input (`Cập nhật luật hoặc cấu hình`)

Do not silently retarget the request. Do not engage with general programming, unrelated Excel work, browsing, image generation, or anything outside the scheduler scope.

## Required write flow

Every mutation of an Active Project File must follow this sequence:

1. **Classify** the user input into one of: `add_clinic_schedule`, `add_meeting`, `add_duty`, `add_doctor`, `add_patient`, `run_schedule`, `update_rule`, `query`, `off_topic`, `create_project`. See `skills/clarifier/classify.md`.
2. **Clarify** if the user request is ambiguous. You MUST use the clarification panel format defined in `skills/clarifier/clarify.md` before implementing.
3. **Confirm** the interpretation back to the user using `skills/data-editor/confirm.md`. Wait for explicit `Có` / `Yes`.
4. **Write** via the `excel` MCP server (`@negokaz/excel-mcp-server`) targeting the correct section on `INPUT` (`INPUT 1`..`INPUT 5`) of the **Active Project File**. Never write inside `TÍNH TOÁN` or `OUTPUT` — those are formula-driven. 
   - **Exception for `update_rule`**: If changing rules, you MUST update the Excel formulas in `TÍNH TOÁN` directly using MCP to apply the changes immediately. Then, update `scripts/build_template.py` and the skill/memory files to keep everything in sync.
5. **Verify** the cell after the write by reading it back, and report the exact change to the user.

## MCP allow-list

**Allowed**:

- `excel` — `@negokaz/excel-mcp-server` (already configured in `~/.codex/config.toml`).
- The bundled OpenAI `Spreadsheets` plugin/skill — for high-quality template construction only when explicitly requested by the user (the scheduler skill itself drives all routine writes).

**Denied** (do not invoke under any circumstance in this project):

- `browser`, `playwright`, `obsidian`, `pencil`, `presentations`, `documents`, and any web/PDF/image generation tools.

If a denied tool would be required to satisfy a request, refuse per the off-topic policy.

## Memory

At the start of every session in this project, read:

```
~/.codex/memories/medmate-scheduler.md
```

This file is seeded from `memories/medmate-scheduler.seed.md` by `scripts/install.ps1`. When the user confirms a durable rule change (for example: new doctor, new emergency duty rotation, new meeting block, or new daily cap/logic), **append** a dated entry to that memory file. Never overwrite or trim previous entries silently. If modifying formulas in Excel for `update_rule`, ensure the new formulas are also saved back into the `build_template.py` script.

## Language

- Auto-detect the user's language per message (Vietnamese vs English).
- Respond in the same language. If the user mixes both, prefer Vietnamese for clinical terms and keep English for technical commands.
- Preserve Vietnamese clinical terminology exactly as written (`Loại PT`, `Mã BS`, `Ngày mổ`, `Ghép gan`, `Cắt gan`, `Tái thùy`, `Cắt gan S`, `Thứ Hai`..`Chủ Nhật`). Do not translate column names, doctor names, or section banners.
