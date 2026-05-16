---
workbook_path: C:\projects\MedMate\excel-scheduler\scheduling-template.xlsx
last_run: null
notes: |
  4-sheet model: INPUT (data entry), TÍNH TOÁN (compute block, rows 65-100),
  and OUTPUT (final schedule, rows 3-38). Daily cap of 11 surgeries is
  hardcoded in the TÍNH TOÁN column-N formula (M{row}<=11); to change the cap,
  edit those formulas directly.
---

# MedMate Scheduler memory

This file is loaded at the start of every session in
`C:\projects\MedMate\excel-scheduler\`. Append (do not overwrite) when the
user confirms a durable rule change.

## Notes

- Workbook of record: `scheduling-template.xlsx` (sheets: `HƯỚNG DẪN`, `INPUT`, `TÍNH TOÁN`, `OUTPUT`).
- Data entry sections on `INPUT` (all horizontal in one sheet):
  - INPUT 1 — Lịch phòng khám (columns A-C, rows 5-11).
  - INPUT 2 — Lịch họp (columns E-G, rows 5-11).
  - INPUT 3 — Lịch trực cấp cứu (columns I-K, rows 5-10).
  - INPUT 4 — Danh sách bác sĩ (columns M-O, rows 5-10).
  - INPUT 5 — Danh sách bệnh nhân / ca cần xếp (columns Q-S, rows 5-40, capacity 36).
- `TÍNH TOÁN` — row 64 headers, rows 65-100 compute (formula-only, read-only, references `INPUT!`).
- `OUTPUT` — row 2 headers, rows 3-38 output list, row 40 summary (formula-only, read-only, references `TÍNH TOÁN!`).
- The daily cap (11 surgeries) is hardcoded in the `TÍNH TOÁN` formulas in
  column N of rows 65-100. It is **not** stored as a parameter.
- Refuse any non-scheduling request with the bilingual refusal in
  `skills/scheduler/prompts/refuse.md`.
- Always confirm before writing; verify by re-reading the cell.
- Language: detect VI vs EN per message, respond in same language.

## Change log (append entries below)

<!-- Format:
- YYYY-MM-DD HH:MM | <intent> | <summary>
-->
