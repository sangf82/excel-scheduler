---
name: scheduler
description: Auto-schedule surgeries in scheduling-template.xlsx. Classifies user input, confirms target, writes to INPUT 1-5 sections on the INPUT sheet, and surfaces the formula-driven OUTPUT list. Medical-domain scheduling only — refuses everything else.
---

# Scheduler

## Operating Scope

This skill operates **only** on the workbook:

```
C:\projects\MedMate\excel-scheduler\scheduling-template.xlsx
```

The workbook has four sheets in this exact order: `HƯỚNG DẪN`, `INPUT`, `TÍNH TOÁN`, `OUTPUT`.

| Sheet        | Section / Range | Purpose                                                                     |
| ------------ | --------------- | --------------------------------------------------------------------------- |
| `INPUT`      | INPUT 1 (rows 4-12)   | Clinic schedule (`Ngày`, `Thứ`, `Mã BS`).                                   |
| `INPUT`      | INPUT 2 (rows 4-12)   | Meeting blocks (`Ngày`, `Thứ`, `Nội dung`).                               |
| `INPUT`      | INPUT 3 (rows 14-21)  | Emergency-duty rota (`Ngày`, `Thứ`, `Mã BS`).                               |
| `INPUT`      | INPUT 4 (rows 14-21)  | Doctor roster (`Mã BS`, `Tên bác sĩ`, `Loại PT phụ trách`).                |
| `INPUT`      | INPUT 5 (rows 24-61)  | Patient list (`Mã BN`, `Tên bệnh nhân`, `Loại PT`).                        |
| `TÍNH TOÁN`  | rows 63-100           | Compute block (formulas referencing `INPUT!`, **do not edit by hand**).    |
| `OUTPUT`     | rows 103-140          | Final scheduled surgeries (formulas referencing `TÍNH TOÁN!`, **read-only**). |
| `OUTPUT`     | Summary (row 142)     | Totals (counts of `Đủ điều kiện` and `Loại`, daily cap).                  |

The skill must:

- Read existing data from INPUT 1-5 before writing.
- Write only into `INPUT` sheet cells (INPUT 1-5 sections). Never write into `TÍNH TOÁN` or `OUTPUT`; those refresh automatically.
- Refuse anything that is not surgical scheduling for this workbook.

Use the `excel` MCP server (`@negokaz/excel-mcp-server`) for all reads and writes. Do not invoke `browser`, `playwright`, `obsidian`, `pencil`, `presentations`, or `documents`.

## Workflow

```
user message
   |
   v
[1] classify  --> off_topic ---------------> refuse (prompts/refuse.md), stop
   |
   v (add_clinic_schedule | add_meeting | add_duty | add_doctor |
   |  add_patient | run_schedule | query)
   |
[2] ambiguous?  --yes--> ask 1 short clarification, stop
   |
   v (no)
[3] confirm interpretation (prompts/confirm.md), wait for Có/Yes
   |
   v
[4] write via excel MCP into the correct INPUT section, OR (for run_schedule)
   |  verify that KHU VỰC TÍNH TOÁN + OUTPUT CUỐI CÙNG have refreshed.
   v
[5] verify by re-reading the cell(s) and report the exact change.
```

## Section contract

### INPUT 1 — Lịch phòng khám (rows 6-12)

| Column | Header | Type   | Notes                                       |
| ------ | ------ | ------ | ------------------------------------------- |
| A      | Ngày   | date   | `dd/mm/yyyy`. Row 5 holds the header.       |
| B      | Thứ    | text   | Formula `=TEXT(A{row},"[$-vi-VN]ddd")`.      |
| C      | Mã BS  | text   | Doctor code, must exist in INPUT 4.         |

### INPUT 2 — Lịch họp (rows 6-12)

| Column | Header   | Type | Notes                                   |
| ------ | -------- | ---- | --------------------------------------- |
| E      | Ngày     | date | `dd/mm/yyyy`. Row 5 holds the header.   |
| F      | Thứ      | text | `=TEXT(E{row},"[$-vi-VN]ddd")`.          |
| G      | Nội dung | text | Free-form meeting name.                  |

### INPUT 3 — Lịch trực cấp cứu (rows 16-21)

Same shape as INPUT 1, header on row 15.

### INPUT 4 — Danh sách bác sĩ (rows 16-21)

| Column | Header             | Type | Notes                                       |
| ------ | ------------------ | ---- | ------------------------------------------- |
| E      | Mã BS              | text | Unique doctor code, e.g. `BS04`.            |
| F      | Tên bác sĩ         | text | Full Vietnamese name with diacritics.        |
| G      | Loại PT phụ trách  | text | One of `Ghép gan`, `Cắt gan`, `Tái thùy`, `Cắt gan S`. |

### INPUT 5 — Danh sách bệnh nhân (rows 26-61)

| Column | Header        | Type | Notes                                            |
| ------ | ------------- | ---- | ------------------------------------------------ |
| A      | Mã BN         | text | Unique patient ID, e.g. `BN001`.                 |
| B      | Tên bệnh nhân | text | Full Vietnamese name with diacritics.             |
| C      | Loại PT       | text | Data-validated against the 4 procedure types.    |

Header on row 25. Up to 36 patients fit (rows 26-61).

### KHU VỰC TÍNH TOÁN (rows 63-100)

Formula-driven compute block. Headers on row 64; one computed row per patient on rows 65-100. Columns:

`Mã BN | Tên BN | Loại PT | Mã BS dự kiến | Tên BS | Ngày dự kiến | Thứ | Check phòng khám | Check họp | Check trực | Check loại PT | Tổng OK | Số ca trong ngày | Kết luận`

The `Kết luận` cell evaluates to `Đủ điều kiện` (green) when every check is OK and `M{row}<=11`, otherwise `Loại` (red).

### OUTPUT CUỐI CÙNG (rows 105-140)

Pulls forward only `Đủ điều kiện` rows from KHU VỰC TÍNH TOÁN. Columns:

`Ngày mổ | Thứ | Bác sĩ | Bệnh nhân | Loại PT | Kết luận`

### Summary (row 142)

`Tổng ca đủ điều kiện | =COUNTIF(... "Đủ điều kiện") | Tổng ca bị loại | =COUNTIF(... "Loại") |  | Ca/ngày tối đa | 11`

## Classification rules

Use `prompts/classify.md` to map every incoming user message to exactly one intent:

- `add_clinic_schedule` — add or update a row in INPUT 1.
- `add_meeting` — add or update a row in INPUT 2.
- `add_duty` — add or update a row in INPUT 3 (emergency duty).
- `add_doctor` — add or update a doctor row in INPUT 4 (roster / PT coverage).
- `add_patient` — add or update a patient row in INPUT 5.
- `run_schedule` — recompute / re-verify KHU VỰC TÍNH TOÁN and OUTPUT CUỐI CÙNG (since formulas auto-recalc, this typically means re-open the workbook and report counts).
- `query` — read a value (no writes).
- `off_topic` — anything else.

If the model's confidence is low or the input could fit two intents, treat it as ambiguous and ask one clarification before writing.

## Confirmation policy

- Always send the bilingual confirmation in `prompts/confirm.md` before any write.
- The confirmation must include the target section (e.g. `INPUT!INPUT 5`), target row (or "next free row"), and a 1-line summary of what will be written.
- Wait for an explicit `Có`/`Yes`. Treat anything else as a decline.
- For `run_schedule`, the confirmation summarises which sheet will be inspected (`OUTPUT`) and reports the counts of `Đủ điều kiện` and `Loại` after the formulas refresh.

## Clarification policy

- Ask **one** short question at a time.
- Ask only for the minimum missing field needed to write safely. Examples:
  - `Loại PT của bệnh nhân là gì?` when procedure type is missing.
  - `Bạn muốn ghi vào INPUT 1 (lịch phòng khám) hay INPUT 3 (lịch trực)?` when the target is ambiguous.
- Never guess between two equally plausible targets.

## Refusal policy

Use `prompts/refuse.md`. Send both VI and EN strings as one message and include the 3 scheduling suggestions. Do not engage further with the off-topic request.

## Algorithm reference

The full conflict-check and date-assignment logic lives in `algorithm.md`. Most of the algorithm is implemented as Excel formulas inside KHU VỰC TÍNH TOÁN; the skill's job is to keep INPUT 1-5 clean and report the verdicts that the formulas produce.
