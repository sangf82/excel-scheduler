# Workbook schema — 4 sheets

All data entry lives on the `INPUT` sheet. The `TÍNH TOÁN` and `OUTPUT` sheets are formula-driven and must never be edited by hand.

## INPUT layout (horizontal, all 5 sections in one row)

Sheet `INPUT` has a single title row, a single subtitle row, one shared header row, then data rows that span all five sections side-by-side.

### Row structure

| Row | Content |
| --- | --- |
| 1 | Title banner (merged A1:S1) |
| 2 | Subtitle banner (merged A2:S2) |
| 3 | Section banners (5 blocks, merged 3 cols each) |
| 4 | Column headers (all 5 sections) |
| 5–11 | Data for sections 1, 2, 3, 4 + start of section 5 |
| 12–40 | Only section 5 continues; sections 1–4 rows are styled blank |

### Column mapping

| Section | Columns | Data rows | Banner | Header |
| --- | --- | --- | --- | --- |
| 1 — Lịch phòng khám | A, B, C | 5–11 | Row 3 (A3:C3) | Row 4 |
| (spacer) | D | — | — | — |
| 2 — Lịch họp | E, F, G | 5–11 | Row 3 (E3:G3) | Row 4 |
| (spacer) | H | — | — | — |
| 3 — Lịch trực cấp cứu | I, J, K | 5–10 | Row 3 (I3:K3) | Row 4 |
| (spacer) | L | — | — | — |
| 4 — Danh sách bác sĩ | M, N, O | 5–10 | Row 3 (M3:O3) | Row 4 |
| (spacer) | P | — | — | — |
| 5 — Danh sách bệnh nhân | Q, R, S | 5–40 | Row 3 (Q3:S3) | Row 4 |

Spacer columns D, H, L, P have width 2, no borders, no fill.

### 1 — Lịch phòng khám (columns A–C, rows 5–11)

| Column | Header | Example | Notes |
| --- | --- | --- | --- |
| A | Ngày | `15/05/2026` | Format `dd/mm/yyyy`. |
| B | Thứ | `T6` | Formula `=TEXT(A{row},"[$-vi-VN]ddd")`. Do not type by hand. |
| C | Mã BS | `BS04` | Must match a `Mã BS` in section 4. |

### 2 — Lịch họp (columns E–G, rows 5–11)

| Column | Header | Example | Notes |
| --- | --- | --- | --- |
| E | Ngày | `15/05/2026` | Format `dd/mm/yyyy`. |
| F | Thứ | `T6` | `=TEXT(E{row},"[$-vi-VN]ddd")`. |
| G | Nội dung | `Hội thảo chuyên môn` | Free-form description. |

A row here marks the **entire day** as a meeting day; every doctor is considered busy.

### 3 — Lịch trực cấp cứu (columns I–K, rows 5–10)

| Column | Header | Example | Notes |
| --- | --- | --- | --- |
| I | Ngày | `15/05/2026` | Format `dd/mm/yyyy`. |
| J | Thứ | `T6` | `=TEXT(I{row},"[$-vi-VN]ddd")`. |
| K | Mã BS | `BS06` | Doctor on emergency duty on that date. |

### 4 — Danh sách bác sĩ (columns M–O, rows 5–10)

| Column | Header | Example | Notes |
| --- | --- | --- | --- |
| M | Mã BS | `BS04` | Unique doctor code. |
| N | Tên bác sĩ | `BS Nguyễn Minh An` | Full Vietnamese name with diacritics. |
| O | Loại PT phụ trách | `Ghép gan` | One of `Ghép gan`, `Cắt gan`, `Tái thùy`, `Cắt gan S`. |

The seed roster covers all four procedure types:

```
Mã BS  Tên bác sĩ            Loại PT phụ trách
BS04   BS Nguyễn Minh An     Ghép gan
BS02   BS Hoàng Hải Nam      Ghép gan
BS05   BS Trần Quốc Bình     Cắt gan
BS09   BS Đỗ Thanh Sơn       Cắt gan
BS08   BS Lê Thu Cúc         Tái thùy
BS06   BS Phạm Đức Dũng      Cắt gan S
```

### 5 — Danh sách bệnh nhân (columns Q–S, rows 5–40)

| Column | Header | Example | Notes |
| --- | --- | --- | --- |
| Q | Mã BN | `BN001` | Unique patient ID. |
| R | Tên bệnh nhân | `Nguyễn Văn A` | Full Vietnamese name with diacritics. |
| S | Loại PT | `Ghép gan` | Data validation pulls from `Ghép gan,Cắt gan,Tái thùy,Cắt gan S`. |

Capacity: 36 patients (rows 5–40). The auto-filter on `Q4:S40` lets users sort and filter patients in place.

## TÍNH TOÁN and OUTPUT

`TÍNH TOÁN` rows 63–100 and `OUTPUT` rows 103–140 are **read-only** for the skill. They are populated entirely by Excel formulas — see `algorithm.md` for the formula behaviour. The skill must never write into these ranges.

The summary row on `OUTPUT` (`A142..G142`) shows `Tổng ca đủ điều kiện`, `Tổng ca bị loại`, and the daily cap (`11`).

## Copyable doctor + patient seed

```
Section 4 (M5:O10):
BS04   BS Nguyễn Minh An     Ghép gan
BS02   BS Hoàng Hải Nam      Ghép gan
BS05   BS Trần Quốc Bình     Cắt gan
BS09   BS Đỗ Thanh Sơn       Cắt gan
BS08   BS Lê Thu Cúc         Tái thùy
BS06   BS Phạm Đức Dũng      Cắt gan S

Section 5 sample (Q5:S7):
BN001  Nguyễn Văn A          Ghép gan
BN002  Trần Thị B            Ghép gan
BN003  Lê Văn C              Cắt gan
```
