# Scheduling algorithm

The scheduler delegates most of the work to Excel formulas inside the `TÍNH TOÁN` sheet (with cross-sheet references to `INPUT` and results surfaced on `OUTPUT`). This document describes how the formulas behave so the skill can reason about each verdict and explain results to the user.

## Inputs (all on sheet `INPUT`)

### INPUT 1 — Lịch phòng khám (rows 6-12)

Columns `A` (Ngày), `B` (Thứ, formula), `C` (Mã BS).

### INPUT 2 — Lịch họp (rows 6-12)

Columns `E` (Ngày), `F` (Thứ, formula), `G` (Nội dung).

### INPUT 3 — Lịch trực cấp cứu (rows 16-21)

Columns `A` (Ngày), `B` (Thứ, formula), `C` (Mã BS).

### INPUT 4 — Danh sách bác sĩ (rows 16-21)

Columns `E` (Mã BS), `F` (Tên bác sĩ), `G` (Loại PT phụ trách — one of `Ghép gan`, `Cắt gan`, `Tái thùy`, `Cắt gan S`).

### INPUT 5 — Danh sách bệnh nhân (rows 26-61)

Columns `A` (Mã BN), `B` (Tên bệnh nhân), `C` (Loại PT, data-validated).

## Compute block — KHU VỰC TÍNH TOÁN on sheet `TÍNH TOÁN` (rows 65-100)

One row per patient, pulling from `INPUT`. Each compute row has 14 columns:

| Col | Header              | Formula (row `r` on `TÍNH TOÁN`, where `r = 65 + i` and `src = 26 + i` on `INPUT`) |
| --- | ------------------- | ---------------------------------------------------------------------------------- |
| A   | Mã BN               | `=INPUT!A{src}`                                                                    |
| B   | Tên BN              | `=INPUT!B{src}`                                                                    |
| C   | Loại PT             | `=INPUT!C{src}`                                                                    |
| D   | Mã BS dự kiến       | Round-robin rule keyed by `Loại PT` (see below).                                   |
| E   | Tên BS              | `=XLOOKUP(D{r},INPUT!$E$16:$E$21,INPUT!$F$16:$F$21,"")`                          |
| F   | Ngày dự kiến        | Seeded date cycling 2026-05-11..2026-05-14.                                        |
| G   | Thứ                 | `=TEXT(F{r},"[$-vi-VN]dddd")`                                                      |
| H   | Check phòng khám    | `=IF(COUNTIFS(INPUT!$A$6:$A$12,F{r},INPUT!$C$6:$C$12,D{r})>0,"Bận phòng khám","OK")` |
| I   | Check họp           | `=IF(COUNTIF(INPUT!$E$6:$E$12,F{r})>0,"Bận họp","OK")`                           |
| J   | Check trực          | `=IF(COUNTIFS(INPUT!$A$16:$A$21,F{r},INPUT!$C$16:$C$21,D{r})>0,"Bận trực","OK")` |
| K   | Check loại PT       | `=IF(XLOOKUP(D{r},INPUT!$E$16:$E$21,INPUT!$G$16:$G$21,"")=C{r},"OK","Sai loại PT")` |
| L   | Tổng OK             | `=IF(AND(H{r}="OK",I{r}="OK",J{r}="OK",K{r}="OK"),"OK","Bận lịch")` |
| M   | Số ca trong ngày    | `=COUNTIF($F$65:$F$100,F{r})`                                                      |
| N   | Kết luận            | `=IF(AND(L{r}="OK",M{r}<=11),"Đủ điều kiện","Loại")`                             |

### Doctor assignment (column D)

```
=IF(C{r}="Ghép gan",
   IF(MOD(COUNTIF($C$65:C{r},"Ghép gan"),2)=1,"BS04","BS02"),
 IF(C{r}="Cắt gan",
   IF(MOD(COUNTIF($C$65:C{r},"Cắt gan"),2)=1,"BS05","BS09"),
 IF(C{r}="Tái thùy","BS08",
 IF(C{r}="Cắt gan S","BS06",""))))
```

- `Ghép gan` cases alternate between `BS04` and `BS02`.
- `Cắt gan` cases alternate between `BS05` and `BS09`.
- `Tái thùy` always goes to `BS08`.
- `Cắt gan S` always goes to `BS06`.
- Anything else yields an empty doctor code (treated as `Loại`).

### Daily cap

The cap of 11 surgeries per day is hardcoded in the `Kết luận` formula (`M{r}<=11`). To change the cap, edit the formulas in column N of rows 65..100.

## OUTPUT CUỐI CÙNG on sheet `OUTPUT` (rows 105-140)

For each compute row `src` on `TÍNH TOÁN` (65..100), the output row `r = src + 40` on `OUTPUT` carries the verdict forward only when `Kết luận = "Đủ điều kiện"`:

| Col | Formula                                                                              |
| --- | ------------------------------------------------------------------------------------ |
| A   | `=IF(TÍNH TOÁN!$N{src}="Đủ điều kiện",TEXT(TÍNH TOÁN!$F{src},"yyyy-mm-dd"),"")`     |
| B   | `=IF(TÍNH TOÁN!$N{src}="Đủ điều kiện",TÍNH TOÁN!$G{src},"")`                        |
| C   | `=IF(TÍNH TOÁN!$N{src}="Đủ điều kiện",TÍNH TOÁN!$E{src},"")`                        |
| D   | `=IF(TÍNH TOÁN!$N{src}="Đủ điều kiện",TÍNH TOÁN!$B{src},"")`                        |
| E   | `=IF(TÍNH TOÁN!$N{src}="Đủ điều kiện",TÍNH TOÁN!$C{src},"")`                        |
| F   | `=TÍNH TOÁN!$N{src}`                                                                 |

Conditional formatting paints `Đủ điều kiện` green (`#D4EDDA`) and `Loại` red (`#F8D7DA`).

## Summary row (R142 on `OUTPUT`)

| A                           | B (formula)                                               | C                  | D (formula)                                           | F                | G   |
| --------------------------- | --------------------------------------------------------- | ------------------ | ----------------------------------------------------- | ---------------- | --- |
| `Tổng ca đủ điều kiện`      | `=COUNTIF(TÍNH TOÁN!$N$65:$N$100,"Đủ điều kiện")`        | `Tổng ca bị loại`  | `=COUNTIF(TÍNH TOÁN!$N$65:$N$100,"Loại")`            | `Ca/ngày tối đa` | `11` |

## Verdicts the skill must explain

- **`Bận phòng khám`** — the proposed doctor already has a clinic block on the proposed date (INPUT 1).
- **`Bận họp`** — the proposed date is fully consumed by a meeting (INPUT 2).
- **`Bận trực`** — the proposed doctor is on emergency duty that day (INPUT 3).
- **`Sai loại PT`** — the proposed doctor's `Loại PT phụ trách` in INPUT 4 does not match the patient's procedure type.
- **`Loại`** — at least one of the above fired, or the day already has 11+ surgeries.
- **`Đủ điều kiện`** — every check passed and the day still has capacity.

## Edge cases

- **Empty INPUT 4 doctor row** — `Tên BS` resolves to `""` and `Check loại PT` returns `Sai loại PT`. Fix by completing the doctor row.
- **Unknown `Loại PT` on INPUT 5** — column `D` (Mã BS dự kiến) returns `""`, leading to `Loại`. Fix by re-classifying the patient.
- **Date in `Ngày dự kiến` falls on a holiday or non-working day** — the formulas do not check weekday or holidays; the user must manually re-seed the date in column F of KHU VỰC TÍNH TOÁN. The skill should warn when a date appears non-working.
- **Daily cap exceeded** — the 12th patient onward on the same date in column F will flip to `Loại` automatically.

## What `run_schedule` actually does

Because every check is a live formula, `run_schedule` consists of:

1. Open the workbook through the `excel` MCP server.
2. Trigger a recalculation (or simply re-read the cells; Excel recalculates on open).
3. Report the values of `OUTPUT!B142` (đủ điều kiện) and `OUTPUT!D142` (loại) plus the first few `Đủ điều kiện` rows from `OUTPUT`.
4. If many rows are `Loại`, recommend which INPUT section to edit (e.g. add more doctors with the right `Loại PT phụ trách`, or shift the proposed dates in `TÍNH TOÁN`).

The skill must never edit cells inside KHU VỰC TÍNH TOÁN or OUTPUT CUỐI CÙNG.
