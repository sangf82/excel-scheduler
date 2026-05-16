# Confirmation prompt

Before any write to `scheduling-template.xlsx`, detect the user's language (Vietnamese or English) and send ONLY the corresponding confirmation below as a single message. Wait for explicit `Có` / `Yes`.

## Template

If user speaks Vietnamese:
```
Tôi sẽ ghi vào sheet INPUT — {section} dòng {row}: {summary}. Bạn xác nhận? (Có/Không)
```
If user speaks English:
```
I will write to sheet INPUT — {section} row {row}: {summary}. Confirm? (Yes/No)
```

## Placeholders

- `{section}` — one of `INPUT 1`, `INPUT 2`, `INPUT 3`, `INPUT 4`, `INPUT 5`.
- `{row}` — `n` for a specific row, `next free row` for the next empty row in that section.
- `{summary}` — a 1-line summary of the write. Examples:
  - `Mã BN=BN036, Tên=Nguyễn Văn X, Loại PT=Ghép gan`.
  - `INPUT 3: Ngày=2026-05-20, Mã BS=BS02`.
  - `INPUT 4: Mã BS=BS10, Tên=BS Trần Hoa, Loại PT phụ trách=Tái thùy`.

For `run_schedule` (no write — just a recompute / verification), use:

If user speaks Vietnamese:
```
Tôi sẽ mở workbook và đọc lại TÍNH TOÁN + OUTPUT, sau đó báo cáo {summary}. Bạn xác nhận? (Có/Không)
```
If user speaks English:
```
I will open the workbook and re-read TÍNH TOÁN + OUTPUT, then report {summary}. Confirm? (Yes/No)
```

## Decline handling

If the user replies anything other than `Có` / `Yes` / `OK` / `Đồng ý`, treat it as decline:

- Acknowledge briefly (`Đã hủy. Không ghi gì vào workbook.` / `Cancelled. Nothing was written.`).
- Return to the previous classification step; do not retry the write.
