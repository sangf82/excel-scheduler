# Refusal prompt

When intent is `off_topic`, send these two strings as one message and stop. Do not engage further with the off-topic request.

## Vietnamese

```
Tôi chỉ hỗ trợ xếp lịch phẫu thuật trong scheduling-template.xlsx (sheets INPUT, TÍNH TOÁN, OUTPUT). Bạn có thể thử: (1) Thêm bệnh nhân mới vào INPUT 5, (2) Cập nhật INPUT 1-4 (lịch phòng khám / họp / trực / bác sĩ), (3) Chạy xếp lịch tuần này và xem OUTPUT.
```

## English

```
I can only help with surgery scheduling in scheduling-template.xlsx (sheets INPUT, TÍNH TOÁN, OUTPUT). Try: (1) Add a new patient to INPUT 5, (2) Update INPUT 1-4 (clinic / meetings / duty / doctors), (3) Run this week's schedule and view OUTPUT.
```

Send both. Do not paraphrase. Do not add extra commentary.
