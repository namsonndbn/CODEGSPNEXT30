# GSP NEXT 30 - CEO Dashboard

Dashboard theo dõi realtime cho ban lãnh đạo GSP NEXT 30: Workstream, Action, RAG, Risk/SOS và các việc CEO cần chốt.

## Link truy cập

- **GitHub Pages (Dashboard):** https://namsonndbn.github.io/CODEGSPNEXT30/
- **Google Apps Script (Backend API):** https://script.google.com/macros/s/AKfycbydqHgolCAEXd34rBDMvudvFET5No28uFtuI3xrXTw-k2ZPHfjBAnwXVfftQuiLMGj6cg/exec
- **GitHub Repo:** https://github.com/namsonndbn/CODEGSPNEXT30

## Kiến trúc hệ thống

```
Google Sheets (Dữ liệu)
        ↓
Google Apps Script (Backend API - ma.gs)
        ↓
GitHub Pages (Frontend - index.html)
        ↓
Người dùng (trình duyệt)
```

## Cấu trúc file

| File | Mô tả |
|------|--------|
| `index.html` | Giao diện dashboard (HTML + CSS + JS) |
| `ma.gs` | Backend Google Apps Script (API + xử lý dữ liệu) |
| `appsscript.json` | Cấu hình Google Apps Script |

## Tính năng

- Xem tổng quan 6 Workstream theo RAG (Green / Amber / Red)
- Danh sách Action Items, deadline, PIC
- Theo dõi Risk / SOS
- CEO cần chốt
- Chat & Comment trực tiếp trên dashboard
- Chế độ họp (ẩn thông tin nhạy cảm)
- Tự động cập nhật mỗi 2 phút

## Cách cập nhật code

```powershell
# 1. Sửa code trong thư mục C:\Users\Admin\Desktop\CODEGSPNEXT30
# 2. Push lên GitHub
git add .
git commit -m "Mô tả thay đổi"
git push

# 3. Đẩy code lên Google Apps Script
clasp push --force
clasp deploy --deploymentId "AKfycbzCzBEykV5gzIfuVm_Y0etw8AhtRT0_QIU8j1DDiY8l48YhjtbHdc__ZEn3IxPwkS9fkw" --description "Mô tả"
```

## Thông tin kỹ thuật

- **Chủ dự án:** namsonndbn
- **Google Sheet nguồn:** GSP_NEXT30_PMO_Tracker_6_Truc_Dashboard_Source
- **Múi giờ:** Asia/Ho_Chi_Minh
- **Runtime GAS:** V8
