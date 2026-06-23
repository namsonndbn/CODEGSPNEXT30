# Bản đồ dữ liệu hệ thống — GSP NEXT 30 CEO Command Center
> Rà soát tự động từ ma.gs + index.html · Ngày: 2026-06-23
> **Không đổi tên / xóa / di chuyển bất kỳ sheet nào ở bước này.**

---

## 1. Bảng tổng hợp tất cả sheet

| Sheet | Nhóm | Module hiển thị | Nhập gì | Ai cập nhật | Hàm đọc chính | Hàm ghi chính | Có thể ẩn | Ghi chú |
|---|---|---|---|---|---|---|---|---|
| `01_Brand` | NHẬP | PMO GSP NEXT 30 | Action, Owner, Deadline, RAG, Status, Risk, CEO Decision | PIC01 – Kinh Doanh | `getDashboardData` | — (PIC tự nhập) | ❌ Không | Sheet gốc của PIC, cần mở để nhập |
| `02_Sales` | NHẬP | PMO GSP NEXT 30 | Như trên | PIC02 – Sales Tech | `getDashboardData` | — | ❌ Không | |
| `03_Operation` | NHẬP | PMO GSP NEXT 30 | Như trên | PIC03-06 – SCM/MH/CĐ/QA | `getDashboardData` | — | ❌ Không | |
| `04_Digital_AI` | NHẬP | PMO GSP NEXT 30 | Như trên | PIC05/07 – IT/ERP/AI | `getDashboardData` | — | ❌ Không | |
| `05_People_Culture` | NHẬP | PMO GSP NEXT 30 | Như trên | PIC09/12 – PC/HCNS | `getDashboardData` | — | ❌ Không | |
| `06_System_KPI` | NHẬP | PMO GSP NEXT 30 | Như trên | PMO tổng hợp | `getDashboardData` | — | ❌ Không | |
| `THEO_DOI_DU_AN_CEO` | NHẬP | CEO Project Tracker | DỰ ÁN, HẠNG MỤC, CÔNG VIỆC, CẬP NHẬT, QUYẾT ĐỊNH (37 cột) | PMO + PIC Lead + CEO | `getCEOProjectTrackerData` | `importCEOProjectSeedData`, `setupCEOProjectTrackerSheet` | ❌ Không | Sheet quan trọng nhất cho CEO Tracker |
| `MEETING_MINUTES` | NHẬP | Biên bản họp | Meeting ID, Ngày họp, Chủ đề, Chỉ đạo CEO, Quyết định, Action, Risk, File biên bản | PMO sau mỗi cuộc họp | `getMeetingMinutesData`, `getMeetingCenterData` | `setupGspNext30Database` (chỉ tạo header) | ❌ Không | Nguồn dữ liệu cho Risk/SOS và CEO Decision |
| `MEETING_TOPICS` | NHẬP | Biên bản họp | Topic ID, Tên topic, Nhóm, Mô tả, Owner, Trạng thái | PMO trước/sau họp | `getMeetingCenterData`, `getDashboardData` | — | ❌ Không | Sheet chưa qua setup tự động |
| `MEETING_ACTIONS` | NHẬP | Biên bản họp | Action ID, Topic ID, Nội dung, Owner, Deadline, RAG, CEO cần chốt | PMO sau họp | `getMeetingCenterData`, `getDashboardData` | — | ❌ Không | Sheet chưa qua setup tự động |
| `CEO_DECISIONS` | NHẬP | CEO Decision Board | Decision ID, Nguồn, Nội dung cần chốt, PIC, Deadline, Trạng thái, Quyết định CEO | PMO khi cần CEO quyết | `getCommandCenterData`, `getDashboardData` | `setupGspNext30Database` (tạo header) | ❌ Không | CEO điền cột "Quyết định CEO" + "Ngày chốt" |
| `RISK_SOS` | NHẬP | Risk/SOS Center | Risk ID, Mô tả, PIC xử lý, Mức độ, RAG, Deadline, Trạng thái, Biện pháp | PIC/PMO khi phát sinh | `getCommandCenterData`, `getDashboardData` | `setupGspNext30Database` (tạo header) | ❌ Không | Ưu tiên cao khi RAG Đỏ |
| `MASTER_PIC` | NHẬP/THAM CHIẾU | Data Quality, Command Center | Tên PIC, Workstream, Nhà máy, Email, SĐT, Vai trò | Admin hệ thống | `getCommandCenterData` | `setupGspNext30Database` (tạo header) | ⚠️ Có thể ẩn | Danh mục PIC — nhập 1 lần, ít thay đổi |
| `PMO_90D_ACTIONS` | NHẬP | PMO (phụ) | Mã dòng, Workstream, Action, PIC, Deadline, RAG, Status | PMO (nhập bổ sung) | `getCommandCenterData` | `setupGspNext30Database` (tạo header) | ⚠️ Có thể ẩn | Sheet dự phòng; dữ liệu chính qua 01-06 hoặc PIC sheets |
| `CEO_PROJECTS` | NHẬP | CEO Decision / Risk | Project ID, Tên dự án, PIC Lead, Deadline, RAG, Vướng mắc, CEO cần chốt | PMO (legacy?) | `getCEOProjectsData`, `getCommandCenterData` | `setupGspNext30Database` (tạo header) | ⚠️ Có thể ẩn | Có thể trùng với THEO_DOI_DU_AN_CEO — cần xác nhận |
| `CEO_PROJECT_UPDATES` | NHẬP | CEO Project Tracker (phụ) | Update ID, Project ID, % HT, Milestone, Vướng mắc, CEO cần biết, RAG | PIC Lead sau milestone | `getCommandCenterData` | `setupGspNext30Database` (tạo header) | ⚠️ Có thể ẩn | Nếu dùng THEO_DOI_DU_AN_CEO thì sheet này dư |
| `CEO_Comments` | HỆ THỐNG | PMO (comment bar) | Comment ID, Người gửi, Lời nhắn, Like, Trạng thái | GAS tự ghi khi user post comment | `getDashboardComments`, `likeDashboardComment` | `saveDashboardComment`, `deleteDashboardComment` | ✅ Có thể ẩn | Không nhập thủ công — GAS quản lý hoàn toàn |
| `PMO_ALL_ACTIONS` | HỆ THỐNG | PMO GSP NEXT 30 | 24 cột chuẩn, sync từ 16 PIC sheets | GAS tự sync (`syncPicSheetsToPmoAllActions`) | `getDashboardData` (Priority 1) | `syncPicSheetsToPmoAllActions` | ✅ Có thể ẩn | Không nhập thủ công — KHÔNG SỬA trực tiếp |
| `PIC01_KinhDoanh` | HỆ THỐNG | PMO GSP NEXT 30 | 24 cột PMO chuẩn (import từ Google Sheet PIC01) | GAS `importPicPlanToDatabase` | `getDashboardData`, `getPicPmoData` | `importPicPlanToDatabase` | ✅ Có thể ẩn | Không nhập thủ công — GAS clear và ghi lại khi import |
| `PIC02_SalesTech` | HỆ THỐNG | PMO | ← như trên | GAS | `getDashboardData` | `importPicPlanToDatabase` | ✅ Có thể ẩn | |
| `PIC03_SCM` | HỆ THỐNG | PMO | ← như trên | GAS | `getDashboardData` | `importPicPlanToDatabase` | ✅ Có thể ẩn | |
| `PIC04_MuaHang` | HỆ THỐNG | PMO | ← như trên | GAS | `getDashboardData` | `importPicPlanToDatabase` | ✅ Có thể ẩn | |
| `PIC05_CongNghe` | HỆ THỐNG | PMO | ← như trên | GAS | `getDashboardData` | `importPicPlanToDatabase` | ✅ Có thể ẩn | |
| `PIC06_CoDien_QLTB` | HỆ THỐNG | PMO | ← như trên | GAS | `getDashboardData` | `importPicPlanToDatabase` | ✅ Có thể ẩn | |
| `PIC07_IT_ERP` | HỆ THỐNG | PMO | ← như trên | GAS | `getDashboardData` | `importPicPlanToDatabase` | ✅ Có thể ẩn | |
| `PIC08_QA` | HỆ THỐNG | PMO | ← như trên | GAS | `getDashboardData` | `importPicPlanToDatabase` | ✅ Có thể ẩn | |
| `PIC09_PC` | HỆ THỐNG | PMO | ← như trên | GAS | `getDashboardData` | `importPicPlanToDatabase` | ✅ Có thể ẩn | |
| `PIC10_GiamSatTuanThu` | HỆ THỐNG | PMO | ← như trên | GAS | `getDashboardData` | `importPicPlanToDatabase` | ✅ Có thể ẩn | |
| `PIC11_KiemSoatNoiBo` | HỆ THỐNG | PMO | ← như trên | GAS | `getDashboardData` | `importPicPlanToDatabase` | ✅ Có thể ẩn | |
| `PIC12_HCNS` | HỆ THỐNG | PMO | ← như trên | GAS | `getDashboardData` | `importPicPlanToDatabase` | ✅ Có thể ẩn | |
| `PIC13_TaiChinhKeToan` | HỆ THỐNG | PMO | ← như trên | GAS | `getDashboardData` | `importPicPlanToDatabase` | ✅ Có thể ẩn | |
| `PIC14_GS1` | HỆ THỐNG | PMO | ← như trên | GAS | `getDashboardData` | `importPicPlanToDatabase` | ✅ Có thể ẩn | |
| `PIC15_GS5` | HỆ THỐNG | PMO | ← như trên | GAS | `getDashboardData` | `importPicPlanToDatabase` | ✅ Có thể ẩn | |
| `PIC16_GS6` | HỆ THỐNG | PMO | ← như trên | GAS | `getDashboardData` | `importPicPlanToDatabase` | ✅ Có thể ẩn | |
| `TO_ACTIONS` | HỆ THỐNG | TO Action Center | 22 cột action (import từ TO report) | GAS `importKHHDToActions_` | `getTOActionsData`, `getToActionCenterData` | `importKHHDToActions_` | ✅ Có thể ẩn | GAS import từ file TO của nhà máy |
| `TO_ISSUES` | HỆ THỐNG | TO Action Center | 16 cột issue (import từ TO report) | GAS `importHienTrangToIssues_` | `getToActionCenterData` | `importHienTrangToIssues_` | ✅ Có thể ẩn | GAS import từ file TO của nhà máy |
| `PMO_IMPORT_LOG` | HỆ THỐNG | — (ẩn) | Log import: Timestamp, PIC, URL, Rows, Status | GAS tự ghi | — | `pmoLogImport_` | ✅ Có thể ẩn | Chỉ dùng để debug import |
| `PMO_SOURCE_FILES` | HỆ THỐNG | — (ẩn) | Source URL, Last Imported, Row Count per PIC | GAS tự ghi | `importPicPlanToDatabase` | `pmoUpdateSourceFiles_` | ✅ Có thể ẩn | Chỉ dùng để tránh import trùng |
| `MASTER_STATUS` | THAM CHIẾU | Data Quality | Giá trị status, Màu HEX, Nhóm (22 dòng seed) | Admin (ít khi thay đổi) | `getCommandCenterData` | `setupGspNext30Database` (seed 22 dòng) | ✅ Có thể ẩn | Danh mục màu RAG — không nhập thủ công thường xuyên |
| `CONFIG` | THAM CHIẾU | — (nội bộ) | Key-Value: APP_VERSION, TIMEZONE, PMO_REVIEW_DAY, v.v. | Admin hệ thống | `getCommandCenterData` | `setupGspNext30Database` (seed 8 dòng) | ✅ Có thể ẩn | Cấu hình hệ thống |
| `00_Dashboard_CEO` | CHƯA XÁC ĐỊNH | — | Không tìm thấy trong code | ? | — | — | ❓ | Có thể là sheet cũ hoặc draft |
| `AI_Review_Input` | CHƯA XÁC ĐỊNH | — | Không tìm thấy trong code | ? | — | — | ❓ | Có thể là sheet thử nghiệm AI |
| `AI_Review_Output` | CHƯA XÁC ĐỊNH | — | Không tìm thấy trong code | ? | — | — | ❓ | Có thể là sheet thử nghiệm AI |

---

## 2. Phân loại chi tiết

### 🟢 NHẬP DỮ LIỆU — Người dùng cần nhập/cập nhật
> Đây là các sheet cần mở Google Sheets và nhập thủ công

| # | Sheet | Người nhập | Tần suất | Module |
|---|---|---|---|---|
| 1 | `01_Brand` | PIC01 | Hàng tuần / sau milestone | PMO GSP NEXT 30 |
| 2 | `02_Sales` | PIC02 | Hàng tuần | PMO GSP NEXT 30 |
| 3 | `03_Operation` | PIC03-06 | Hàng tuần | PMO GSP NEXT 30 |
| 4 | `04_Digital_AI` | PIC05/07 | Hàng tuần | PMO GSP NEXT 30 |
| 5 | `05_People_Culture` | PIC09/12 | Hàng tuần | PMO GSP NEXT 30 |
| 6 | `06_System_KPI` | PMO tổng hợp | Hàng tuần | PMO GSP NEXT 30 |
| 7 | `THEO_DOI_DU_AN_CEO` | PMO + PIC Lead | Hàng ngày / mỗi update | CEO Project Tracker |
| 8 | `MEETING_MINUTES` | PMO | Sau mỗi cuộc họp | Biên bản họp |
| 9 | `MEETING_TOPICS` | PMO | Trước/sau họp | Biên bản họp |
| 10 | `MEETING_ACTIONS` | PMO | Sau mỗi cuộc họp | Biên bản họp |
| 11 | `CEO_DECISIONS` | PMO | Khi phát sinh | CEO Decision Board |
| 12 | `RISK_SOS` | PIC / PMO | Khi phát sinh | Risk/SOS Center |
| 13 | `MASTER_PIC` | Admin | Khi có PIC mới | Data Quality |
| 14 | `PMO_90D_ACTIONS` | PMO | Khi bổ sung action | PMO (bổ sung) |
| 15 | `CEO_PROJECTS` | PMO | Khi tạo mới dự án | CEO Decision / Risk |
| 16 | `CEO_PROJECT_UPDATES` | PIC Lead | Sau milestone | CEO Tracker (phụ) |

### 🔵 HỆ THỐNG — Apps Script tự đọc/ghi, **KHÔNG nhập thủ công**
> Nhập vào đây sẽ bị GAS ghi đè hoặc gây lỗi

| # | Sheet | GAS ghi khi nào | Module |
|---|---|---|---|
| 1 | `CEO_Comments` | Khi user post comment trên dashboard | PMO (comment bar) |
| 2 | `PMO_ALL_ACTIONS` | Chạy `syncPicSheetsToPmoAllActions()` | PMO GSP NEXT 30 |
| 3 | `PIC01_KinhDoanh` | Chạy `importPicPlanToDatabase()` | PMO GSP NEXT 30 |
| 4-16 | `PIC02` … `PIC16_GS6` | Chạy `importPicPlanToDatabase()` | PMO GSP NEXT 30 |
| 17 | `TO_ACTIONS` | Chạy `importKHHDToActions_()` từ file TO report | TO Action Center |
| 18 | `TO_ISSUES` | Chạy `importHienTrangToIssues_()` từ file TO report | TO Action Center |
| 19 | `PMO_IMPORT_LOG` | Sau mỗi lần import PIC | — (log nội bộ) |
| 20 | `PMO_SOURCE_FILES` | Sau mỗi lần import PIC | — (tracking nội bộ) |

### ⚫ THAM CHIẾU — Danh mục/cấu hình, nhập 1 lần, ít thay đổi

| # | Sheet | Dữ liệu | Ai quản lý |
|---|---|---|---|
| 1 | `MASTER_STATUS` | Danh mục trạng thái + màu RAG (22 dòng seed) | Admin hệ thống |
| 2 | `CONFIG` | Cấu hình: version, timezone, ngày review (8 key-value) | Admin hệ thống |

### ❓ CHƯA XÁC ĐỊNH — Không tìm thấy trong code

| # | Sheet | Tình trạng | Khuyến nghị |
|---|---|---|---|
| 1 | `00_Dashboard_CEO` | Không có trong `ma.gs` hay `index.html` | Xác nhận với PMO trước khi ẩn/xóa |
| 2 | `AI_Review_Input` | Không có trong code | Xác nhận — có thể là thử nghiệm AI cũ |
| 3 | `AI_Review_Output` | Không có trong code | Xác nhận — có thể là thử nghiệm AI cũ |

---

## 3. Nguồn dữ liệu cho từng module

| Module dashboard | Sheet nguồn chính | Sheet nguồn phụ |
|---|---|---|
| 🏠 Tổng quan | Tổng hợp từ tất cả | — |
| 📋 PMO GSP NEXT 30 | `01_Brand`→`06_System_KPI` (Priority 3) hoặc `PIC01`-`PIC16` (Priority 2) hoặc `PMO_ALL_ACTIONS` (Priority 1) | `CEO_Comments` |
| 🎯 CEO Project Tracker | `THEO_DOI_DU_AN_CEO` | — |
| ⚡ TO Action Center | `TO_ACTIONS`, `TO_ISSUES` | — |
| 📄 Biên bản họp | `MEETING_MINUTES`, `MEETING_TOPICS`, `MEETING_ACTIONS` | — |
| 🚨 Risk/SOS Center | `RISK_SOS` (nhập thủ công) + `CEO_PROJECTS` + `MEETING_MINUTES` + `TO_ACTIONS` (trích xuất) | — |
| ✅ CEO Decision Board | `CEO_DECISIONS` (nhập thủ công) + `MEETING_MINUTES` + `CEO_PROJECTS` + `TO_ACTIONS` (trích xuất) | — |
| 📊 Data Quality | `PMO_ALL_ACTIONS`, `CEO_PROJECTS`, `TO_ACTIONS`, `MEETING_MINUTES`, `MASTER_STATUS` | `MASTER_PIC` |
| 💬 Luồng chat Chủ tịch | File `.docx` → Python script → `data/luong_chat_chu_tich.json` → nhúng vào `index.html` | — |

---

## 4. Đề xuất sắp xếp tab và màu tab

> **Không thực hiện ngay.** Xác nhận với user trước khi thay đổi màu/thứ tự.

### Thứ tự đề xuất:
```
[Tổng quan] | [01_Brand] [02_Sales] [03_Operation] [04_Digital_AI] [05_People_Culture] [06_System_KPI]
| [THEO_DOI_DU_AN_CEO] | [MEETING_TOPICS] [MEETING_MINUTES] [MEETING_ACTIONS]
| [CEO_DECISIONS] [RISK_SOS]
| [MASTER_PIC] [MASTER_STATUS] [CONFIG]
---ẩn---
| [PMO_ALL_ACTIONS] [PIC01..PIC16] [TO_ACTIONS] [TO_ISSUES] [CEO_Comments] [PMO_IMPORT_LOG] [PMO_SOURCE_FILES]
| [CEO_PROJECTS] [CEO_PROJECT_UPDATES] [PMO_90D_ACTIONS]
---chưa xác định---
| [00_Dashboard_CEO] [AI_Review_Input] [AI_Review_Output]
```

### Màu tab đề xuất:
| Màu | Nhóm | Áp dụng cho |
|---|---|---|
| 🟢 Xanh lá (`#15803d`) | NHẬP DỮ LIỆU | `01_Brand`-`06_System_KPI`, `THEO_DOI_DU_AN_CEO`, `MEETING_*`, `CEO_DECISIONS`, `RISK_SOS` |
| 🔵 Xanh dương (`#1d4ed8`) | HỆ THỐNG | `PMO_ALL_ACTIONS`, `PIC01`-`PIC16`, `TO_ACTIONS`, `TO_ISSUES`, `CEO_Comments`, `PMO_IMPORT_LOG`, `PMO_SOURCE_FILES` |
| ⚫ Xám (`#374151`) | THAM CHIẾU | `MASTER_STATUS`, `CONFIG`, `MASTER_PIC` |
| 🟡 Vàng (`#b45309`) | CHƯA XÁC ĐỊNH | `00_Dashboard_CEO`, `AI_Review_Input`, `AI_Review_Output` |
| 🟠 Cam nhạt | Có thể cũ | `CEO_PROJECTS`, `CEO_PROJECT_UPDATES`, `PMO_90D_ACTIONS` |

---

## 5. Sheet có thể ẩn an toàn (không ảnh hưởng hệ thống)

Ẩn trong Google Sheets: chuột phải tab → "Hide sheet". GAS vẫn đọc/ghi được.

| Sheet | Lý do có thể ẩn |
|---|---|
| `CEO_Comments` | GAS quản lý tự động, không cần thấy |
| `PMO_ALL_ACTIONS` | Chỉ là bảng tổng hợp, không nhập trực tiếp |
| `PIC01` đến `PIC16` | GAS ghi đè, không nhập thủ công |
| `TO_ACTIONS` | GAS import, không nhập thủ công |
| `TO_ISSUES` | GAS import, không nhập thủ công |
| `PMO_IMPORT_LOG` | Chỉ log nội bộ |
| `PMO_SOURCE_FILES` | Chỉ tracking nội bộ |
| `MASTER_STATUS` | Ít thay đổi, có thể ẩn sau khi setup xong |
| `CONFIG` | Ít thay đổi, có thể ẩn sau khi setup xong |

---

## 6. Sheet nghi là cũ — cần xác nhận trước khi xóa

| Sheet | Nghi ngờ vì | Cần xác nhận |
|---|---|---|
| `00_Dashboard_CEO` | Không có trong code, tên gợi ý là phiên bản dashboard cũ | PMO xác nhận còn dùng không |
| `AI_Review_Input` | Không có trong code | PMO xác nhận còn dùng không |
| `AI_Review_Output` | Không có trong code | PMO xác nhận còn dùng không |
| `CEO_PROJECTS` | Có thể trùng chức năng với `THEO_DOI_DU_AN_CEO` | Xác nhận có còn nhập vào đây không |
| `CEO_PROJECT_UPDATES` | Có thể trùng với cột CẬP NHẬT trong `THEO_DOI_DU_AN_CEO` | Xác nhận |
| `PMO_90D_ACTIONS` | Không rõ đang dùng hay không — PMO_ALL_ACTIONS có vẻ là nguồn chính | Xác nhận |

---

## 7. Rủi ro khi đổi tên hoặc xóa

| Sheet | Rủi ro nếu đổi tên / xóa |
|---|---|
| `THEO_DOI_DU_AN_CEO` | **Nghiêm trọng** — CEO Tracker sẽ trắng hoàn toàn, mất toàn bộ dữ liệu dự án |
| `01_Brand` đến `06_System_KPI` | **Nghiêm trọng** — PMO dashboard sẽ mất dữ liệu workstream tương ứng |
| `MEETING_MINUTES` | **Nghiêm trọng** — module Biên bản họp mất dữ liệu, Risk/SOS và CEO Decision mất nguồn phụ |
| `TO_ACTIONS` / `TO_ISSUES` | **Nghiêm trọng** — TO Action Center trắng |
| `CEO_DECISIONS` | **Cao** — CEO Decision Board mất dữ liệu nhập thủ công |
| `RISK_SOS` | **Cao** — Risk/SOS Center mất dữ liệu nhập thủ công |
| `CEO_Comments` | **Trung bình** — Mất lịch sử comment trên dashboard |
| `PMO_ALL_ACTIONS` | **Thấp** — GAS sẽ tạo lại khi sync, nhưng mất data nếu chưa sync |
| `PIC01`-`PIC16` | **Thấp** — GAS tạo lại khi import, nhưng mất data cũ nếu chưa backup |
| `MASTER_STATUS` / `CONFIG` | **Thấp** — GAS hoạt động bình thường, chỉ mất cấu hình màu sắc |
| `00_Dashboard_CEO`, `AI_Review_*` | **Không rõ** — Cần xác nhận trước |

---

## 8. Thống kê

| Chỉ số | Số lượng |
|---|---|
| **Tổng số sheet tìm thấy trong code** | **36** |
| Sheet NHẬP DỮ LIỆU (người dùng cần nhập) | 16 |
| Sheet HỆ THỐNG (GAS tự ghi) | 20 |
| Sheet THAM CHIẾU | 2 |
| Sheet CHƯA XÁC ĐỊNH (không có trong code) | 3 (*) |

> (*) `00_Dashboard_CEO`, `AI_Review_Input`, `AI_Review_Output` được user đề cập nhưng không tìm thấy trong `ma.gs` hay `index.html`. Có thể là sheet cũ hoặc dùng ngoài hệ thống này.

---

## 9. Hàm Apps Script quan trọng cần biết

| Hàm | Chức năng | Khi nào chạy |
|---|---|---|
| `setupCEOProjectTrackerSheet()` | Tạo/cập nhật header sheet THEO_DOI_DU_AN_CEO | Chạy 1 lần khi setup |
| `importCEOProjectSeedData()` | Nhập dữ liệu profile 8 dự án + 48 hạng mục | Chạy 1 lần để seed data |
| `setupGspNext30Database()` | Tạo tất cả sheet DB (TO_ACTIONS, MEETING_MINUTES, v.v.) | Chạy 1 lần khi setup |
| `setupPicSheets()` | Tạo 16 PIC sheets + PMO_ALL_ACTIONS + log sheets | Chạy 1 lần khi setup |
| `syncPicSheetsToPmoAllActions()` | Gộp 16 PIC sheets → PMO_ALL_ACTIONS | Chạy định kỳ (hàng tuần) |
| `importPicPlanToDatabase(url, ...)` | Import kế hoạch từ Google Sheet của PIC | Chạy khi PIC cập nhật file |
| `importToReportToDatabase(url, ...)` | Import TO report → TO_ACTIONS + TO_ISSUES | Chạy khi TO báo cáo |
| `getDashboardData()` | Tổng hợp tất cả dữ liệu PMO → trả về cho dashboard | Chạy mỗi khi refresh dashboard |
| `getCommandCenterData()` | Đọc tất cả sheet DB → trả về cho Command Center | Chạy khi dashboard refresh |

---

*Tài liệu này được tạo tự động bằng cách phân tích ma.gs + index.html — không có thay đổi nào được thực hiện trên dữ liệu hoặc sheet.*
