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

## 10. Meeting Hub V3 — Kiến trúc dữ liệu thật

> Cập nhật: 2026-10-06 · Thay thế cách đọc legacy (MEETING_TOPICS / MEETING_MINUTES / MEETING_ACTIONS)
> Config qua `PropertiesService.getScriptProperties()` — **KHÔNG hardcode ID**.

### 10.1 Script Properties cần cấu hình

| Key | Giá trị cần điền | Bắt buộc |
|---|---|---|
| `MEETING_HUB_SPREADSHEET_ID` | ID Google Spreadsheet chứa 5 sheets mới | ✅ |
| `DRIVE_ROOT_FOLDER_ID` | ID folder gốc trên Google Drive cho Meeting Hub | ✅ |
| `MINUTES_TEMPLATE_FILE_ID` | ID file Google Doc template biên bản | ⚠️ Tuỳ chọn |
| `ALLOWED_DOMAIN` | `goldsunpackaging.vn` | ✅ (default có sẵn) |

### 10.2 Sheets mới (tạo bằng `setupMeetingHubDataStore()`)

| Sheet | Nhóm | Số cột | Hàm đọc chính | Hàm ghi chính | Ghi chú |
|---|---|---|---|---|---|
| `MEETINGS` | NHẬP | 29 | `getMeetingHubDataFromSheets_()` | `createMeeting()` | Chứa driveFolderId, minutesFileId, v.v. |
| `ACTIONS` | NHẬP | 25 | `getMeetingHubDataFromSheets_()` | `updateAction()`, `importMeetingActionsV3()` | ACTION_LOG tự ghi khi update |
| `PEOPLE` | THAM CHIẾU | 5 | `getPersonDetail()` | `seedMeetingHubPeopleFromArray()` | email, fullName, department, title, active |
| `TOPICS` | THAM CHIẾU | 6 | `getTopicDetail()` | — (nhập thủ công) | topicId, topicName, description, module, active, sortOrder |
| `ACTION_LOG` | HỆ THỐNG | 10 | — (audit only) | `updateAction()`, `importMeetingActionsV3()` | Ghi tự động khi Status/Deadline/Owner thay đổi |

#### Schema MEETINGS (29 cột)
`meetingId` · `meetingDate` · `meetingTitle` · `topicId` · `topicName` · `meetingType` · `chairperson` · `secretary` · `participants` · `summary` · `mainContent` · `analysis` · `conclusions` · `decisions` · `ceoDirections` · `needCeoDecision` · `driveFolderId` · `driveFolderUrl` · `minutesFolderId` · `attachmentsFolderId` · `evidenceFolderId` · `otherFolderId` · `minutesFileId` · `minutesFileName` · `minutesFileType` · `minutesFileUrl` · `status` · `createdAt` · `updatedAt`

#### Schema ACTIONS (25 cột)
`actionId` · `meetingId` · `topicId` · `actionTitle` · `actionDescription` · `ownerEmail` · `ownerName` · `coOwnerEmails` · `department` · `deadline` · `expectedOutput` · `priority` · `status` · `rag` · `progress` · `actualResult` · `evidenceFileId` · `evidenceFileName` · `evidenceFileUrl` · `blocker` · `needCeoDecision` · `ceoDecision` · `lastUpdate` · `createdAt` · `updatedAt`

#### Schema ACTION_LOG (10 cột)
`logId` · `actionId` · `meetingId` · `updatedAt` · `updatedBy` · `oldStatus` · `newStatus` · `oldDeadline` · `newDeadline` · `progress` · `note`

### 10.3 Cấu trúc Google Drive

```
DRIVE_ROOT_FOLDER_ID/
└── YYYY/
    └── MM/
        └── MTG-YYYYMMDD-XXX_TenCuocHop/   ← driveFolderId trong MEETINGS
            ├── 01_Bien_ban/                 ← minutesFolderId
            ├── 02_Tai_lieu_hop/             ← attachmentsFolderId
            ├── 03_Bang_chung_Action/        ← evidenceFolderId
            └── 04_Khac/                     ← otherFolderId
```

- Meeting ID format: `MTG-YYYYMMDD-XXX` (XXX = số thứ tự trong ngày, 3 chữ số)
- Action ID format: `ACT-YYYYMMDD-XXXX` (XXXX = số thứ tự trong ngày, 4 chữ số)
- File preview URL: Google Docs → `/document/d/{ID}/preview` · Sheets → `/spreadsheets/d/{ID}/preview` · Khác → `/file/d/{ID}/preview`
- Trash file: `file.setTrashed(true)` — **KHÔNG bao giờ xóa vĩnh viễn**

### 10.4 Hàm Apps Script quan trọng — Meeting Hub V3

| Hàm | Chức năng | Khi nào chạy |
|---|---|---|
| `setupMeetingHubDataStore()` | Tạo/chuẩn hóa 5 sheets với headers — không xóa data cũ | 1 lần khi setup |
| `setupMeetingHubConfigExample()` | In hướng dẫn cấu hình Script Properties vào log | Khi cần hỗ trợ |
| `testMeetingHubConnection()` | Diagnostic: config, SS, Drive, row counts | Sau mỗi lần cấu hình |
| `getMeetingHubData()` | V3 main — đọc thật, fallback demo nếu chưa cấu hình | Mỗi khi dashboard load |
| `createMeeting(params)` | Tạo meeting + Drive folder + ghi MEETINGS | Khi tạo cuộc họp mới |
| `updateAction(actionId, changes)` | Cập nhật action + ghi ACTION_LOG tự động | Khi PIC cập nhật tiến độ |
| `importMeetingActionsV3(rows, mode, meetingId)` | Import Excel → ACTIONS + validate + ACTION_LOG | Sau cuộc họp |
| `seedMeetingHubPeopleFromArray(arr)` | Thêm bulk PEOPLE (skip duplicate email) | Setup ban đầu |
| `migrateLegacyMeetingHubData(dryRun)` | Migrate dữ liệu cũ → 5 sheets mới | 1 lần, sau setup |
| `createMeetingDriveFolder(meetingId, title)` | Tạo cấu trúc 4 subfolder trên Drive | Auto-call bởi createMeeting |
| `createMinutesFromTemplate(meetingId, title)` | Copy template → 01_Bien_ban, replace {{TOKEN}} | Khi cần biên bản |

### 10.5 Quan hệ fallback

```
getMeetingHubData()
  ├── validateMeetingHubConfig() → valid?
  │   ├── YES → getMeetingHubDataFromSheets_() [REAL DATA]
  │   └── NO  → getMeetingHubDataLegacy_()     [DEMO FALLBACK, _isDemo:true]
  │              ↑ đọc MEETING_TOPICS + MEETING_MINUTES + MEETING_ACTIONS (sheets cũ)
  └── Frontend nhận _isDemo:true → hiển thị banner "Demo mode"
```

---

## 11. LOS — LEADERSHIP OPERATING SYSTEM

> Cập nhật: 2026-10-06 · Module mới — theo dõi năng lực điều hành cán bộ quản lý
> 2 chỉ số độc lập: **LOS** (Leadership Operating System) và **ME** (Meeting Effectiveness)
> KHÔNG tạo Overall Score. KHÔNG cộng LOS + ME.

### 11.1 Phạm vi áp dụng

Cán bộ từ **Trưởng bộ phận trở lên** — lọc bằng `isLosEligibleTitle_()`.
TUYỆT ĐỐI không tự tạo Leader giả. Nếu chưa có dữ liệu → Empty State.

### 11.2 Sheets mới (tạo bằng `setupLosDataStore()`)

| Sheet | Số cột | Mô tả |
|---|---|---|
| `LOS_LEADER` | 18 | Profile Leader đã enroll vào chương trình LOS |
| `LOS_SCORE` | 27 | LOS score theo 8 tiêu chí, mỗi tiêu chí có AI Proposed + Verified |
| `LOS_MEETING` | 8 | Reference tới Meeting Hub (KHÔNG duplicate) — chỉ lưu meetingId + leaderRole |
| `LOS_ME_SCORE` | 24 | ME score theo 6 tiêu chí, có ME Cap ≤ 5 nếu thiếu Decision+Action |
| `LOS_COACHING` | 20 | Coaching Actions theo từng Gap |
| `LOS_EVIDENCE` | 11 | Reference tới Meeting/Action/OC/Drive (chỉ lưu sourceId, không copy binary) |
| `LOS_REVIEW` | 17 | Gate Review: Baseline → D7 → D14 → D30 → D60 → D90 |

### 11.3 LOS Score — 8 Tiêu chí (tổng weight = 10)

| Tiêu chí | Weight |
|---|---|
| Data-driven | 1.5 |
| Planning & Prioritization | 1.5 |
| Decision Making | 1.5 |
| Ownership | 1.5 |
| Delegation/PIC | 1.0 |
| Problem Solving | 1.0 |
| Follow-through | 1.0 |
| Coaching/Team Development | 1.0 |

### 11.4 ME Score — 6 Tiêu chí (tổng weight = 10)

| Tiêu chí | Weight | Rule đặc biệt |
|---|---|---|
| Data Ready | 1.5 | |
| Problem/Gap rõ | 1.5 | |
| Root Cause | 1.5 | |
| Decision | 2.0 | **Cap trigger** |
| Action – PIC – Deadline | 2.0 | **Cap trigger** |
| Follow-up / Closure | 1.5 | |

**ME Cap**: Nếu Decision < 1.0 HOẶC Action-PIC-Deadline < 1.0 → ME final ≤ 5.0

### 11.5 Luồng dữ liệu LOS

```
LOS_LEADER (profile)
  └── LOS_SCORE (8 tiêu chí × gate)
  └── LOS_MEETING → MEETINGS.meetingId (reference Meeting Hub)
       └── LOS_ME_SCORE (6 tiêu chí × meeting)
  └── LOS_COACHING (Gap → Action → Before/After)
  └── LOS_EVIDENCE → Meeting/Action/OC/Drive sourceId (reference only)
  └── LOS_REVIEW (Gate snapshot: Baseline/D7/D14/D30/D60/D90)
```

### 11.6 AI vs Verified Score

- `aiProposedScore`: AI đề xuất — hiển thị badge "AI đề xuất"
- `verifiedScore`: Score chính thức — chỉ human mới set
- AI KHÔNG được tự override verifiedScore
- Score hiển thị trên Dashboard: verifiedScore nếu có, sinon aiProposedScore + badge

### 11.7 Tích hợp Meeting Hub

```
LOS_MEETING.meetingId  →  MEETINGS.meetingId
LOS_EVIDENCE.sourceId  →  MEETINGS.meetingId / ACTIONS.actionId / ...
Click Meeting trong LOS  →  openMeetingHub()  →  mhOpenMeeting(meetingId)
```

### 11.8 Hàm Apps Script quan trọng — LOS

| Hàm | Chức năng | Khi nào chạy |
|---|---|---|
| `setupLosDataStore()` | Tạo 7 sheets LOS với headers | 1 lần khi setup |
| `testLosConnection()` | Diagnostic: config, sheets, user | Sau mỗi lần cấu hình |
| `getLosData()` | Tất cả Leader + scores + coaching + gates | Mỗi khi module LOS load |
| `getLosEligibleLeaders()` | Lấy danh sách Leader eligible (từ LOS_LEADER hoặc PEOPLE) | Sub-function |
| `calculateLosScore_(obj)` | Tính LOS /10 từ 8 tiêu chí | Helper |
| `calculateMeScore_(obj)` | Tính ME /10 từ 6 tiêu chí + ME Cap | Helper |
| `applyMeDecisionActionCap_(raw, obj)` | Áp ME Cap ≤ 5 nếu thiếu Decision/Action | Helper |
| `saveLosVerifiedScore(payload)` | Ghi Verified Score vào LOS_SCORE | Khi PMO/CEO xác nhận điểm |
| `saveMeVerifiedScore(payload)` | Ghi ME Verified Score vào LOS_ME_SCORE | Khi PMO xác nhận ME |
| `createLosCoachingAction(payload)` | Tạo Coaching Action mới | Sau mỗi Review |
| `saveLosReview(payload)` | Lưu Gate Review snapshot | Tại Baseline/D30/D60/D90 |

### 11.9 Script Properties (dùng chung với Meeting Hub)

| Key | Mô tả |
|---|---|
| `MEETING_HUB_SPREADSHEET_ID` | ID Spreadsheet chứa cả MH sheets và LOS sheets |
| `ALLOWED_DOMAIN` | goldsunpackaging.vn |

---

## 12. CLC — CHI PHÍ LỖI CHẤT LƯỢNG

### 12.1 Tổng quan module

Module CLC quản lý vòng đời PDCA cho chi phí lỗi chất lượng:  
**Data → Case → Root Cause → Action (CA/PA) → PIC Commitment → Weekly Review → Monthly KPI**

### 12.2 Sheet: `CLC_COMMITMENT`

Lưu cam kết của từng PIC gắn với Case/Action cụ thể. Hỗ trợ xác nhận 2 cấp.

| Cột | Kiểu | Mô tả |
|---|---|---|
| `commitmentId` | Text (PK) | Mã cam kết, ví dụ: `C-001` |
| `caseId` | Text (FK) | Mã Case CLC liên quan |
| `actionId` | Text (FK) | Mã Action (CA/PA) liên quan |
| `picId` | Text (FK) | Mã PIC chịu trách nhiệm |
| `picName` | Text | Tên đầy đủ PIC |
| `department` | Text | Bộ phận/Phòng ban |
| `commitmentText` | Text (long) | Nội dung cam kết (ngôn ngữ tự nhiên) |
| `commitmentDate` | Date | Ngày PIC đặt cam kết |
| `deadline` | Date | Deadline hiện tại (có thể carry-forward) |
| `originalDeadline` | Date | Deadline gốc ban đầu (KHÔNG thay đổi khi carry-forward) |
| `expectedOutput` | Text | Đầu ra bắt buộc (kết quả, bằng chứng cụ thể) |
| `status` | Text | Trạng thái: Chờ PIC xác nhận / Đang thực hiện / Quá hạn / Hoàn thành chờ xác minh / Đã xác minh |
| `signed` | Boolean | TRUE = Cấp 1 đã xác nhận điện tử |
| `signedBy` | Text | Email/Tên người ký |
| `signedDate` | Date | Ngày ký xác nhận |
| `signatureEvidenceUrl` | URL | Link file biên bản ký (Cấp 2 — tùy chọn) |
| `actualResult` | Text | Kết quả thực tế PIC tự cập nhật |
| `evidenceUrl` | URL | Link bằng chứng hoàn thành (ảnh/file) |
| `verifiedBy` | Text | Người xác minh (QA/Manager) |
| `verifiedDate` | Date | Ngày xác minh |
| `carryForwardFrom` | Text | commitmentId gốc nếu là carry-forward |
| `overdueDays` | Number | Số ngày quá hạn (tính từ originalDeadline) |
| `_savedAt` | Timestamp | Tự động — GAS ghi khi save/upsert |
| `_savedBy` | Email | Tự động — email người gọi GAS |
| `_confirmedAt` | Timestamp | Tự động — thời điểm xác nhận điện tử |

**Hàm GAS:** `saveClcCommitment(payload)`, `confirmClcCommitment(payload)`, `getClcCommitmentsData_(ss)`  
**doPost actions:** `saveClcCommitment`, `confirmClcCommitment`  
**doGet actions:** `getClcDashboard` (trả về mảng commitments)

---

### 12.3 Sheet: `CLC_WEEKLY_PIC_REVIEW`

Snapshot đánh giá từng PIC mỗi cuối tuần. Dùng cho Weekly PIC Review và lịch sử KPI.

| Cột | Kiểu | Mô tả |
|---|---|---|
| `reviewId` | Text (PK) | Mã đánh giá, ví dụ: `WR-P01-W4` |
| `picId` | Text (FK) | Mã PIC |
| `year` | Number | Năm, ví dụ: 2026 |
| `month` | Number | Tháng (1–12) |
| `week` | Number | Tuần trong tháng (1–4) |
| `kpiMonthlyTarget` | Number | Mục tiêu KPI tổng hợp của tháng (0–100) |
| `weekScore_k1` | Number | Điểm thực tế KPI1 — Giảm phát sinh CLC mới (%) |
| `weekScore_k2` | Number | Điểm thực tế KPI2 — Xử lý CLC hiện hữu (%) |
| `weekScore_k3` | Number | Điểm thực tế KPI3 — Hoàn thành Action đúng hạn (%) |
| `weekScore_k4` | Number | Điểm thực tế KPI4 — Phòng ngừa/chống tái diễn (%) |
| `weekScore_k5` | Number | Điểm thực tế KPI5 — Kỷ luật dữ liệu & báo cáo (%) |
| `weeklyScore` | Number | Điểm tổng hợp có trọng số (0–100) |
| `mtdActual` | Number | Điểm MTD trung bình tích lũy (0–100) |
| `commitmentStatus` | Text | Tổng trạng thái cam kết tuần: Đang thực hiện / Quá hạn / Đã xác minh / Chờ PIC xác nhận |
| `overdueActions` | Number | Số Action quá hạn tính đến cuối tuần |
| `repeatIssues` | Number | Số lỗi tái diễn phát hiện trong tuần |
| `reviewNote` | Text (long) | Nhận xét CEO/Manager sau buổi đánh giá |
| `nextWeekCommitment` | Text (long) | Cam kết hành động tuần tới |
| `reviewedBy` | Text | Người đánh giá (thường là CEO/Manager) |
| `snapshotDate` | Date | Ngày chụp snapshot đánh giá |
| `_savedAt` | Timestamp | Tự động — GAS ghi |
| `_savedBy` | Email | Tự động — email người gọi |

**Hàm GAS:** `saveClcWeeklyReview(payload)`, `getClcWeeklyReviewData_(ss)`  
**doPost actions:** `saveClcWeeklyReview`  
**doGet actions:** `getClcDashboard` (trả về mảng weeklyReviews)

---

### 12.4 Cài đặt ban đầu

1. Chạy `setupClcDataStore()` từ GAS Editor để tạo 2 sheet với headers đúng chuẩn.
2. Hai sheet được tạo trong Spreadsheet hiện tại (hoặc ID trong `CLC_SPREADSHEET_ID` nếu cấu hình riêng).
3. Module hoạt động ở **Demo Mode** ngay lập tức — không cần dữ liệu thật để test.

### 12.5 Weights KPI (có thể điều chỉnh)

| KPI | Trọng số mặc định | Target mặc định |
|---|---|---|
| K1 — Giảm phát sinh CLC mới | 30% | 30% |
| K2 — Xử lý CLC hiện hữu | 20% | 20% |
| K3 — Hoàn thành Action đúng hạn | 20% | 85% |
| K4 — Phòng ngừa / chống tái diễn | 20% | 80% |
| K5 — Kỷ luật dữ liệu & báo cáo | 10% | 90% |

> Trọng số chỉ chính thức sau khi CEO phê duyệt. Điều chỉnh trong `CLC_KPI_CONFIG` ở frontend JS.

---

*Tài liệu này được tạo tự động bằng cách phân tích ma.gs + index.html — không có thay đổi nào được thực hiện trên dữ liệu hoặc sheet.*
