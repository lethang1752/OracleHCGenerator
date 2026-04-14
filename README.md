# 🚀 Oracle HC Generator - Hướng Dẫn Sử Dụng

[![Version](https://img.shields.io/badge/version-2.1.1-blue.svg)](https://github.com/lethang1752/OracleHCGenerator)
[![Platform](https://img.shields.io/badge/platform-Windows-0078D4.svg)](https://www.microsoft.com/windows/)
[![UI](https://img.shields.io/badge/UI-Fluent_Design-0067C0.svg)](https://docs.microsoft.com/en-us/windows/apps/design/)

**Oracle HC Generator** là công cụ Desktop chuyên dụng giúp tự động hóa việc tạo báo cáo sức khỏe cơ sở dữ liệu (Health Check) cho hệ thống Oracle RAC. Ứng dụng này được thiết kế theo phong cách hiện đại, trực quan và không cần cài đặt (Portable).

---

## 📖 Mục Lục
1. [Chuẩn Bị Dữ Liệu](#1-chuẩn-bị-dữ-liệu)
2. [Tạo Báo Cáo Phụ Lục (Appendix Generator)](#2-tạo-báo-cáo-phụ-lục)
3. [Xử Lý Đồ Thị Hiệu Năng (OSWBB Graph)](#3-xử-lý-đồ-thị-hiệu-năng)
4. [Gộp Tài Liệu (Merge Documents)](#4-gộp-tài-liệu)
5. [Bộ Công Cụ Thu Thập (Collection Tools)](#5-bộ-công-cụ-thu-thập)
6. [Các Lưu Ý Quan Trọng](#6-các-lưu-ý-quan-trọng)

---

## 1. Chuẩn Bị Dữ Liệu
Để ứng dụng có thể parse dữ liệu chính xác, bạn cần tổ chức các thư mục của từng Node như sau:
- **Alert Log:** Tệp `.xml` hoặc `.log` (Ví dụ: `alert_prod1.log`).
- **AWR Reports:** Các tệp báo cáo AWR định dạng `.html`.
- **Database Info:** Tệp `database_information.html` (chi tiết thông tin cấu hình DB).

> [!TIP]
> Hãy gom tất cả dữ liệu của **Node 1** vào một thư mục, **Node 2** vào một thư mục riêng biệt để dễ dàng quản lý.

---

## 2. Tạo Báo Cáo Phụ Lục (Appendix Generator)
Đây là tính năng cốt lõi của ứng dụng.
1. **Thêm thư mục Node:** Nhấn nút **➕ Add Node Folder** để chọn thư mục dữ liệu của từng Node (hỗ trợ tối đa 8 Nodes).
2. **Cấu hình thời gian:** Nhập số ngày (`Alert Log Scan Days`) bạn muốn quét lỗi ORA-. Thông thường là 10-15 ngày.
3. **Cấu hình báo cáo:**
   - Chọn phông chữ (`Font Choice`): Times New Roman (truyền thống) hoặc Segoe UI (hiện đại).
   - Đặt tên tệp đầu ra.
4. **Thực thi:** Nhấn **🚀 GENERATE COMPREHENSIVE REPORT**. Ứng dụng sẽ parse dữ liệu song song và thông báo khi hoàn tất.

---

## 3. Xử Lý Đồ Thị Hiệu Năng (OSWBB Graph)
Tính năng này giúp trực quan hóa dữ liệu từ công cụ OSWatcher.
1. **Input Directory:** Chọn thư mục chứa các tệp nén `.tar.bz2` hoặc thư mục đã giải nén của OSW.
2. **Output Directory:** Nơi lưu trữ các đồ thị sinh ra.
3. **Sync & Push (Nâng cao):** 
   - Bạn có thể thêm các thư mục đích vào danh sách **Target Folders**.
   - Sau khi tạo đồ thị xong, ứng dụng sẽ tự động copy kết quả sang các thư mục này.
4. **Thực thi:** Nhấn **🚀 GENERATE & PUSH**.

---

## 4. Gộp Tài Liệu (Merge Documents)
Dùng để ghép nhiều tệp Word (.docx) thành một tệp duy nhất.
- **Thao tác:** Kéo và thả (Drag & Drop) các tệp tin vào danh sách.
- **Sắp xếp:** Sử dụng nút **▲ Move Up** / **▼ Move Down** để thay đổi thứ tự các phần trong báo cáo.
- **Thực thi:** Nhấn **🔗 MERGE DOCUMENTS**.

---

## 5. Bộ Công Cụ Thu Thập (Collection Tools)
Ứng dụng tích hợp sẵn các script thu thập dữ liệu tiêu chuẩn.
- Nhấn vào tab **📦 Collection Tools** ở dưới cùng sidebar.
- Tại đây liệt kê các script SQL/Shell. Bạn có thể nhấn **📂 Open Folder** để lấy script và chạy trên Server Oracle.
- Nếu danh sách trống khi vừa mở app, nhấn **🔄 Refresh List** để cập nhật.

---

## 6. Các Lưu Ý Quan Trọng

> [!IMPORTANT]
> **Quyền truy cập tệp tin:** Đảm bảo bạn không mở tệp DOCX kết quả bằng Word trong khi ứng dụng đang thực hiện ghi đè, nếu không sẽ gặp lỗi "Permission Denied".

> [!NOTE]
> **Hỗ trợ Java:** Ứng dụng này đã được đóng gói kèm một bộ chạy Java thu nhỏ (JRE 8) bên trong thư mục `jre_mini`. Bạn **KHÔNG CẦN** cài đặt thêm Java trên máy tính cá nhân để chạy tính năng đồ thị OSWBB.

---
**Tác giả:** Victor Le
*Cập nhật lần cuối: 14/04/2026*
