# Oracle HC Generator (v2.5.0)

[![Release](https://img.shields.io/badge/Release-v2.5.0-blue.svg)](https://github.com/lethang1752/OracleHCGenerator/releases)
[![Python](https://img.shields.io/badge/Python-3.8+-green.svg)](https://www.python.org/)

**Oracle HC Generator** là giải pháp Desktop mạnh mẽ giúp tự động hóa quá trình thu thập, phân tích và tạo báo cáo kiểm tra sức khỏe (Health Check) cho hệ thống Oracle Database RAC. Ứng dụng được thiết kế theo phong cách hiện đại, trực quan và không cần cài đặt (Portable).

---

## Các Tính Năng Nổi Bật

- **Phân tích song song (N-Node):** Xử lý đồng thời AWR và Alert Log từ tất cả các Node trong Cluster.
- **Tích hợp OSWBB & ExaWatcher:** Tự động tạo đồ thị hiệu năng hệ điều hành và hỗ trợ nhận diện ảnh thông minh cho cả hệ thống thông thường và Exadata.
- **Tùy chọn Standby Database:** Tính năng rút gọn báo cáo dành riêng cho site Standby, tự động thêm hậu tố `-STB`.
- **Hệ thống Kéo & Thả (Universal Drag & Drop):** Kéo thả file/folder thực tế từ Windows Explorer vào ứng dụng để nhập liệu cực nhanh.
- **Báo cáo chuyên nghiệp:** Xuất file Word định dạng chuẩn, bảng biểu cố định (16.4cm), tự động map thông tin SQL ID.
- **Tự động hóa Đóng gói:** Quy trình build tự động hóa việc tăng phiên bản và đồng bộ tài liệu hệ thống.
- **Giao diện Windows 11 Fluent:** Hiện đại, tối ưu hóa chiều cao input 28px giúp giao diện gọn gàng.

## Hướng Dẫn Sử Dụng Nhanh

1. **Appendix Generator:** Thêm các thư mục log của từng Node -> Chọn Font -> Nhấn Generate.
2. **OSWBB Graph Generator:** Chọn thư mục chứa log OSWBB -> Chọn thư mục lưu ảnh -> Nhấn Generate.
3. **Merge Documents:** Thêm các file Word lẻ -> Sắp xếp thứ tự (hoặc Auto Sort theo DB) -> Nhấn Merge.

---

## Dành Cho Nhà Phát Triển

Nếu bạn muốn chạy ứng dụng từ mã nguồn:

```bash
# 1. Tạo môi trường ảo
python -m venv venv

# 2. Kích hoạt môi trường ảo (Windows)
.\venv\Scripts\activate

# 3. Cài đặt thư viện
pip install -r requirements.txt

# 4. Chạy ứng dụng
python main.py

# 5. Đóng gói ứng dụng thành file .exe duy nhất
pyinstaller build_onefile.spec
```

---
**Tác giả:** Victor Le  
*Cập nhật lần cuối: 15/04/2026*
