# PROJECT_MAP (v2.5.0)

## Tổng quan Functional
Ứng dụng **Oracle HC Generator** (trước đây là Oracle Report Generator) là một Desktop App (viết bằng Python/PyQt5) dùng để tự động phân tích và trích xuất dữ liệu từ các file log của hệ thống Oracle Database RAC hỗ trợ đa nút (N-Node, N >= 1).

### Các tính năng chính:
- **Phân tích song song (Parallel Parsing):** Xử lý AWR Reports, Alert Logs và System Information từ nhiều Node cùng lúc.
- **Tích hợp OSWBB & ExaWatcher:** Tự động tạo đồ thị hiệu năng (CPU, Memory, IO) với khả năng fallback thông minh và tùy chỉnh màu sắc biểu đồ.
- **Hệ thống Kéo & Thả (Universal Drag & Drop):** Hỗ trợ kéo thả file/folder vào mọi ô nhập liệu và danh sách.
- **Tùy chọn Database Role (Primary/Standby):** Cung cấp khả năng lọc và rút gọn báo cáo dành riêng cho site Standby, hỗ trợ hậu tố `-STB` trong tên file và tiêu đề báo cáo.
- **Báo cáo chuẩn hóa:** Xuất file Word (`.docx`) với định dạng bảng biểu 16.4cm, tự động map SQL ID.
- **Giao diện Windows 11 Fluent:** Thiết kế hiện đại, tối ưu hóa kích thước (Standard Height 28px), đồng bộ chiều cao giữa các module.
- **Tự động hóa Đóng gói (Packaging Automation):** Tự động tăng phiên bản và đồng bộ hóa tài liệu khi thực hiện build ứng dụng.

## Cấu trúc File/Thư mục hiện tại
```text
NewApplication/
├── main.py                    # Entry point chính của ứng dụng
├── build_onefile.spec         # Cấu hình PyInstaller (Đóng gói JRE, QSS, Icon)
├── PROJECT_MAP.md              # Sơ đồ dự án và tài liệu hệ thống (File này)
├── DATA_FLOW_MAP.md            # Bản đồ luồng dữ liệu chi tiết
├── README.md                  # Hướng dẫn sử dụng cho người dùng cuối
│
├── styles/                    # Giao diện và Asset
│   ├── main.qss               # Stylesheet trung tâm (Đã chuẩn hóa 28px)
│   └── app_icon.ico           # Biểu tượng ứng dụng
│
├── src/                       # Source code chính
│   ├── config.py              # Cấu hình hệ thống (Version 2.3.0)
│   ├── ui/                    # Giao diện đa tab và logic sự kiện (Drag & Drop)
│   ├── generators/            # Engine tạo báo cáo (Fallback image discovery)
│   ├── parsers/               # Bộ phân tích log (Alert, AWR, Patch)
│   ├── models/                # Cơ sở dữ liệu (SQLite stub)
├── scripts/                   # Scripts quản lý và build
│   ├── build_app.bat          # Script thực thi đóng gói
│   └── packager.py            # Engine tự động tăng phiên bản & đóng gói
│
└── venv/                      # Môi trường ảo Python
```

## Thống kê Tham Số Hệ Thống
- **Tên App:** Oracle HC Generator
- **Phiên bản:** v2.5.1
- **Ngôn ngữ:** Python 3.8+ (PyQt5)
- **Thiết kế UI:** Fluent Design (Accent: #0067C0, Height: 28px)
- **Đóng gói:** PyInstaller One-file (JRE mini đi kèm trong dist/)
- **Tự động hóa:** Packager Script (Auto Bump: x.9.0 -> x+1.0.0)

---
*Cập nhật lần cuối: 17/04/2026*
