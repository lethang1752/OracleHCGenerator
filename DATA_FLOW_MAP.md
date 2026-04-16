# Bản Đồ Luồng Dữ Liệu - Oracle HC Generator
> *Tài liệu kỹ thuật nội bộ - Dùng để tham chiếu khi phát triển tính năng mới*

---

## TỔNG QUAN KIẾN TRÚC

```
📁 Thư mục Node (log_dir)
    │
    ├── alert_*.log               ──► AlertLogParser      ──► alerts_data (dict)
    ├── awrrpt_*.html             ──► AWRParser           ──► awr_data    (dict)
    └── database_information.html ──► DatabaseInfoParser  ──► db_info     (dict)
                                                │
                                   ParseWorker (ProcessPoolExecutor)
                                   [Chạy song song cho mỗi Node]
                                                │
                                    ┌───────────▼──────────────┐
                                    │      parsed_data (dict)   │
                                    │  ├─ db_name: str          │
                                    │  └─ nodes: List[dict]     │
                                    └───────────┬──────────────┘
                                                │
                              ComprehensiveHealthcareReportGenerator
                                                │
                                          📄 output.docx
```

---

## PARSER 1: AlertLogParser
**File nguồn:** `alert_*.log` — glob pattern: `ALERT_LOG_PATTERN = "alert_*.log"`
**Class:** `src/parsers/alert_parser.py`

### Đầu ra: `get_data()` → `dict`
```python
{
    "alerts": [
        {
            "timestamp":  "2025-01-01T10:30:45",  # str, ISO format (chuẩn hóa từ cả 11g và 12c)
            "message":    "ORA-00600: ...",         # str, dòng ORA- đầu tiên trong block
            "error_code": "ORA-00600",              # str, mã lỗi trích xuất bằng regex
            "full_text":  "ORA-00600: ...\n..."     # str, toàn bộ block lỗi (nhiều dòng)
        },
        # ...
    ],
    "count":         15,       # int, tổng số lỗi ORA- tìm thấy
    "num_days":      30,       # int, khoảng thời gian đã quét (ngày)
    "instance_name": "PROD1"   # str, lấy từ tên file: alert_PROD1.log
}
```

### Quy tắc parse:
- Chỉ lấy entri có chứa dòng bắt đầu bằng `ORA-`
- Lọc thời gian: từ `(ngày sửa file - num_days ngày)` trở đi
- Hỗ trợ 2 format timestamp:
  - **Oracle 12c:** `YYYY-MM-DDTHH:MM:SS` (32 ký tự)
  - **Oracle 11g:** `DDD MMM DD HH:MM:SS YYYY` (24 ký tự)

---

## PARSER 2: AWRParser
**File nguồn:** `awrrpt_*.html` — glob: `AWR_REPORT_PATTERN = "awrrpt_*.html"`
**Class:** `src/parsers/awr_parser.py`

### Đầu ra: `get_data()` → `dict`
```python
{
    "tables": [
        {
            "title":     "This table displays...",  # str, thuộc tính `summary` của thẻ <table>
            "rows":      [["Col1", "Col2"], ...],   # List[List[str]], hàng đầu là header
            "row_count": 15                         # int
        },
        # ...
    ],
    "table_count":   8,
    "db_name":       "PRODDB",     # str, lấy từ rows[1][0] của bảng đầu tiên
    "instance_name": "PRODDB1",    # str, lấy từ rows[1][1]
    "db_version":    "19.0.0.0.0"  # str, lấy từ rows[1][2]
}
```

### Danh sách bảng AWR được lọc (config.py → `AWR_TABLES_12C` / `AWR_TABLES_11G`):

| Keyword tìm trong `table.summary`                              | Dùng trong section |
|----------------------------------------------------------------|--------------------|
| `"This table displays host information"`                       | Metadata (db_name) |
| `"This table displays instance efficiency percentages"`        | 1.3.2 Buffer Ratio |
| `"This table displays wait class statistics..."`               | 1.3.3 Wait Classes |
| `"This table displays top SQL by elapsed time"`                | 1.3.4 Top Queries  |
| `"This table displays the text of the SQL statements"`         | 1.3.5 SQL Text     |
| `"Top 10 Foreground Events by Total Wait Time"`                | 1.3.6 Wait Events  |
| `"Top 5 Timed Foreground Events"` *(11g)*                     | 1.3.6 Wait Events  |

---

## PARSER 3: DatabaseInfoParser
**File nguồn:** `database_information.html`
**Class:** `src/parsers/database_info_parser.py`

### Đầu ra: `get_all_data()` → `dict`
```python
{
    # Key (str)              : Giá trị (List[List[str]])   - Ghi chú
    "ASM"                    : [[header], [row], ...],  # Thông tin disk group
    "TABLESPACE"             : [[header], [row], ...],  # Sử dụng tablespace
    "INDEX_FRAGMENT"         : [[header], [row], ...],  # Index fragment thông thường
    "INDEX_PARTITION_FRAGMENT": [[header], [row], ...], # Index fragment phân vùng
    "TABLE_FRAGMENT"         : [[header], [row], ...],  # Table fragment thông thường
    "TABLE_PARTITION_FRAGMENT": [[header], [row], ...], # Table fragment phân vùng
    "INVALID_OBJECT"         : [[header], [row], ...],  # Object không hợp lệ
    "TABLE_STATISTICS"       : [[header], [row], ...],  # Thống kê bảng
    "INDEX_STATISTICS"       : [[header], [row], ...],  # Thống kê index
    "CHECK_CLUSTER"          : [[header], [row], ...],  # Trạng thái cluster CRS
    "RESOURCE_CRS"           : [[header], [row], ...],  # Chi tiết tài nguyên CRS
    "DISK_USAGE"             : [[header], [row], ...],  # Dung lượng ổ đĩa theo Node
    "CHECK_BACKUP"           : [[header], [row], ...],  # Trạng thái backup gần nhất
    "BACKUP_POLICY"          : [[header], [row], ...],  # Chính sách RMAN backup
    "DBA_ROLE"               : [[header], [row], ...],  # User có quyền DBA
    "OBJECT_IN_SYSTEM"       : [[header], [row], ...],  # Object trong SYSTEM / SYSAUX
    "CHECK_PATCHES"          : [[header], [row], ...],  # Thông tin patches đã cài
}
```

**Cơ chế parse:** Tìm thẻ HTML (`<b>`, `<p>`, `<td>`, `<span>`...) có nội dung bắt đầu bằng `+`
(ví dụ: `+ASM`, `+TABLESPACE`), sau đó lấy thẻ `<table>` liền kề tiếp theo.

---

## CẤU TRÚC DỮ LIỆU TRUNG GIAN (`parsed_data`)

Sau khi tất cả parsers hoàn tất, `ParseWorker` tổng hợp thành:

```python
parsed_data = {
    "db_name": "PRODDB",    # str, từ awr.db_name của node đầu tiên có giá trị hợp lệ

    "nodes": [
        {
            "node_id"      : 1,
            "data_dir"     : "D:/path/to/node1/",  # str, đường dẫn thư mục gốc của Node
            "instance_name": "PRODDB1",             # str, ưu tiên từ alert_parser, fallback sang awr
            "alerts"       : { ...alerts_data... }, # dict - kết quả của AlertLogParser.get_data()
            "awr"          : { ...awr_data... },    # dict - kết quả của AWRParser.get_data()
            "database_info": { ...db_info... }      # dict - kết quả của DatabaseInfoParser.get_all_data()
        },
        {
            "node_id": 2,
            # ... tương tự với Node 2, 3, ...
        }
    ]
}
```

---

## MAPPING VÀO DOCX

### Bảng tra cứu nhanh:

| Section DOCX              | Nguồn dữ liệu             | Key / Cách truy cập                                                        |
|---------------------------|---------------------------|----------------------------------------------------------------------------|
| **1. Tiêu đề DB**         | `nodes[0]`                | `instance_name` (loại bỏ số cuối bằng vòng lặp `while db_name[-1].isdigit()`) |
| **1.1 Status Check**      | ⚠️ HARDCODE               | `"Running"`, `"Open"`, `"Online"`, `"Running"`                             |
| **1.2 Alert Log**         | `node['alerts']`          | `alerts['alerts'][i]['timestamp']` + `['full_text']` — Lặp qua mọi node   |
| **1.3.1 CPU**             | File ảnh + disk           | `data_dir/generated_files/OSWg_OS_Cpu_Idle.jpg`, `_System.jpg`, `_User.jpg` |
| **1.3.2 Memory**          | File ảnh + AWR            | `OSWg_OS_Memory_Free.jpg`, `_Swap.jpg` + table `"instance efficiency"`     |
| **1.3.3 I/O**             | File ảnh + AWR            | `OSWg_OS_IO_PB.jpg` + table keyword `"wait class"`                         |
| **1.3.4 Top Queries**     | `node['awr']`             | Table keyword `"top sql elapsed"`, bỏ cột `"SQL Module"` (`drop_cols`)     |
| **1.3.5 SQL Text**        | `node['awr']`             | SQL ID từ `row[6]` của Top SQL → lookup trong table `"text of the sql"`    |
| **1.3.6 Wait Events**     | `node['awr']`             | Table keyword `"Top 10 Foreground Events by Total Wait Time"`               |
| **1.3.7 Disk Group**      | `nodes[0]['database_info']` | Key `"ASM"`                                                              |
| **1.3.8 Tablespace**      | `nodes[0]['database_info']` | Key `"TABLESPACE"`                                                       |
| **1.3.9 Index Fragment**  | `nodes[0]['database_info']` | Keys `"INDEX_FRAGMENT"`, `"INDEX_PARTITION_FRAGMENT"`                     |
| **1.3.10 Table Fragment** | `nodes[0]['database_info']` | Keys `"TABLE_FRAGMENT"`, `"TABLE_PARTITION_FRAGMENT"`                     |
| **1.4.1 Invalid Objects** | `nodes[0]['database_info']` | Key `"INVALID_OBJECT"` (filter_nulls=True)                               |
| **1.4.2 Table Stats**     | `nodes[0]['database_info']` | Key `"TABLE_STATISTICS"` (tìm ngày ở cột có 'ANALYZED'/'DATE')           |
| **1.4.2 Index Stats**     | `nodes[0]['database_info']` | Key `"INDEX_STATISTICS"`                                                 |
| **1.5.1 Cluster Status**  | **Mọi node** (gộp)        | `node['database_info']['CHECK_CLUSTER']` — Bỏ header từ node 2 trở đi    |
| **1.5.2 CRS Resources**   | Chỉ `nodes[0]`            | `nodes[0]['database_info']['RESOURCE_CRS']`                                |
| **1.6 Storage**           | Từng node riêng           | `node['database_info']['DISK_USAGE']` — Mỗi node tạo một bảng riêng      |
| **1.7.1 Backup Status**   | `nodes[0]['database_info']` | Key `"CHECK_BACKUP"` (lọc bỏ hàng có cột STATUS = `"NULL"`)             |
| **1.7.2 Scheduling**      | ⚠️ HARDCODE               | Text cố định (backup schedule)                                             |
| **1.7.3 Policy**          | `nodes[0]['database_info']` | Key `"BACKUP_POLICY"`, lấy cell có chứa `"configure"` trong row[1]      |
| **1.8 Dataguard**         | ⚠️ HARDCODE               | `"SUCCESS"`, `"REAL TIME"`, `"0"`, `"0"`                                  |
| **1.9.1 DBA Users**       | `nodes[0]['database_info']` | Key `"DBA_ROLE"` (filter_nulls=True)                                     |
| **1.9.2 Objects System**  | `nodes[0]['database_info']` | Key `"OBJECT_IN_SYSTEM"` (lọc pattern NULL NULL)                         |
| **1.10 Patches**          | **Mọi node** (gộp)        | `node['database_info']['CHECK_PATCHES']` — Bỏ header từ node 2 trở đi    |

---

## GHI CHÚ QUAN TRỌNG KHI PHÁT TRIỂN TÍNH NĂNG MỚI

### [HARDCODE] Section cần được cải thiện về sau:
- **1.1 Status Check:** Hiện trả về `"Running"/"Open"` cứng → cần parser đọc từ `database_information.html` nếu có dữ liệu thực.
- **1.8 Dataguard:** Hiện trả về `"SUCCESS"/"0"` cứng → cần nguồn dữ liệu thực từ Data Guard logs.
- **1.7.2 Scheduling:** Text backup schedule là hardcode → có thể cho phép user cấu hình.

### [ẢNH OSWBB & EXAWATCHER] Điều kiện tiên quyết:
Các section 1.3.1, 1.3.2, 1.3.3 đọc file `.jpg`/`.gif`/`.png` từ `node['data_dir']/generated_files/` hoặc `exawatcher_files/`.  
**Quy tắc thư mục mặc định (v2.3.0):**
- Nếu không chọn Output Folder, OSWBB sẽ tạo folder `generated_files` tại thư mục chứa file `.exe`.
- Nếu không chọn Output Folder, ExaWatcher sẽ tạo folder `exawatcher_files` tại thư mục chứa file `.exe`.
- **Tính năng Sync:** Hỗ trợ đẩy (Push) các folder ảnh này vào các thư mục Node đích (Target Folders for Sync) ngay sau khi tạo xong.

### [CHỈ NODE 0] Các section không lặp qua tất cả node:
Các section `1.3.7`, `1.3.8`, `1.3.9`, `1.3.10`, `1.4.x`, `1.5.2`, `1.7`, `1.9` chỉ đọc từ `nodes[0]`.  
Đây là thiết kế cố ý (dữ liệu DB-wide như tablespace, patches là như nhau ở mọi node).

### [SQL ID MAPPING] Vị trí cột cứng:
Trong `_add_filtered_sql_text_table()`:
- SQL ID được lấy từ **cột index 6** (`row[6]`) của bảng Top SQL.
- Nếu format AWR báo cáo thay đổi số cột, cần cập nhật giá trị này.

---

*Cập nhật lần cuối: 16/04/2026*
