import json
import logging
import os
from pathlib import Path
from typing import Dict, Any, List

from ..config import BASE_DIR

logger = logging.getLogger(__name__)

RULES_FILE = BASE_DIR / "data" / "recommendation_rules.json"

DEFAULT_RULES = {
    "1.2": {
        "title": "Alert Log",
        "category": "LOG",
        "column": "MESSAGE",
        "condition": "contains_any",
        "threshold": ["ORA-00600", "ORA-07445", "ORA-04031", "ORA-01555", "ORA-00060"],
        "rec_en": "Critical database errors detected. Please check alert log and trace files.",
        "rec_vi": "Phát hiện lỗi cơ sở dữ liệu nghiêm trọng. Vui lòng kiểm tra alert log và các file trace.",
        "risk_en": "Risk of instance crash or data corruption.",
        "risk_vi": "Nguy cơ treo instance hoặc lỗi cấu trúc dữ liệu.",
        "severity_en": "Critical",
        "severity_vi": "Nghiêm trọng"
    },
    "1.3.1": {
        "title": "CPU Usage",
        "category": "AWR",
        "column": "%CPU",
        "condition": ">",
        "threshold": 80,
        "rec_en": "High CPU utilization detected. Consider optimizing top SQL or adding CPU resources.",
        "rec_vi": "Sử dụng CPU cao. Xem xét tối ưu hóa SQL top đầu hoặc bổ sung tài nguyên CPU.",
        "risk_en": "System performance degradation and slow response times.",
        "risk_vi": "Giảm hiệu năng hệ thống và thời gian phản hồi chậm.",
        "severity_en": "High",
        "severity_vi": "Cao"
    },
    "1.3.2": {
        "title": "Memory Usage",
        "category": "AWR",
        "column": "%MEMORY",
        "condition": ">",
        "threshold": 90,
        "rec_en": "High memory consumption. Check for memory leaks or resize SGA/PGA components.",
        "rec_vi": "Tiêu thụ bộ nhớ cao. Kiểm tra rò rỉ bộ nhớ hoặc điều chỉnh kích thước SGA/PGA.",
        "risk_en": "Risk of OOM (Out of Memory) errors and swapping.",
        "risk_vi": "Nguy cơ lỗi thiếu bộ nhớ (OOM) và hiện tượng swap.",
        "severity_en": "High",
        "severity_vi": "Cao"
    },
    "1.3.4": {
        "title": "Top Queries",
        "category": "AWR_SQL",
        "column": "ELAPSED TIME (S)",
        "condition": ">",
        "threshold": 1000,
        "rec_en": "Long running SQL statements detected. Perform SQL tuning/execution plan analysis.",
        "rec_vi": "Phát hiện các câu lệnh SQL chạy lâu. Thực hiện tinh chỉnh SQL/phân tích kế hoạch thực thi.",
        "risk_en": "Blocking sessions and excessive resource consumption.",
        "risk_vi": "Gây nghẽn (blocking) và tiêu tốn quá mức tài nguyên.",
        "severity_en": "Medium",
        "severity_vi": "Trung bình"
    },
    "1.3.6": {
        "title": "Top Wait Events",
        "category": "AWR_WAIT",
        "column": "AVG WAIT (MS)",
        "condition": ">",
        "threshold": 20,
        "rec_en": "Significant wait events detected. Investigate I/O, network or locking issues.",
        "rec_vi": "Phát hiện sự kiện chờ đợi đáng kể. Điều tra các vấn đề về I/O, mạng hoặc tranh chấp khóa.",
        "risk_en": "Application latency and database performance bottlenecks.",
        "risk_vi": "Tăng độ trễ ứng dụng và tạo ra điểm nghẽn hiệu năng.",
        "severity_en": "Medium",
        "severity_vi": "Trung bình"
    },
    "1.3.7": {
        "title": "Disk Group Usage",
        "category": "DB_INFO",
        "target": "ASM",
        "column": "USED (%)",
        "condition": ">",
        "threshold": 85,
        "rec_en": "ASM Disk Group capacity is reaching threshold. Add more disks or clean up old files.",
        "rec_vi": "Dung lượng ASM Disk Group sắp đạt ngưỡng. Bổ sung đĩa hoặc dọn dẹp các tệp cũ.",
        "risk_en": "Database might stop if ASM disk groups become full.",
        "risk_vi": "Cơ sở dữ liệu có thể dừng hoạt động nếu ASM disk group đầy.",
        "severity_en": "High",
        "severity_vi": "Cao"
    },
    "1.3.8": {
        "title": "Tablespace Usage",
        "category": "DB_INFO",
        "target": "TABLESPACE",
        "column": "USED (%)",
        "condition": ">",
        "threshold": 90,
        "rec_en": "Tablespace usage is high. Add datafiles or enable autoextend.",
        "rec_vi": "Sử dụng Tablespace đang ở mức cao. Bổ sung datafile hoặc bật autoextend.",
        "risk_en": "Transaction failures and application downtime due to space allocation errors.",
        "risk_vi": "Lỗi giao dịch và dừng ứng dụng do không thể cấp phát thêm không gian.",
        "severity_en": "Critical",
        "severity_vi": "Nghiêm trọng"
    },
    "1.3.9": {
        "title": "Index Fragmentation",
        "category": "DB_INFO",
        "target": "INDEX_FRAGMENT",
        "column": "FRAG (%)",
        "condition": ">",
        "threshold": 30,
        "rec_en": "High index fragmentation detected. Perform index rebuild or coalesce.",
        "rec_vi": "Phát hiện phân mảnh chỉ mục cao. Thực hiện rebuild hoặc coalesce index.",
        "risk_en": "Wasted space and slower query performance.",
        "risk_vi": "Lãng phí không gian và giảm hiệu năng truy vấn.",
        "severity_en": "Low",
        "severity_vi": "Thấp"
    },
    "1.3.10": {
        "title": "Table Fragmentation",
        "category": "DB_INFO",
        "target": "TABLE_FRAGMENT",
        "column": "WASTED (%)",
        "condition": ">",
        "threshold": 30,
        "rec_en": "High table fragmentation detected. Perform table shrinking or move.",
        "rec_vi": "Phát hiện phân mảnh bảng cao. Thực hiện shrink hoặc move table.",
        "risk_en": "Increased I/O and slower full table scans.",
        "risk_vi": "Tăng I/O và làm chậm các thao tác quét toàn bộ bảng (full table scan).",
        "severity_en": "Low",
        "severity_vi": "Thấp"
    },
    "1.4.1": {
        "title": "Invalid Objects",
        "category": "DB_INFO",
        "target": "INVALID_OBJECT",
        "column": "COUNT",
        "condition": ">",
        "threshold": 0,
        "rec_en": "Invalid objects found. Recompile them using utlrp.sql.",
        "rec_vi": "Tìm thấy các object bị invalid. Thực hiện recompile bằng utlrp.sql.",
        "risk_en": "Functional errors in application logic and stored procedures.",
        "risk_vi": "Lỗi chức năng trong logic ứng dụng và các stored procedure.",
        "severity_en": "Low",
        "severity_vi": "Thấp"
    },
    "1.4.2": {
        "title": "Table/Index Statistics",
        "category": "DB_INFO",
        "target": "TABLE_STATISTICS",
        "column": "COUNT",
        "condition": ">",
        "threshold": 0,
        "rec_en": "Stale or missing statistics detected. Gather statistics to ensure optimal execution plans.",
        "rec_vi": "Phát hiện thống kê bị cũ hoặc thiếu. Thực hiện gather statistics để đảm bảo kế hoạch thực thi tối ưu.",
        "risk_en": "Poor query performance due to incorrect optimizer decisions.",
        "risk_vi": "Hiệu năng truy vấn kém do bộ tối ưu hóa đưa ra quyết định sai.",
        "severity_en": "Low",
        "severity_vi": "Thấp"
    },
    "1.6": {
        "title": "Storage Capacity",
        "category": "DB_INFO",
        "target": "DISK_USAGE",
        "column": "USED (%)",
        "condition": ">",
        "threshold": 90,
        "rec_en": "OS partition usage is critical. Purge logs or expand file system.",
        "rec_vi": "Sử dụng phân vùng hệ điều hành ở mức tới hạn. Xóa log hoặc mở rộng file system.",
        "risk_en": "OS instability and alert log writing failures.",
        "risk_vi": "Mất ổn định hệ điều hành và không thể ghi log hệ thống.",
        "severity_en": "Critical",
        "severity_vi": "Nghiêm trọng"
    },
    "1.7.1": {
        "title": "Backup Status",
        "category": "DB_INFO",
        "target": "CHECK_BACKUP",
        "column": "STATUS",
        "condition": "!=",
        "threshold": "COMPLETED",
        "rec_en": "Recent backup failed or did not run. Investigate RMAN logs immediately.",
        "rec_vi": "Sao lưu gần đây thất bại hoặc chưa chạy. Kiểm tra log RMAN ngay lập tức.",
        "risk_en": "Risk of total data loss if a disaster occurs.",
        "risk_vi": "Nguy cơ mất trắng dữ liệu nếu xảy ra thảm họa.",
        "severity_en": "Critical",
        "severity_vi": "Nghiêm trọng"
    },
    "1.9.1": {
        "title": "DBA Users",
        "category": "DB_INFO",
        "target": "DBA_ROLE",
        "column": "GRANTEE",
        "condition": "count",
        "threshold": 5,
        "rec_en": "Excessive number of users with DBA role detected. Review and revoke unnecessary privileges.",
        "rec_vi": "Phát hiện số lượng lớn người dùng có quyền DBA. Rà soát và thu hồi các quyền không cần thiết.",
        "risk_en": "Security risk and potential for unauthorized or accidental changes.",
        "risk_vi": "Rủi ro bảo mật và khả năng xảy ra các thay đổi trái phép hoặc vô ý.",
        "severity_en": "Medium",
        "severity_vi": "Trung bình"
    },
    "1.9.2": {
        "title": "Users with Objects in SYSTEM/SYSAUX",
        "category": "DB_INFO",
        "target": "OBJECT_IN_SYSTEM",
        "column": "OWNER",
        "condition": "count",
        "threshold": 0,
        "rec_en": "Non-system objects found in SYSTEM/SYSAUX tablespaces. Move them to appropriate user tablespaces.",
        "rec_vi": "Tìm thấy các object không thuộc hệ thống trong tablespace SYSTEM/SYSAUX. Di chuyển chúng sang các tablespace dữ liệu phù hợp.",
        "risk_en": "Potential to cause system tablespace exhaustion, impacting database operation.",
        "risk_vi": "Có khả năng gây cạn kiệt tablespace hệ thống, ảnh hưởng đến hoạt động của cơ sở dữ liệu.",
        "severity_en": "Medium",
        "severity_vi": "Trung bình"
    },
    "1.10": {
        "title": "Patch Update Status",
        "category": "DB_INFO",
        "target": "CHECK_PATCHES",
        "column": "DAYS_SINCE",
        "condition": ">",
        "threshold": 180,
        "rec_en": "Database has not been patched for over 6 months. Plan for latest PSU/RU update.",
        "rec_vi": "Hệ quản trị CSDL chưa được vá lỗi trên 6 tháng. Lập kế hoạch cập nhật PSU/RU mới nhất.",
        "risk_en": "Vulnerability to known security threats and bugs.",
        "risk_vi": "Rủi ro trước các lỗ hổng bảo mật và lỗi đã được công bố.",
        "severity_en": "Low",
        "severity_vi": "Thấp"
    }
}

class RulesManager:
    """Manages recommendation rules and thresholds persistence"""
    
    @staticmethod
    def load_rules() -> Dict[str, Any]:
        """Load rules from JSON, gộp thêm các mục mặc định còn thiếu (Smart Merge)"""
        current_rules = {}
        if RULES_FILE.exists():
            try:
                with open(RULES_FILE, 'r', encoding='utf-8') as f:
                    current_rules = json.load(f)
            except Exception as e:
                logger.error(f"Failed to load rules: {e}")
        
        # Smart Merge: Kiểm tra xem có mục nào trong DEFAULT_RULES bị thiếu không
        is_modified = False
        for rid, default_rule in DEFAULT_RULES.items():
            if rid not in current_rules:
                current_rules[rid] = default_rule
                is_modified = True
                logger.info(f"Smart Merge: Added missing rule section {rid}")
        
        # Lưu lại nếu có thay đổi để lần sau không phải merge lại
        if is_modified or not RULES_FILE.exists():
            RulesManager.save_rules(current_rules)
            
        return current_rules

    @staticmethod
    def save_rules(rules: Dict[str, Any]) -> bool:
        """Save rules to JSON file"""
        try:
            os.makedirs(RULES_FILE.parent, exist_ok=True)
            with open(RULES_FILE, 'w', encoding='utf-8') as f:
                json.dump(rules, f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            logger.error(f"Failed to save rules: {e}")
            return False

    @staticmethod
    def reset_rules():
        """Reset to built-in defaults"""
        RulesManager.save_rules(DEFAULT_RULES)
        return DEFAULT_RULES
