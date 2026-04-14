"""
Database Models
"""
from datetime import datetime
from pathlib import Path

try:
    import sqlite3
    HAS_SQLITE = True
except ImportError:
    HAS_SQLITE = False


class ReportDatabase:
    """Simple SQLite database for storing reports"""
    
    def __init__(self, db_path: str):
        """
        Initialize database
        
        Args:
            db_path: Path to SQLite database
        """
        if not HAS_SQLITE:
            raise ImportError("sqlite3 not available")
        
        self.db_path = db_path
        Path(db_path).parent.mkdir(parents=True, exist_ok=True)
        self._init_db()
    
    def _init_db(self):
        """Initialize database schema"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Reports table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS reports (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                db_name TEXT NOT NULL,
                node1 TEXT,
                node2 TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                file_path TEXT,
                file_format TEXT
            )
        """)
        
        conn.commit()
        conn.close()
    
    def add_report(self, db_name: str, node1: str, node2: str, 
                   file_path: str, file_format: str) -> int:
        """Add report to database"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute("""
            INSERT INTO reports (db_name, node1, node2, file_path, file_format)
            VALUES (?, ?, ?, ?, ?)
        """, (db_name, node1, node2, file_path, file_format))
        
        report_id = cursor.lastrowid
        conn.commit()
        conn.close()
        
        return report_id
    
    def get_reports(self) -> list:
        """Get all reports"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute("SELECT * FROM reports ORDER BY created_at DESC")
        reports = cursor.fetchall()
        
        conn.close()
        return reports
    
    def delete_report(self, report_id: int) -> bool:
        """Delete report"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Get file path first
        cursor.execute("SELECT file_path FROM reports WHERE id = ?", (report_id,))
        result = cursor.fetchone()
        
        if result:
            file_path = result[0]
            # Delete from DB
            cursor.execute("DELETE FROM reports WHERE id = ?", (report_id,))
            conn.commit()
            
            # Delete file
            try:
                Path(file_path).unlink()
            except:
                pass
        
        conn.close()
        return True
