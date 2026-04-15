"""
GitHub Synchronization Worker for Collection Tools
"""
import os
import json
import urllib.request
import logging
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal
from ..config import GITHUB_TOOLS_API_URL, COLLECT_TOOL_DIR

logger = logging.getLogger(__name__)

class GitHubSyncWorker(QThread):
    """Worker thread to sync collection tools from GitHub"""
    
    progress = pyqtSignal(int, str)  # Percentage and message
    finished = pyqtSignal(bool, str) # Success status and final message
    error = pyqtSignal(str)          # Error message
    
    def __init__(self):
        super().__init__()
        self._os_auth_token = os.environ.get("GITHUB_TOKEN") # Optional for higher rate limits

    def run(self):
        try:
            self.progress.emit(10, "Đang kết nối tới GitHub API...")
            
            # 1. Fetch file list from GitHub contents API
            headers = {'User-Agent': 'Oracle-HC-Generator-Client'}
            if self._os_auth_token:
                headers['Authorization'] = f'token {self._os_auth_token}'
                
            req = urllib.request.Request(GITHUB_TOOLS_API_URL, headers=headers)
            with urllib.request.urlopen(req, timeout=15) as response:
                if response.status != 200:
                    raise Exception(f"GitHub API returned status {response.status}")
                
                data = json.loads(response.read().decode('utf-8'))
            
            if not isinstance(data, list):
                raise Exception("Phản hồi từ GitHub không đúng định dạng mong đợi (List expected).")
            
            # Filter for files only
            files_to_download = [item for item in data if item['type'] == 'file']
            total_files = len(files_to_download)
            
            if total_files == 0:
                self.finished.emit(True, "Không tìm thấy file nào trong thư mục đích trên GitHub.")
                return

            self.progress.emit(20, f"Tìm thấy {total_files} tệp tin. Bắt đầu tải về...")
            
            # Ensure local directory exists
            os.makedirs(COLLECT_TOOL_DIR, exist_ok=True)
            
            # 2. Download each file
            for idx, file_info in enumerate(files_to_download):
                file_name = file_info['name']
                download_url = file_info.get('download_url')
                
                if not download_url:
                    logger.warning(f"File {file_name} không có download_url, bỏ qua.")
                    continue
                
                self.progress.emit(
                    20 + int((idx / total_files) * 70), 
                    f"Đang tải ({idx+1}/{total_files}): {file_name}..."
                )
                
                local_path = COLLECT_TOOL_DIR / file_name
                
                # Perform download
                file_req = urllib.request.Request(download_url, headers=headers)
                with urllib.request.urlopen(file_req, timeout=30) as file_response:
                    with open(local_path, 'wb') as f:
                        f.write(file_response.read())
            
            self.progress.emit(100, "Đồng bộ hoàn tất!")
            self.finished.emit(True, f"Đã cập nhật thành công {total_files} tệp tin từ GitHub.")
            
        except Exception as e:
            logger.error(f"GitHub Sync Error: {e}")
            self.error.emit(f"Lỗi đồng bộ: {str(e)}")
            self.finished.emit(False, str(e))
