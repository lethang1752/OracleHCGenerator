import re
import json
import logging
import tarfile
import io
import os
import traceback
from pathlib import Path
from typing import List, Dict, Optional, Tuple, Union
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime
from PyQt5.QtCore import QObject, pyqtSignal

from .logger import setup_logger

logger = setup_logger(__name__)

class LogSource:
    """Abstract base class for log data sources"""
    def __init__(self, progress_callback=None):
        self.progress_callback = progress_callback

    def log(self, msg: str):
        if self.progress_callback:
            self.progress_callback(msg)
        else:
            logger.info(msg)

    def get_content(self, suffix: str) -> Optional[str]:
        raise NotImplementedError

    def exists(self) -> bool:
        raise NotImplementedError

class DirectoryLogSource(LogSource):
    """Source for logs in a local directory"""
    def __init__(self, path: Union[str, Path], progress_callback=None):
        super().__init__(progress_callback)
        self.path = Path(path)

    def exists(self) -> bool:
        return self.path.exists() and self.path.is_dir()

    def get_content(self, suffix: str) -> Optional[str]:
        for p in self.path.rglob(f"*{suffix}"):
            if p.is_file():
                try:
                    return p.read_text(encoding='utf-8', errors='ignore')
                except Exception as e:
                    logger.warning(f"Failed to read {p}: {e}")
        return None

class TarLogSource(LogSource):
    """Source for logs inside a .tar.bz2 archive (streaming in RAM)"""
    def __init__(self, path: Union[str, Path], progress_callback=None):
        super().__init__(progress_callback)
        self.path = Path(path)
        self._content_cache = {}
        self._scanned = False

    def exists(self) -> bool:
        return self.path.exists() and self.path.is_file()

    def _scan_if_needed(self, suffixes: List[str]):
        if self._scanned:
            return
        
        found_suffixes = set()
        try:
            self.log(f"[INFO] Đang quét hệ thống tệp trong archive: {self.path.name}")
            with tarfile.open(self.path, "r:bz2") as tar:
                # Use iterator instead of getmembers() to avoid indexing whole archive
                file_count = 0
                for member in tar:
                    file_count += 1
                    if member.isfile():
                        for suffix in suffixes:
                            if suffix not in found_suffixes and member.name.endswith(suffix):
                                f = tar.extractfile(member)
                                if f:
                                    self._content_cache[suffix] = f.read().decode('utf-8', errors='ignore')
                                    found_suffixes.add(suffix)
                                    self.log(f"[OK] Đã tìm thấy {member.name}")
                                    
                    # Stop early if we have everything we need
                    if len(found_suffixes) >= len(suffixes):
                        self.log(f"[INFO] Kết thúc sớm: Đã tìm đủ các tệp cần thiết sau khi quét {file_count} mục.")
                        break
                else:
                    self.log(f"[INFO] Quét xong: Đã duyệt {file_count} mục. Tìm thấy {len(found_suffixes)}/{len(suffixes)} tệp yêu cầu.")
            self._scanned = True
        except Exception as e:
            logger.error(f"Failed to stream archive {self.path}: {e}")

    def get_content(self, suffix: str) -> Optional[str]:
        # We pre-scan for common ExaWatcher suffixes to avoid multiple passes
        self._scan_if_needed(["_mp.html", "_meminfo.html", "_iosummary.html"])
        return self._content_cache.get(suffix)

class ExaWatcherGraphGenerator(QObject):
    """
    Consolidates ExaWatcher HTML data and generates high-quality static images.
    Mimics OSWBB output structure.
    """
    progress = pyqtSignal(str)      # Sends text logs
    progress_val = pyqtSignal(int)  # Sends numerical percentage (0-100)
    finished = pyqtSignal(bool)     # Finished (True = Success)

    def __init__(self, db_node_source: str, cell_node_source: str, output_folder: str, 
                 push_targets: list = None, push_mode: str = "overwrite"):
        super().__init__()
        self.db_source_path = Path(db_node_source)
        self.cell_source_path = Path(cell_node_source)
        # Ensure output folder is absolute, especially important for frozen EXE
        self.output_folder = Path(output_folder).resolve()
        self.push_targets = push_targets or []
        self.push_mode = push_mode # "overwrite" | "timestamp"
        self._stop_requested = False

        # Factory for log sources
        self.db_log_source = self._create_source(self.db_source_path)
        self.cell_log_source = self._create_source(self.cell_source_path)

    def _create_source(self, path: Path) -> LogSource:
        if path.is_file() and (path.suffix == '.bz2' or '.tar' in path.name):
            return TarLogSource(path, progress_callback=self.progress.emit)
        return DirectoryLogSource(path, progress_callback=self.progress.emit)

    def stop(self):
        self._stop_requested = True

    def run(self):
        import concurrent.futures
        try:
            self.progress_val.emit(5)
            self.output_folder.mkdir(parents=True, exist_ok=True)
            self.progress.emit(f"[INFO] Bắt đầu xử lý ExaWatcher (Parallel Mode)...")
            self.progress.emit(f"[PATH] Output folder: {self.output_folder}")

            image_count = 0
            
            with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
                future_to_source = {}
                
                # 1. Process CPU & Memory (Source A: DB/VM)
                if self.db_log_source.exists():
                    future_to_source[executor.submit(self._process_cpu_mem)] = "DB/VM Source"
                    self.progress_val.emit(25)
                else:
                    self.progress.emit(f"[WARN] Nguồn dữ liệu DB/VM không hợp lệ: {self.db_source_path.name}")

                # 2. Process IO (Source B: Cell)
                if self.cell_log_source.exists():
                    future_to_source[executor.submit(self._process_io)] = "Cell Source"
                    self.progress_val.emit(50)
                else:
                    self.progress.emit(f"[WARN] Nguồn dữ liệu Cell không hợp lệ: {self.cell_source_path.name}")

                # Collect results and catch exceptions from threads
                for future in concurrent.futures.as_completed(future_to_source):
                    source_name = future_to_source[future]
                    try:
                        count = future.result()
                        image_count += count
                        current_prog = 50 + (25 if image_count > 0 else 0)
                        self.progress_val.emit(min(current_prog, 85))
                        self.progress.emit(f"[SUCCESS] {source_name} xử lý xong (Tạo được {count} ảnh).")
                    except Exception as e:
                        err_detail = traceback.format_exc()
                        logger.error(f"Error in {source_name}: {e}\n{err_detail}")
                        self.progress.emit(f"[ERROR] Thất bại tại {source_name}: {str(e)}")

            if self._stop_requested:
                self.progress.emit("[INFO] Đã dừng theo yêu cầu người dùng.")
                self.finished.emit(False)
                return

            if image_count == 0:
                self.progress.emit(f"[ERROR] Không tạo được ảnh nào. Vui lòng kiểm tra lại định dạng tệp log.")
                self.finished.emit(False)
                return

            self.progress.emit(f"[SUCCESS] Hoàn tất tạo {image_count} biểu đồ ExaWatcher.")
            self.progress_val.emit(90)
            
            # 3. PUSH RESULTS (Synchronize files to target folders)
            if self.push_targets:
                self.progress_val.emit(95)
                push_success = self._push_results()
                if not push_success:
                    self.finished.emit(False)
                    return

            self.progress_val.emit(100)
            self.finished.emit(True)

        except Exception as e:
            logger.error(f"ExaWatcher processing failed: {e}", exc_info=True)
            self.progress.emit(f"[ERROR] Lỗi xử lý: {str(e)}")
            self.finished.emit(False)

    def _robust_rmtree(self, path):
        """Robustly delete directory, handling read-only files and retries"""
        import time
        import shutil
        import os
        import stat

        def remove_readonly(func, path, excinfo):
            os.chmod(path, stat.S_IWRITE)
            func(path)

        for i in range(3):
            try:
                if os.path.exists(path):
                    shutil.rmtree(path, onerror=remove_readonly)
                return True
            except Exception:
                time.sleep(0.5)
        return False

    def _push_results(self) -> bool:
        """Copy generated images to target folders"""
        import shutil
        import datetime
        from pathlib import Path
        
        # We push all files from the output_folder
        if not self.output_folder.exists() or not any(self.output_folder.iterdir()):
            self.progress.emit("[ERROR] Không có dữ liệu kết quả để Push.")
            return False
            
        self.progress.emit(f"[SYNC] Bắt đầu đẩy dữ liệu tới {len(self.push_targets)} mục tiêu...")
        
        try:
            for target in self.push_targets:
                target_path = Path(target)
                if not target_path.exists():
                    self.progress.emit(f"[SKIP] Thư mục đích không tồn tại: {target}")
                    continue
                
                # Determine folder name
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                folder_name = f"exawatcher_files_{timestamp}" if self.push_mode == "timestamp" else "exawatcher_files"
                final_dest = target_path / folder_name
                
                self.progress.emit(f"[COPYING] Đang sao chép tới {final_dest}...")
                
                if final_dest.exists():
                    if self.push_mode == "overwrite":
                        if not self._robust_rmtree(final_dest):
                            self.progress.emit(f"[WARNING] Không thể ghi đè '{final_dest}', folder đang bị khóa.")
                            continue
                
                # Robust Copy with retries
                copy_success = False
                for attempt in range(3):
                    try:
                        shutil.copytree(str(self.output_folder), str(final_dest))
                        copy_success = True
                        break
                    except Exception as e:
                        if attempt < 2:
                            import time
                            time.sleep(1)
                        else:
                            self.progress.emit(f"[ERROR] Không thể sao chép tới {final_dest}: {e}")
                
                if copy_success:
                    self.progress.emit(f"[SUCCESS] Đã đẩy tới {target}")
                
            return True
        except Exception as e:
            self.progress.emit(f"[CRITICAL ERROR] Quá trình Push thất bại: {str(e)}")
            return False


    def _extract_js_var(self, html_content: str, var_names: List[str]) -> Dict:
        """Extract JavaScript variable values using regex (supports 'var x=', 'self.x =')"""
        results = {}
        # Pre-compile patterns for speed
        for var_name in var_names:
            pattern = re.compile(rf"(?:var\s+|self\.){var_name}\s*=\s*")
            start_match = pattern.search(html_content)
            if start_match:
                start_pos = start_match.end()
                end_pos = html_content.find(';', start_pos)
                if end_pos != -1:
                    raw_val = html_content[start_pos:end_pos].strip()
                    try:
                        results[var_name] = json.loads(raw_val)
                        logger.info(f"Extracted variable {var_name}")
                    except Exception as e:
                        logger.warning(f"Failed to parse variable {var_name}: {e}")
            else:
                logger.debug(f"Variable {var_name} not found")
        return results

    def _setup_plot(self, title: str):
        """Standard matplotlib setup for Oracle-style charts (1000x350px)"""
        # 10x3.5 inches at default 100dpi = 1000x350px
        fig, ax = plt.subplots(figsize=(10, 3.5))
        ax.set_title(title, fontsize=12, fontweight='bold', pad=15)
        ax.grid(True, linestyle='--', alpha=0.3, linewidth=0.5)
        
        # Design: Light sky blue outer background (#C6D9F1), White plotting area
        ax.set_facecolor('white')
        fig.patch.set_facecolor('#C6D9F1') 
        
        # 3-hour interval timeline with 24h hour-only labels
        ax.xaxis.set_major_locator(mdates.HourLocator(byhour=[0, 3, 6, 9, 12, 15, 18, 21]))
        
        def time_formatter(x, pos):
            dt = mdates.num2date(x)
            if dt.hour == 0:
                return dt.strftime('%Hh\n%d %b')
            return dt.strftime('%Hh')
            
        ax.xaxis.set_major_formatter(plt.FuncFormatter(time_formatter))
        plt.xticks(fontsize=9, rotation=0)
        plt.yticks(fontsize=9)
        
        # Remove padding at the beginning and end of the X axis
        ax.set_xmargin(0)
        
        # We will use tight_layout() right before saving in each specific 
        # process method to ensure it considers all labels/titles.
        return fig, ax

    def _process_cpu_mem(self) -> int:
        images_generated = 0
        # CPU
        content = self.db_log_source.get_content("_mp.html")
        if content:
            self.progress.emit("[INFO] Đang vẽ biểu đồ CPU từ nguồn dữ liệu...")
            data_vars = self._extract_js_var(content, ["xAxis", "series"])
            
            if "xAxis" in data_vars and "series" in data_vars:
                times = [datetime.fromisoformat(t) for t in data_vars["xAxis"]]
                series_data = data_vars["series"].get("all", [])
                
                targets = {
                    "idl": ("CPU Idle Utilization", "Exa_Cpu_Idle.png"),
                    "sys": ("CPU System Utilization", "Exa_Cpu_Sys.png"),
                    "usr": ("CPU User Utilization", "Exa_Cpu_Usr.png")
                }
                
                for sid, (title, out_name) in targets.items():
                    target_series = next((s for s in series_data if s.get('id') == sid), None)
                    if target_series and len(target_series.get('items', [])) == len(times):
                        fig, ax = self._setup_plot(title)
                        vals = target_series['items']
                        
                        # Scale fractions to percentages (0.27 -> 27.0)
                        max_val = max((v for v in vals if v is not None), default=0)
                        if max_val <= 1.05 and max_val > 0:
                            vals = [v * 100 if v is not None else None for v in vals]
                        
                        label_name = target_series.get('name', sid)
                        ax.plot(times, vals, label=label_name, color='#0067C0', linewidth=1.2)
                        ax.fill_between(times, vals, color='#0067C0', alpha=0.1)
                        
                        ax.set_ylim(0, 105)
                        # Format Y-ticks as % (No decimals)
                        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f"{int(x)}%"))
                        ax.legend(loc='upper right', fontsize=9, framealpha=0.8)
                        
                        # Use fixed size export with dynamic margin calculation
                        fig.tight_layout()
                        fig.savefig(self.output_folder / out_name)
                        plt.close(fig)
                        images_generated += 1
                    else:
                        logger.warning(f"Series {sid} not found or length mismatch in _mp.html")

        # Memory
        content = self.db_log_source.get_content("_meminfo.html")
        if content:
            self.progress.emit("[INFO] Đang vẽ biểu đồ Memory từ nguồn dữ liệu...")
            # Look for all common memory variable names
            data_vars = self._extract_js_var(content, ["xAxis", "data", "refObjects"])
            
            if "xAxis" in data_vars and "data" in data_vars:
                times = [datetime.fromisoformat(t) for t in data_vars["xAxis"]]
                mem_data_root = data_vars["data"]
                
                # Chart 1: OS Memory (usually 'osmem' or 'memChart')
                os_mem_key = "osmem" if "osmem" in mem_data_root else "memChart"
                if os_mem_key in mem_data_root:
                    fig, ax = self._setup_plot("OS Memory Utilization")
                    series = mem_data_root[os_mem_key]
                    for s in series:
                        if len(s.get('items', [])) == len(times):
                            ax.plot(times, s['items'], label=s['name'], linewidth=1.1)
                    
                    # Add MemTotal if available in refObjects
                    ref = data_vars.get("refObjects", {}).get(os_mem_key, [])
                    for r in ref:
                        if r.get("text") == "MemTotal" and "value" in r:
                            ax.axhline(y=r["value"], color='#C74634', linestyle='--', label='MemTotal', linewidth=1.2)
                    
                    ax.legend(loc='upper right', fontsize=9, framealpha=0.8)
                    # Format Y-ticks as GB (No decimals, with comma)
                    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f"{int(x):,}GB"))
                    fig.tight_layout()
                    fig.savefig(self.output_folder / "Exa_Mem_OS.png")
                    plt.close(fig)
                    images_generated += 1

                # Chart 2: HugePages (usually 'hp' or 'hugePagesChart')
                hp_key = "hp" if "hp" in mem_data_root else "hugePagesChart"
                if hp_key in mem_data_root:
                    fig, ax = self._setup_plot("HugePages Utilization")
                    series = mem_data_root[hp_key]
                    for s in series:
                        if len(s.get('items', [])) == len(times):
                            ax.plot(times, s['items'], label=s['name'], linewidth=1.1)
                    
                    ax.legend(loc='upper right', fontsize=9, framealpha=0.8)
                    # Format Y-ticks as GB/Unit (e.g., 1,200GB)
                    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f"{int(x):,}GB"))
                    fig.tight_layout()
                    fig.savefig(self.output_folder / "Exa_Mem_HugePages.png")
                    plt.close(fig)
                    images_generated += 1
        return images_generated

    def _process_io(self) -> int:
        images_generated = 0
        content = self.cell_log_source.get_content("_iosummary.html")
        if content:
            self.progress.emit("[INFO] Đang vẽ biểu đồ IO từ nguồn dữ liệu...")
            data_vars = self._extract_js_var(content, ["xAxis", "data"])
            
            if "xAxis" in data_vars and "data" in data_vars:
                times = [datetime.fromisoformat(t) for t in data_vars["xAxis"]]
                io_data_root = data_vars["data"]
                
                # Use 'flash' if available and has data, otherwise 'disk'
                dtype = "flash"
                if not io_data_root.get(dtype) or not io_data_root[dtype].get("iops") or not io_data_root[dtype]["iops"][0].get("items"):
                    dtype = "disk"
                
                if dtype in io_data_root:
                    stats = io_data_root[dtype]
                    # We want IOPS, RPS, WPS in one chart
                    fig, ax = self._setup_plot(f"Cell {dtype.capitalize()} I/O Summary")
                    
                    # Mapping of internal IDs to legend names
                    plot_targets = {
                        "iops": "io/s",
                        "rps": "r/s",
                        "wps": "w/s"
                    }
                    
                    colors = {"iops": "#0067C0", "rps": "#4C825C", "wps": "#AA643B"}
                    
                    has_data = False
                    # Extract r/s and w/s which are usually in the same array as 'iops'
                    # Jet structure: "iops": [{"id": "iops", ...}, {"id": "rps", ...}, {"id": "wps", ...}]
                    iops_series_list = stats.get("iops", [])
                    
                    for stat_id, label in plot_targets.items():
                        # Find the specific series within the iops list
                        s = next((ser for ser in iops_series_list if ser.get('id') == stat_id), None)
                        if s:
                            items = s.get('items', [])
                            if len(items) == len(times):
                                # Convert nulls to NaN for matplotlib
                                plot_items = [x if x is not None else float('nan') for x in items]
                                ax.plot(times, plot_items, label=label, color=colors.get(stat_id), linewidth=1.2)
                                has_data = True
                    
                    if has_data:
                        ax.legend(loc='upper right', fontsize=9, framealpha=0.8)
                        # Format Y-ticks with comma and /s (e.g., 20,000/s)
                        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f"{int(x):,}/s"))
                        fig.tight_layout()
                        fig.savefig(self.output_folder / "Exa_IO_Summary.png")
                        images_generated += 1
                    
                    plt.close(fig)
        return images_generated
