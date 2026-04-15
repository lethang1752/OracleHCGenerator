import os
import sys
import subprocess
import threading
from pathlib import Path
from PyQt5.QtCore import QObject, pyqtSignal

class OSWBBGraphGenerator(QObject):
    """
    Wrapper để chạy công cụ OSWBBA Java.
    Khởi tạo cùng một JRE nội bộ được tạo từ jlink.
    """
    progress = pyqtSignal(str)   # Gửi log dạng text
    finished = pyqtSignal(bool)  # Kết thúc (True = Thành công)
    
    def __init__(self, log_folder: str, output_folder: str, gen_dashboard: bool = True, push_targets: list = None, push_mode: str = "overwrite", jar_filename: str = "oswbba.jar"):
        super().__init__()
        self.log_folder = log_folder
        self.output_folder = output_folder
        self.gen_dashboard = gen_dashboard
        self.push_targets = push_targets or []
        self.push_mode = push_mode # "overwrite" | "timestamp"
        self.jar_filename = jar_filename
        self.process = None
        self.temp_filtered_dir = None
        
        self.ALLOWED_FOLDERS = [
            'oswcpuinfo', 'oswiostat', 'oswmeminfo', 
            'oswnetstat', 'oswvmstat'
        ]

    def run(self):
        """Tiến trình chính chạy trong Worker Thread"""
        self._run_java_process()

    def stop(self):
        """Dừng tiến trình Java cưỡng bách"""
        if self.process and self.process.poll() is None:
            try:
                self.progress.emit("[SYSTEM] Đang yêu cầu dừng tiến trình...")
                self.process.kill()
                self._cleanup_temp_dir()
                self.progress.emit("[SYSTEM] Tiến trình đã bị dừng bởi người dùng.")
            except Exception as e:
                self.progress.emit(f"[ERROR] Không thể dừng tiến trình: {e}")

    def _cleanup_temp_dir(self):
        """Dọn dẹp thư mục junctions tạm thời"""
        if self.temp_filtered_dir and self.temp_filtered_dir.exists():
            try:
                # On Windows, we need to remove junctions carefully
                # os.remove() or os.unlink() works on junctions as well
                for item in self.temp_filtered_dir.iterdir():
                    if item.is_dir():
                        if os.name == 'nt':
                            subprocess.run(f'rmdir "{item}"', shell=True, capture_output=True)
                        else:
                            os.unlink(item)
                
                import shutil
                shutil.rmtree(self.temp_filtered_dir)
                self.progress.emit("[CLEANUP] Đã dọn dẹp thư mục lọc tạm thời.")
            except Exception as e:
                self.progress.emit(f"[WARNING] Không thể tự động dọn dẹp folder tạm: {e}")
        
    def _get_base_path(self) -> Path:
        if getattr(sys, 'frozen', False):
            return Path(sys._MEIPASS)
        return Path.cwd()
        return Path.cwd()
        
    def _run_java_process(self):
        try:
            base_path = self._get_base_path()
            
            # Paths to Check (Bundle vs Source)
            jre_candidates = [
                base_path / "jre_mini" / "bin" / "java.exe",
                base_path / "dist" / "jre_mini" / "bin" / "java.exe",
            ]
            
            # Find Java
            java_cmd = None
            for candidate in jre_candidates:
                if candidate.exists():
                    java_cmd = str(candidate)
                    break
            
            if not java_cmd:
                self.progress.emit("[WARNING] Cảnh báo: Không tìm thấy JRE nội bộ. Đang thử fallback sang 'java' hệ thống.")
                java_cmd = "java"
                
            # Find JAR (Selected by user)
            jar_candidates = [
                base_path / "oswbb" / self.jar_filename,
                base_path / self.jar_filename,
            ]
            
            jar_path = None
            for candidate in jar_candidates:
                if candidate.exists():
                    jar_path = candidate
                    break

            if not jar_path:
                self.progress.emit(f"[ERROR] Không tìm thấy tệp {self.jar_filename} tại bất kỳ vị trí nào trong {base_path}")
                self.finished.emit(False)
                return

            # --- FILTERING LOGIC ---
            self.progress.emit("[INFO] Đang lọc dữ liệu chỉ lấy 5 nhóm thông số quan trọng để tăng tốc...")
            import tempfile
            import time
            # Use system temp directory to avoid OneDrive sync locks (Access Denied)
            self.temp_filtered_dir = Path(tempfile.gettempdir()) / f"osw_filter_{int(time.time())}"
            try:
                self.temp_filtered_dir.mkdir(parents=True, exist_ok=True)
                
                # Create junctions for allowed folders
                source_root = Path(self.log_folder)
                found_any = False
                for folder_name in self.ALLOWED_FOLDERS:
                    src_sub = source_root / folder_name
                    if src_sub.exists() and src_sub.is_dir():
                        dest_sub = self.temp_filtered_dir / folder_name
                        # Use mklink /J on Windows for fast directory junctions
                        if os.name == 'nt':
                            subprocess.run(f'mklink /J "{dest_sub}" "{src_sub}"', shell=True, capture_output=True)
                        else:
                            os.symlink(src_sub, dest_sub, target_is_directory=True)
                        found_any = True
                
                if not found_any:
                    self.progress.emit("[WARNING] Không thấy 5 thư mục osw* tiêu chuẩn, gửi toàn bộ thư mục gốc cho Java.")
                    analysis_input = os.path.normpath(self.log_folder)
                else:
                    analysis_input = os.path.normpath(str(self.temp_filtered_dir))
                    self.progress.emit("[INFO] Đã chuẩn bị bộ dữ liệu tinh gọn.")
            except Exception as e:
                self.progress.emit(f"[ERROR] Lỗi khi tạo folder lọc: {e}. Quay lại dùng folder gốc.")
                analysis_input = os.path.normpath(self.log_folder)

            # Build command list
            cmd = [
                java_cmd,
                "-Duser.language=en",
                "-Duser.country=US",
                "-jar", str(jar_path)
            ]

            if self.gen_dashboard:
                # Option D: Dashboard (Non-interactive)
                # Put -i first to ensure validation passes
                cmd.extend(["-i", analysis_input, "-D"])
            else:
                # Standard Interactive Mode (i.e. generate GIFs)
                cmd.extend(["-i", analysis_input])

            self.progress.emit(f"[INFO] Khởi chạy Java (Low-Impact Mode): {' '.join(cmd)}")
            
            # Windows specific: Lower priority to keep system responsive
            creation_flags = 0
            if os.name == 'nt':
                # CREATE_NO_WINDOW (0x08000000) | BELOW_NORMAL_PRIORITY_CLASS (0x4000)
                creation_flags = 0x08000000 | 0x00004000

            # Run Process with stdin/stdout pipes
            self.process = subprocess.Popen(
                cmd,
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                creationflags=creation_flags
            )

            # Track if we have triggered the automation
            automation_triggered = False

            # Scan stdout realtime
            for line in iter(self.process.stdout.readline, ''):
                if line:
                    stripped_line = line.strip()
                    self.progress.emit(stripped_line)
                    
                    if not self.gen_dashboard:
                        # Standard Interactive mode needs stdin input
                        if "Please Select an Option:" in stripped_line or "Parsing Completed" in stripped_line:
                            if not automation_triggered:
                                self.progress.emit("[SYSTEM] Đang gửi lệnh tự động: GD, GM, GC, Q...")
                                # Send commands to generate all graphs and quit
                                self.process.stdin.write("GD\nGM\nGC\nQ\n")
                                self.process.stdin.flush()
                                automation_triggered = True

            self.process.stdout.close()
            return_code = self.process.wait()

            # --- POST-PROCESSING & CLEANUP ---
            import shutil
            
            # 1. Handle GIF mode
            temp_gif_dir = Path.cwd() / "gif"
            if temp_gif_dir.exists():
                self.progress.emit(f"[INFO] Đang di chuyển ảnh GIF từ {temp_gif_dir} vào {self.output_folder}...")
                for f in temp_gif_dir.glob("*.gif"):
                    shutil.move(str(f), str(Path(self.output_folder) / f.name))
                shutil.rmtree(temp_gif_dir, ignore_errors=True)

            # 2. Handle Dashboard (-D) mode
            # Search for 'generated_files' inside 'analysis' folders
            analysis_base = Path.cwd() / "analysis"
            if analysis_base.exists():
                self.progress.emit("[INFO] Đang tìm kiếm thư mục 'generated_files' trong 'analysis'...")
                found_gen_folder = False
                for root, dirs, files in os.walk(analysis_base):
                    if "generated_files" in dirs:
                        gen_path = Path(root) / "generated_files"
                        dest_path = Path(self.output_folder) / "generated_files"
                        
                        # Move the folder as requested
                        if dest_path.exists():
                            shutil.rmtree(dest_path)
                        
                        shutil.move(str(gen_path), str(dest_path))
                        self.progress.emit(f"[SUCCESS] Đã chuyển folder {gen_path.name} vào {self.output_folder}")
                        found_gen_folder = True
                        break
                
                if not found_gen_folder and self.gen_dashboard:
                    self.progress.emit("[WARNING] Không tìm thấy thư mục 'generated_files' sau khi chạy Option D.")
            
            # --- END OF MAIN PROCESSING ---
            self._cleanup_temp_dir()
            
            if return_code == 0 or automation_triggered or self.gen_dashboard:
                # 3. PUSH RESULTS (New Feature)
                if self.push_targets:
                    push_success = self._push_results()
                    if not push_success:
                        self.finished.emit(False)
                        return

                self.progress.emit("[SUCCESS] OSWBBA Generator hoàn tất thành công.")
                self.finished.emit(True)
            else:
                self.progress.emit(f"[ERROR] OSWBBA thoái ra với lỗi code {return_code}.")
                self.finished.emit(False)

        except FileNotFoundError as e:
            self.progress.emit("[ERROR] Lỗi: Không thể tìm thấy trình khởi chạy Java (java.exe).")
            self.progress.emit("[ERROR] Nguyên nhân: Bạn chưa cài đặt Java trên máy (phiên bản DEV) hoặc chưa đóng gói JRE thu nhỏ.")
            self.progress.emit(">> Để khắc phục: Vui lòng chạy file scripts\\create_jre_mini.bat trên một máy có cài JDK để đóng gói JRE đi kèm ứng dụng.")
            self.finished.emit(False)
        except Exception as e:
            self.progress.emit(f"[EXCEPTION] Lỗi khi triệu gọi Java: {str(e)}")
            self.finished.emit(False)
        finally:
            self._final_cleanup()

    def _robust_rmtree(self, path):
        """Xóa thư mục một cách mạnh mẽ, xử lý cả tệp read-only và thử lại nhiều lần"""
        import time
        import shutil
        import stat

        def remove_readonly(func, path, excinfo):
            os.chmod(path, stat.S_IWRITE)
            func(path)

        for i in range(3): # Thử lại tối đa 3 lần
            try:
                if os.path.exists(path):
                    shutil.rmtree(path, onerror=remove_readonly)
                return True
            except Exception:
                time.sleep(0.5) # Đợi 0.5s để hệ thống giải phóng file
        return False

    def _final_cleanup(self):
        """Đảm bảo xóa sạch các thư mục tạm sau cùng"""
        analysis_path = os.path.join(os.getcwd(), "analysis")
        gif_path = os.path.join(os.getcwd(), "gif")
        
        if os.path.exists(analysis_path):
            self.progress.emit("[CLEANUP] Đang thực hiện xóa cưỡng bức thư mục 'analysis'...")
            if self._robust_rmtree(analysis_path):
                self.progress.emit("[CLEANUP] Đã xóa thư mục 'analysis'.")
            else:
                self.progress.emit("[WARNING] Không thể xóa thư mục 'analysis', vui lòng xóa thủ công.")

        if os.path.exists(gif_path):
            self.progress.emit("[CLEANUP] Đang thực hiện xóa cưỡng bức thư mục 'gif'...")
            if self._robust_rmtree(gif_path):
                self.progress.emit("[CLEANUP] Đã xóa thư mục 'gif'.")
            else:
                self.progress.emit("[WARNING] Không thể xóa thư mục 'gif', vui lòng xóa thủ công.")

    def _push_results(self) -> bool:
        """Sao chép thư mục generated_files tới các đích được chọn"""
        import shutil
        import datetime
        from pathlib import Path
        
        source_gen = Path(self.output_folder) / "generated_files"
        if not source_gen.exists():
            # If we generated GIFs instead of Dashboard
            if not any(Path(self.output_folder).iterdir()):
                self.progress.emit("[ERROR] Không có dữ liệu kết quả để Push.")
                return False
            # If generated_files is missing, do not fall back to the entire directory 
            # to avoid copying the whole project/app root.
            self.progress.emit("[ERROR] Không tìm thấy thư mục 'generated_files'. Không thể thực hiện Push.")
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
                folder_name = f"generated_files_{timestamp}" if self.push_mode == "timestamp" else "generated_files"
                final_dest = target_path / folder_name
                
                self.progress.emit(f"[COPYING] Đang sao chép tới {final_dest}...")
                
                if final_dest.exists():
                    if self.push_mode == "overwrite":
                        if not self._robust_rmtree(final_dest):
                            self.progress.emit(f"[WARNING] Không thể ghi đè '{final_dest}', folder đang bị khóa bởi tiến trình khác (có thể là OneDrive).")
                            continue # Try next target
                
                # Robust Copy with retries
                copy_success = False
                for attempt in range(3):
                    try:
                        shutil.copytree(str(source_gen), str(final_dest))
                        copy_success = True
                        break
                    except Exception as e:
                        if attempt < 2:
                            self.progress.emit(f"[RETRY] Thử lại sao chép ({attempt+1}/3)...")
                            import time
                            time.sleep(1)
                        else:
                            self.progress.emit(f"[ERROR] Không thể sao chép tới {final_dest}: {e}")
                
                if copy_success:
                    self.progress.emit(f"[SUCCESS] Đã đẩy tới {target}")
                
            return True
        except Exception as e:
            self.progress.emit(f"[CRITICAL ERROR] Quá trình Push thất bại tại {target}: {str(e)}")
            self.progress.emit("[SYSTEM] Dừng toàn bộ tiến trình theo yêu cầu.")
            return False
