@echo off
REM ==========================================
REM Script build ứng dụng Oracle RAC Report Generator
REM Yêu cầu: Đã cài đặt PyInstaller trong venv
REM = [1] Chạy scripts\create_jre_mini.bat trước nếu chưa có thư mục dist\jre_mini
REM ==========================================

echo [INFO] Kiem tra JRE noi bo...
if not exist "dist\jre_mini\bin\java.exe" (
    echo [WARNING] Khong tim thay JRE tai dist\jre_mini.
    echo [INFO] Dang tu dong chay create_jre_mini.bat...
    call scripts\create_jre_mini.bat
)

echo [INFO] Dang bat dau qua trinh tang version va dong goi...
venv\Scripts\python scripts\packager.py

if %ERRORLEVEL% EQU 0 (
    echo [SUCCESS] Dong goi hoan tat! Vui long kiem tra thu muc dist.
) else (
    echo [ERROR] Co loi xay ra trong qua trinh dong goi.
)
pause
