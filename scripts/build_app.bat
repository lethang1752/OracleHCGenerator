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

echo [INFO] Dang bat dau qua trinh dong goi voi PyInstaller...
pyinstaller --clean build.spec

if %ERRORLEVEL% EQU 0 (
    echo [SUCCESS] Dong goi thanh cong! 
    echo [INFO] File thuc thi nam tai: dist\OracleRACReport_OSWBB.exe
) else (
    echo [ERROR] Qua trinh dong goi gap loi.
)
pause
