@echo off
REM ==========================================
REM Script copy JRE 8 cho OSWBBA Graph Generator
REM OSWBBA YÊU CẦU Java 8 do lỗi parse java.util.Date trên Java 9+ (kể cả có cấu hình locale).
REM Script này sẽ copy JRE 8 đã được cài (hoặc có sẵn) vào thư mục đích.
REM ==========================================

REM Đường dẫn Java 8 trên máy (có thể tùy chỉnh lại nếu máy khác)
set JAVA8_HOME=C:\Program Files\Eclipse Adoptium\jre-8.0.482.8-hotspot

echo [INFO] Kiem tra duong dan JRE 8...
if not exist "%JAVA8_HOME%" (
    echo [ERROR] Khong tim thay JRE 8 tai %JAVA8_HOME%. Vui long cai dat Eclipse Adoptium JRE 8 hoac sua duong dan trong file nay.
    pause
    exit /b 1
)

set OUTPUT_DIR=%~dp0..\dist\jre_mini

REM Xoa thu muc jre cu neu co de tap moi
if exist "%OUTPUT_DIR%" (
    echo [INFO] Phat hien ban build JRE cu. Dang xoa thu muc %OUTPUT_DIR%...
    rmdir /S /Q "%OUTPUT_DIR%"
)

echo [INFO] Dang copy JRE tu %JAVA8_HOME% den %OUTPUT_DIR%...
mkdir "%OUTPUT_DIR%"
xcopy "%JAVA8_HOME%" "%OUTPUT_DIR%" /E /I /H /C /Q

if %ERRORLEVEL% EQU 0 (
    echo [SUCCESS] Copy JRE 8 thanh cong tai %OUTPUT_DIR%
) else (
    echo [ERROR] Qua trinh copy JRE that bai.
)
pause
