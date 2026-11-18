@echo off
chcp 65001 >nul
title ĐẨY DỰ ÁN LÊN GITHUB ĐỂ DEPLOY STREAMLIT CLOUD
cls
echo ======================================================================
echo       TỰ ĐỘNG ĐẨY DỰ ÁN LÊN GITHUB ĐỂ CHẠY STREAMLIT CLOUD (24/7)
echo ======================================================================
echo.
echo Bước 1: Hãy vào https://github.com/new để tạo 1 Repository mới.
echo         (Ví dụ đặt tên là: mrir-web-app, để chế độ Public hoặc Private)
echo.
echo Bước 2: Dán đường dẫn HTTPS của Repository vừa tạo vào dưới đây.
echo         (Ví dụ: https://github.com/username/mrir-web-app.git)
echo ----------------------------------------------------------------------
echo.

set /p REPO_URL="Nhập URL GitHub Repository: "

if "%REPO_URL%"=="" (
    echo [!] URL không được để trống!
    pause
    exit /b
)

echo.
echo [*] Đang cấu hình nhánh chính 'main'...
git branch -M main

echo [*] Đang liên kết với remote GitHub: %REPO_URL%
git remote remove origin 2>nul
git remote add origin %REPO_URL%

echo [*] Đang tải mã nguồn lên GitHub...
git push -u origin main

if %ERRORLEVEL% equ 0 (
    echo.
    echo ======================================================================
    echo [THÀNH CÔNG] Đã đẩy mã nguồn lên GitHub thành công!
    echo.
    echo Bước tiếp theo:
    echo 1. Truy cập: https://share.streamlit.io
    echo 2. Đăng nhập bằng tài khoản GitHub của bạn.
    echo 3. Nhấn nút "New app" và chọn Repository bạn vừa tải lên.
    echo 4. Mục "Main file path" điền: app_web.py
    echo 5. Nhấn "Deploy"! Sau 2-3 phút, Tool sẽ chạy Online 24/7 vĩnh viễn!
    echo ======================================================================
) else (
    echo.
    echo [!] Đẩy lên GitHub không thành công. Hãy kiểm tra lại quyền đăng nhập GitHub.
)

pause
