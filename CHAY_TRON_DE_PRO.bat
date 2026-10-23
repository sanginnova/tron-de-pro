@echo off
chcp 65001 >nul
title HỆ THỐNG TRỘN ĐỀ THI TRẮC NGHIỆM PRO - THS. NGUYỄN VĂN SANG
color 1F

echo ======================================================================
echo          CHƯƠNG TRÌNH TRỘN ĐỀ THI TRẮC NGHIỆM PRO (GDPT 2018)
echo      Tác giả: ThS. Nguyễn Văn Sang - Trường Cao đẳng Nghề số 1 - BQP
echo ======================================================================
echo.

set /p INPUT_DOCX="[+] Kéo thả file Word đề thi nguồn (.docx) vào đây rồi nhấn Enter: "
set INPUT_DOCX=%INPUT_DOCX:"=%

if "%INPUT_DOCX%"=="" (
    echo [!] Lỗi: Bạn chưa nhập đường dẫn file đề!
    pause
    exit /b
)

set /p NUM_CODES="[+] Nhập số lượng mã đề cần tạo (Mặc định 4): "
if "%NUM_CODES%"=="" set NUM_CODES=4

set /p START_CODE="[+] Nhập mã đề bắt đầu (Mặc định 101): "
if "%START_CODE%"=="" set START_CODE=101

set /p EXAM_TIME="[+] Nhập thời gian làm bài tính bằng phút (Mặc định 45): "
if "%EXAM_TIME%"=="" set EXAM_TIME=45

echo.
echo [*] Đang xử lý trộn đề chuyên nghiệp... Vui lòng chờ...
echo.

python "%~dp0cli.py" "%INPUT_DOCX%" --codes %NUM_CODES% --start %START_CODE% --time %EXAM_TIME%

echo.
echo ======================================================================
echo [*] Đã hoàn thành! Nhấn phím bất kỳ để thoát...
echo ======================================================================
pause >nul
