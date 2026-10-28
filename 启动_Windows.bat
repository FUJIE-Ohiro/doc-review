@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON=py -3"
) else (
    set "PYTHON=python"
)
%PYTHON% "%~dp0word_audit.py" --gui
if not errorlevel 1 goto end

echo.
echo [提示] 图形界面无法启动。可改用命令行：
echo 请将稿件文件夹拖入下方窗口，按 Enter。
set /p "INPUT_DIR=稿件文件夹路径: "
if not defined INPUT_DIR goto end
%PYTHON% "%~dp0word_audit.py" --input "%INPUT_DIR%"
if errorlevel 1 echo 运行失败，请检查 Python 版本或路径。
:end
echo.
pause