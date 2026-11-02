@echo off
REM 一键打包：隐藏 CMD 控制台窗口（--windowed），产物 dist/设计周期定额计算_1.2.0.exe
chcp 65001 >nul
cd /d "%~dp0"
pip install -r requirements.txt
pyinstaller --noconfirm --clean --onefile --windowed --name "设计周期定额计算_1.2.0" --paths src src\app.py
echo.
echo 打包完成: dist\设计周期定额计算_1.2.0.exe
pause
