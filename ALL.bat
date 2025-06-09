@echo off
cd /d D:\SynologyDrive\flask

:: 執行 Excel_Edge.py
python "Excel_Edge.py"
if errorlevel 1 goto error

pause

:: 執行 run_update2.py
python "run_update2.py"
if errorlevel 1 goto error

pause

:: 執行 data_updw.py 並啟動 save_excel.exe
python "data_updw.py"
if errorlevel 1 goto error

start "" "save_excel.exe"
pause

:: 進入 Git 專案資料夾
cd /d D:\SynologyDrive\flask

:: 拉取遠端最新版本
git pull
if errorlevel 1 goto error

:: 加入所有變更
git add -A

:: 建立時間戳記
for /f "tokens=1-3 delims=/ " %%a in ("%date%") do (
    set y=%%c
    set m=%%a
    set d=%%b
)
set timestamp=%y%-%m%-%d_%time%
git commit -m "Auto commit on %timestamp%"
if errorlevel 1 goto error

git push
if errorlevel 1 goto error

pause
exit

:error
echo 發生錯誤，請檢查程式執行狀況。
pause
exit /b
