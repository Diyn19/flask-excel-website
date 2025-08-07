@echo off
setlocal enabledelayedexpansion

cd /d D:\flask

:: 新增功能：詢問是否直接開始 Git 操作
set /p skipAll=是否直接開始 Git 操作 [Y/N]：
if /i "!skipAll!"=="Y" (
    goto git_only
)

:: 原本詢問是否跳過下載檔案
set /p skipDownload=是否跳過下載檔案 [Y/N]：

if /i "!skipDownload!"=="Y" (
    echo [1/6] 已選擇跳過下載檔案。
) else (
    echo [1/6] 執行 Excel_Edge.py...
    python "Excel_Edge.py"
    if errorlevel 1 goto error
)

echo [2/6] 執行 run_update2.py...
python "run_update2.py"
if errorlevel 1 goto error

echo [3/6] 執行 run_MFP_update.py...
python "run_MFP_update.py"
if errorlevel 1 goto error

:: 第 4 步：輸入版本號（允許直接 Enter 自動填入）
set /p vernum=請輸入版本號（直接 Enter 則使用目前時間 mmddhhmm）：
if "!vernum!"=="" (
    for /f %%a in ('powershell -command "Get-Date -Format MMddHHmm"') do set vernum=%%a
)

echo [4/6] 寫入版本號 %vernum% 到 Excel...
python "add_ver.py" %vernum%
if errorlevel 1 goto error

echo [5/6] 執行 data_updw.py...
python "data_updw.py"
if errorlevel 1 goto error

echo [6/6] 啟動 save_excel.exe...
start /wait "" "save_excel.exe"
if errorlevel 1 goto error

:: 執行 Git 操作
goto git_operation

:git_only
:: 只執行 Git，也需輸入版本號
set /p vernum=請輸入版本號（將寫入 Excel 並作為 Git 訊息，Enter 自動填入）：
if "!vernum!"=="" (
    for /f %%a in ('powershell -command "Get-Date -Format MMddHHmm"') do set vernum=%%a
)

echo [1/3] 寫入版本號 %vernum% 到 Excel...
python "add_ver.py" %vernum%
if errorlevel 1 goto error

echo [2/3] 儲存 Excel（save_excel.exe）...
start /wait "" "save_excel.exe"
if errorlevel 1 goto error

echo [3/3] 進行 Git 操作...
cd /d D:\flask

git pull
if errorlevel 1 goto error

git add -A
git commit -m "Auto commit - %vernum%"
if errorlevel 1 goto error

git push
if errorlevel 1 goto error

echo 所有程序已完成！
pause
exit /b

:git_operation
echo [Git] 進行 Git 操作...
cd /d D:\flask

git pull
if errorlevel 1 goto error

git add -A
git commit -m "Auto commit - %vernum%"
if errorlevel 1 goto error

git push
if errorlevel 1 goto error

echo 所有程序已完成！
pause
exit /b

:error
echo 錯誤：某個步驟執行失敗，流程中止。
pause
exit /b
