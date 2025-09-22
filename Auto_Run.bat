@echo off
setlocal enabledelayedexpansion

cd /d D:\flask

:: 詢問是否直接開始 Git 操作
set /p skipAll=是否直接開始 Git 操作 [Y/N]：
if /i "!skipAll!"=="Y" (
    goto git_only
)

:: 詢問是否跳過下載檔案
set /p skipDownload=是否跳過下載檔案 [Y/N]：
if /i "!skipDownload!"=="Y" (
    echo [1/7] 已選擇跳過下載檔案。
) else (
    echo [1/7] 執行 Excel_Edge.py...
    python "Excel_Edge.py"
    if errorlevel 1 (
        echo 檔案下載異常
        pause
        exit /b
    )
)

:: [2/7] 複製 data.xlsx 到臨時備份
echo [2/7] 複製 data.xlsx 到臨時備份資料夾...
copy /Y "data.xlsx" "臨時備份\自動備份\data.xlsx"
if errorlevel 1 goto error

:: [3/7] 執行 run_update2.py
echo [3/7] 執行 run_update2.py...
python "run_update2.py"
if errorlevel 1 goto error

:: [4/7] 執行 run_MFP_update.py
echo [4/7] 執行 run_MFP_update.py...
python "run_MFP_update.py"
if errorlevel 1 goto error

:: [5/7] 執行 add_ver.py，自動或手動版本號
echo [5/7] 執行 add_ver.py（版本號可手動輸入，10 秒後自動填入）...
for /f %%a in ('python "add_ver.py"') do set vernum=%%a
echo 使用版本號: %vernum%

:: [6/7] 執行 data_updw.py
echo [6/7] 執行 data_updw.py...
python "data_updw.py"
if errorlevel 1 goto error

:: [7/7] 執行 save_excel.exe
echo [7/7] 啟動 save_excel.exe...
start /wait "" "save_excel.exe"
if errorlevel 1 goto error

:: 執行 Git 操作
goto git_operation

:git_only
:: 只執行 Git，也需取得版本號
for /f %%a in ('python "add_ver.py"') do set vernum=%%a
echo 使用版本號: %vernum%

echo [1/3] 寫入版本號到 Excel（已由 add_ver.py 寫入）
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
