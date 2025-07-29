@echo off
setlocal enabledelayedexpansion

cd /d D:\flask

:: �s�W�\��G�߰ݬO�_�����}�l Git �ާ@
set /p skipAll=�O�_�����}�l Git �ާ@ [Y/N]�G
if /i "!skipAll!"=="Y" (
    goto git_operation
)

:: �쥻�߰ݬO�_���L�U���ɮ�
set /p skipDownload=�O�_���L�U���ɮ� [Y/N]�G

if /i "!skipDownload!"=="Y" (
    echo [1/6] �w��ܸ��L�U���ɮסC
) else (
    echo [1/6] ���� Excel_Edge.py...
    python "Excel_Edge.py"
    if errorlevel 1 goto error
)

echo [2/6] ���� run_update2.py...
python "run_update2.py"
if errorlevel 1 goto error

echo [3/6] ���� run_MFP_update.py...
python "run_MFP_update.py"
if errorlevel 1 goto error

:: �쥻���B�J 3/6 �{�b�אּ�� 4/6�G��J�������ð��� add_ver.py
set /p vernum=�п�J�������G
echo [4/6] �g�J������ %vernum% �� Excel...
python "add_ver.py" %vernum%
if errorlevel 1 goto error

echo [5/6] ���� data_updw.py...
python "data_updw.py"
if errorlevel 1 goto error

echo [6/6] �Ұ� save_excel.exe...
start /wait "" "save_excel.exe"
if errorlevel 1 goto error

:git_operation
echo [Git] �i�� Git �ާ@...
cd /d D:\flask

git pull
if errorlevel 1 goto error

git add -A

:: �إ� timestamp�]�榡�GYYYY-MM-DD_HH:MM:SS�^
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

echo �Ҧ��{�Ǥw�����I
pause
exit /b

:error
echo ���~�G�Y�ӨB�J���楢�ѡA�y�{����C
pause
exit /b
