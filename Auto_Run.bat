@echo off
setlocal enabledelayedexpansion

cd /d D:\flask

:: �s�W�\��G�߰ݬO�_�����}�l Git �ާ@
set /p skipAll=�O�_�����}�l Git �ާ@ [Y/N]�G
if /i "!skipAll!"=="Y" (
    goto git_only
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

:: �� 4 �B�G��J�������]���\���� Enter �۰ʶ�J�^
set /p vernum=�п�J�������]���� Enter �h�ϥΥثe�ɶ� mmddhhmm�^�G
if "!vernum!"=="" (
    for /f %%a in ('powershell -command "Get-Date -Format MMddHHmm"') do set vernum=%%a
)

echo [4/6] �g�J������ %vernum% �� Excel...
python "add_ver.py" %vernum%
if errorlevel 1 goto error

echo [5/6] ���� data_updw.py...
python "data_updw.py"
if errorlevel 1 goto error

echo [6/6] �Ұ� save_excel.exe...
start /wait "" "save_excel.exe"
if errorlevel 1 goto error

:: ���� Git �ާ@
goto git_operation

:git_only
:: �u���� Git�A�]�ݿ�J������
set /p vernum=�п�J�������]�N�g�J Excel �ç@�� Git �T���AEnter �۰ʶ�J�^�G
if "!vernum!"=="" (
    for /f %%a in ('powershell -command "Get-Date -Format MMddHHmm"') do set vernum=%%a
)

echo [1/3] �g�J������ %vernum% �� Excel...
python "add_ver.py" %vernum%
if errorlevel 1 goto error

echo [2/3] �x�s Excel�]save_excel.exe�^...
start /wait "" "save_excel.exe"
if errorlevel 1 goto error

echo [3/3] �i�� Git �ާ@...
cd /d D:\flask

git pull
if errorlevel 1 goto error

git add -A
git commit -m "Auto commit - %vernum%"
if errorlevel 1 goto error

git push
if errorlevel 1 goto error

echo �Ҧ��{�Ǥw�����I
pause
exit /b

:git_operation
echo [Git] �i�� Git �ާ@...
cd /d D:\flask

git pull
if errorlevel 1 goto error

git add -A
git commit -m "Auto commit - %vernum%"
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
