﻿chcp 65001 >nul
@echo off
Title XOA KEY OFFICE
mode con: cols=96 lines=35
chcp 65001 >nul
@echo.
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo  Run CMD as Administrator...
    goto goUAC 
) else (
 goto goADMIN )

:goUAC
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    set params = %*:"=""
    echo UAC.ShellExecute "cmd.exe", "/c %~s0 %params%", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /B

:goADMIN
    pushd "%CD%"
    CD /D "%~dp0"
	
:main
cls
color f0
@echo. 
echo        XOA KEY OFFICE
echo     Chon Phien Ban Office Can Xoa Key
echo =========================================
echo [  1. Office 2010     : Nhan phim so 1  ]
echo [  2. Office 2013     : Nhan phim so 2  ]
echo [  3. Office 2016     : Nhan phim so 3  ]
echo [  4. Office 2019     : Nhan phim so 4  ]
echo [  5. Office 2021     : Nhan phim so 5  ]
echo [  6. Office 365      : Nhan phim so 6  ]
echo =========================================
Choice /N /C 12345 /M "* Nhap lua chon : 
if %errorlevel% == 6 ( set "xx=16" & goto vogia)
if %errorlevel% == 5 ( set "xx=16" & goto vogia)
if %errorlevel% == 4 ( set "xx=16" & goto vogia)
if %errorlevel% == 3 ( set "xx=16" & goto vogia)
if %errorlevel% == 2 ( set "xx=15" & goto vogia)
if %errorlevel% == 1 ( set "xx=14" & goto vogia)

:vogia
if exist "%ProgramFiles%\Microsoft Office\Office%xx%\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office%xx%"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office%xx%\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office%xx%"
cscript ospp.vbs /dstatus >dstatus.txt
start dstatus.txt
goto office
)

:office
set /p key= * NHAP 5 KY TU CUOI CUA KEY : 
@echo  ...DANG XOA KEY OFFICE...
cscript OSPP.VBS /unpkey:%key%
@echo =========================================
@echo      DA XOA KEY OFFICE THANH CONG !
@echo =========================================
goto office
)