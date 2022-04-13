if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
set "cmd=cscript //nologo ospp.vbs"
%cmd% /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul 2>&1
 
%cmd% /dstatus | findstr "Office19ProPlus2019VL"
 
if not %errorlevel% == 0 (for /f %x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do %cmd% /inslic:"..\root\Licenses16\%x")
%cmd% /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
 
%cmd% /sethst:kms.lotro.cc & %cmd% /act
 
cls & %cmd% /dstatus
echo