if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" (cd /d "%ProgramFiles%\Microsoft Office\Office16")
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" (cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")
set k1=KBQNC-JMVHB-WTMKW-F2G48-4VXG3
cls
@echo on&mode con: cols=20 lines=2
cscript ospp.vbs /inpkey:%k1%
@mode con: cols=100 lines=30
cscript ospp.vbs /dinstid>id.txt 
start id.txt