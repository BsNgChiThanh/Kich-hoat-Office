# Kich-hoat-Office
Kích hoạt Office 2010, 2013, 2016, 2019, 2021, 365

**[Source nguồn Office](https://docs.google.com/spreadsheets/d/e/2PACX-1vRlK-vRwPJHDaANT81EjyG4m5ZnLXdKRYfS0eKXyCzGymEfUDmKHRhxvUbtWYTfVn7MJ3E2jk7v3cGi/pubhtml?gid=605361024&single=true)**. **Password giải nén nếu có của nguồn này [bấm vào đây](https://docs.google.com/document/d/1nskNEcAVu1SbhSzdRfGQWwT3aYtKUzLN/edit?usp=drivesdk&ouid=108710666609351868901&rtpof=true&sd=true)**

**Nếu bạn muốn tìm key windows và office thì [bấm vào đây](https://t.me/+yqfFsJPOciwyY2Vl)**

# Kích hoạt Office bằng cmd!!!
 
## Chạy kích hoạt này sẽ được 180 ngày sử dụng, gần hết thì các bạn chạy kính hoạt này một lần nữa sẽ được 180 ngày và hãy lập lại như thế sẽ xem như vĩnh viễn ##

## Nếu bạn hay quên, hãy cho file cmd của mình tạo ra khởi động cùng Windows, khi đó mỗi lần khởi động máy thì file cmd này tự động chạy đồng nghĩa với office của mình mới được active lại và thời hạn sử dụng cũng đồng nghĩa với tự động gia hạn, để làm đều này bạn xem hướng dẫn sau đây: ##

Hãy nhấn phím logo **Windows + R**, nhập **shell:startup**, rồi chọn OK. Thao tác này sẽ mở thư mục Khởi động.

Sao chép và dán file cmd của mình vào đấy là Ok

# 1. Office 365 Prolus #

Mở Notepad lên dán đoạn mã dưới đây vào, bấm save as và lưu tên **kichhoatoffice365Prolus.cmd** sau đó run file này bằng quyền administrator là OK.

```php
@echo off
title Activate Office 365 ProPlus for FREE - MSGuides.com&cls&echo =====================================================================================&echo #Project: Activating Microsoft software products for FREE without additional software&echo =====================================================================================&echo.&echo #Supported products: Office 365 ProPlus (x86-x64)&echo.&echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&echo.&echo ============================================================================&echo Activating your Office...&cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:WFG99 >nul&cscript //nologo ospp.vbs /unpkey:DRTFM >nul&cscript //nologo ospp.vbs /unpkey:BTDRB >nul&set i=1&cscript //nologo ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 >nul||cscript //nologo ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul||goto notsupported
:skms
if %i% GTR 10 goto busy
if %i% EQU 1 set KMS=kms7.MSGuides.com
if %i% EQU 2 set KMS=s8.uk.to
if %i% EQU 3 set KMS=sv9.uk.to
if %i% GTR 3 goto ato
cscript //nologo ospp.vbs /sethst:%KMS% >nul
:ato
echo ============================================================================&echo.&echo.&cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo ============================================================================&echo.&echo #My official blog: MSGuides.com&echo.&echo #How it works: bit.ly/kms-server&echo.&echo #Please feel free to contact me at msguides.com@gmail.com if you have any questions or concerns.&echo.&echo #Please consider supporting this project: donate.msguides.com&echo #Your support is helping me keep my servers running 24/7!&echo.&echo ============================================================================&choice /n /c YN /m "Would you like to visit my blog [Y,N]?" & if errorlevel 2 exit) || (echo The connection to my KMS server failed! Trying to connect to another one... & echo Please wait... & echo. & echo. & set /a i+=1 & goto skms)
explorer "http://MSGuides.com"&goto halt
:notsupported
echo ============================================================================&echo.&echo Sorry, your version is not supported.&echo.&goto halt
:busy
echo ============================================================================&echo.&echo Sorry, the server is busy and can't respond to your request. Please try again.&echo.
:halt
pause >nul
```

# 2. Office 365 Mondo 2016 #

Mở Notepad lên dán đoạn mã dưới đây vào, bấm save as và lưu tên **kichhoatoffice365Mondo.cmd** sau đó run file này bằng quyền administrator là OK.

```php
@echo off
title Activate Office 365 ProPlus for FREE - MSGuides.com&cls&echo ============================================================================&echo BS. NGUYEN CHI THANH, TRUONG KHOA CAP CUU BV DAM DOI, Phone 0914678254 &echo ============================================================================&echo.&echo Bs Nguyen Chi Thanh, Activating Office 365 Mondo 2016(x86-x64)&echo.&echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&echo.&echo ============================================================================&echo Activating your Office...&cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:WFG99 >nul&cscript //nologo ospp.vbs /unpkey:DRTFM >nul&cscript //nologo ospp.vbs /unpkey:BTDRB >nul&cscript //nologo ospp.vbs /inpkey:HFTND-W9MK4-8B7MJ-B6C4G-XQBR2 >nul&set i=1
:server
if %i%==1 set KMS=kms7.MSGuides.com
if %i%==2 set KMS=kms8.MSGuides.com
if %i%==3 set KMS=kms9.MSGuides.com
if %i%==4 goto notsupported
cscript //nologo ospp.vbs /sethst:%KMS% >nul&echo ============================================================================&echo.&echo.
cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo ============================================================================&echo.&echo My blog: bsnguyenchithanh.business.site Phone 0914678254.&echo Hay ung ho Phong kham Noi tong hop tu nhan cua toi. &echo Rat vinh du duoc don tiep qui khach. &echo Chuc qui khach Van su nhu y&echo.&echo ============================================================================&choice /n /c YN /m "Would you like to visit blog of Bs Nguyen Chi Thanh, BV Da Khoa Dam Doi [Y,N]?" & if errorlevel 2 exit) || (echo The connection to my KMS server failed! Trying to connect to another one... & echo Please wait... & echo. & echo. & set /a i+=1 & goto server)
explorer "http://bsnguyenchithanh.business.site"&goto halt
:notsupported
echo.&echo ============================================================================&echo Sorry! Your version is not supported.&echo Please try installing the latest version here: bit.ly/odt2k16
:halt
pause >nul
```

# 3. Office 365 Enterprise #

Bạn dùng tài khoản để kích hoạt nhé! [Bấm vào đây để lấy tài khoản](https://bsthanh-my.sharepoint.com/:w:/g/personal/laptopxiaomi_bsthanh_tk/EQa9vlOr8JdOqcUEYGyjjfQBvW7eHmeqtjR1KMf__A2lHw?e=YgQkSj)

**Tham khảo thêm về office 365 E5 _ Office 365 Enterprise [bấm vào đây](https://github.com/BsNgChiThanh/Tao-office-365-E5-kich-hoat-Office-365-for-desktop)**

# 4. Office 2021 #

Mở Notepad lên dán đoạn mã dưới đây vào, bấm save as và lưu tên **kichhoatoffice2021.cmd** sau đó run file này bằng quyền administrator là OK.

```php
@echo off
title Activate Microsoft Office 2021 (ALL versions) for FREE - MSGuides.com&cls&echo =====================================================================================&echo #Project: Activating Microsoft software products for FREE without additional software&echo =====================================================================================&echo.&echo #Supported products:&echo - Microsoft Office Standard 2021&echo - Microsoft Office Professional Plus 2021&echo.&echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&echo.&echo =====================================================================================&echo Activating your product...&cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:6F7TH >nul&set i=1&cscript //nologo ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH >nul||goto notsupported
:skms
if %i% GTR 10 goto busy
if %i% EQU 1 set KMS=kms7.MSGuides.com
if %i% EQU 2 set KMS=s8.uk.to
if %i% EQU 3 set KMS=s9.us.to
if %i% GTR 3 goto ato
cscript //nologo ospp.vbs /sethst:%KMS% >nul
:ato
echo =====================================================================================&echo.&echo.&cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo =====================================================================================&echo.&echo #My official blog: MSGuides.com&echo.&echo #How it works: bit.ly/kms-server&echo.&echo #Please feel free to contact me at msguides.com@gmail.com if you have any questions or concerns.&echo.&echo #Please consider supporting this project: donate.msguides.com&echo #Your support is helping me keep my servers running 24/7!&echo.&echo =====================================================================================&choice /n /c YN /m "Would you like to visit my blog [Y,N]?" & if errorlevel 2 exit) || (echo The connection to my KMS server failed! Trying to connect to another one... & echo Please wait... & echo. & echo. & set /a i+=1 & goto skms)
explorer "http://MSGuides.com"&goto halt
:notsupported
echo =====================================================================================&echo.&echo Sorry, your version is not supported.&echo.&goto halt
:busy
echo =====================================================================================&echo.&echo Sorry, the server is busy and can't respond to your request. Please try again.&echo.
:halt
pause >nul
```

# 5. Office 2019 #

Mở Notepad lên dán đoạn mã dưới đây vào, bấm save as và lưu tên **kichhoatoffice2019.cmd** sau đó run file này bằng quyền administrator là OK.

```php
@echo off  title Kich hoat Microsoft Office 2019 ALL versions mienphi!&cls&echo ============================================================================&echo #Kich hoat Microsoft Office 2019 hop phap - Khong su dung phan mem&echo ============================================================================&echo.&echo #San pham ho tro:&echo - Microsoft Office Standard 2019&echo - Microsoft Office Professional Plus 2019&echo.&echo.&(if exist "%ProgramFiles%Microsoft OfficeOffice16ospp.vbs" cd /d "%ProgramFiles%Microsoft OfficeOffice16")&(if exist "%ProgramFiles(x86)%Microsoft OfficeOffice16ospp.vbs" cd /d "%ProgramFiles(x86)%Microsoft OfficeOffice16")&(for /f %%x in ('dir /b ..rootLicenses16ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..rootLicenses16%%x" >nul)&(for /f %%x in ('dir /b ..rootLicenses16ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..rootLicenses16%%x" >nul)&echo.&echo ============================================================================&echo Activating your Office...&cscript //nologo ospp.vbs /unpkey:6MWKP >nul&cscript //nologo ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul&set i=1  :server  if %i%==1 set KMS_Sev=kms7.MSGuides.com  if %i%==2 set KMS_Sev=kms8.MSGuides.com  if %i%==3 set KMS_Sev=kms9.MSGuides.com  if %i%==4 goto notsupported  cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul&echo ============================================================================&echo.&echo.  cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo ============================================================================&choice /n /c YN /m "Truy cap topthuthuat.vn: [Y,N]?" & if errorlevel 2 exit) || (echo Ket noi voi may chu KMS khong thanh cong! Dang ket noi lai... & echo Vui long cho... & echo. & echo. & set /a i+=1 & goto server)  explorer "http://topthuthuat.vn"&goto halt  :notsupported  echo.&echo ============================================================================&echo Phien ban Office cua ban khong duoc ho tro.&echo Download phien ban moi nhat tai day: http://topthuthuat.vn/:halt  pause >nul
```

Hoặc:

```php
@echo off
title Bs Nguyen Chi Thanh, Kich hoat Microsoft Office 2019 ALL versions!&cls&echo ==========================Bs Nguyen Chi Thanh======================================&echo # Bs Nguyen Chi Thanh, Khoa CC_HSTC_CD BV Dam Doi Kich hoat Microsoft Office 2019&echo ==========================Bs Nguyen Chi Thanh======================================&echo.&echo #San pham ho tro:&echo - Microsoft Office Standard 2019&echo - Microsoft Office Professional Plus 2019&echo.&echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&echo.&echo ==========================Bs Nguyen Chi Thanh======================================&echo Activating your Office...&cscript //nologo ospp.vbs /unpkey:6MWKP >nul&cscript //nologo ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul&set i=1
:server
if %i%==1 set KMS_Sev=kms7.MSGuides.com
if %i%==2 set KMS_Sev=kms8.MSGuides.com
if %i%==3 set KMS_Sev=kms9.MSGuides.com
if %i%==4 goto notsupported
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul&echo ==========================Bs Nguyen Chi Thanh======================================&echo.&echo.
cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo ==========================Bs Nguyen Chi Thanh======================================&choice /n /c YN /m "Truy cap trang Web Bs Nguyen Chi Thanh: [Y,N]?" & if errorlevel 2 exit) || (echo Ket noi voi may chu KMS khong thanh cong! Dang ket noi lai... & echo Vui long cho... & echo. & echo. & set /a i+=1 & goto server)
explorer "https://phong-kham-bsck1-nguyen-chi-thanh.business.site/?m=true"&goto halt
:notsupported
echo.&echo ==========================Bs Nguyen Chi Thanh======================================&echo Phien ban Office cua ban khong duoc ho tro.&echo Download phien ban moi nhat tai day: http://topthuthuat.vn/:halt
pause >nul
```

**Office 2019 Prolus:**

Mở Command Prompt bằng quyền Run Administrator (Tức bấm tìm kiếm Command Prompt, bấm chuột phải chọn Run Administrator)

![1](https://user-images.githubusercontent.com/82578024/168939483-8b9dd175-677f-4493-9ed0-f6731ce10b40.gif)

Dán đoạn mã sau vào:

```php
if exist “%ProgramFiles%\Microsoft Office\Office16\ospp.vbs” cd /d “%ProgramFiles%\Microsoft Office\Office16” if exist “%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs” cd /d “%ProgramFiles(x86)%\Microsoft Office\Office16” set “cmd=cscript //nologo ospp.vbs” %cmd% /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul 2>&1 %cmd% /dstatus | findstr “Office19ProPlus2019VL” if not %errorlevel% == 0 (for /f %x in (‘dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms’) do %cmd% /inslic:”..\root\Licenses16\%x”) %cmd% /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP %cmd% /sethst:kms.lotro.cc & %cmd% /act cls & %cmd% /dstatus echo.
```

# 6. Office 2016 #

Mở Notepad lên dán đoạn mã dưới đây vào, bấm save as và lưu tên **kichhoatoffice2016.cmd** sau đó run file này bằng quyền administrator là OK.

```php
@echo off
title Activate Microsoft Office 2016 ALL versions for FREE!&cls&echo ============================================================================&echo #Project: Activating Microsoft software products for FREE without software&echo ============================================================================&echo.&echo #Supported products:&echo - Microsoft Office Standard 2016&echo - Microsoft Office Professional Plus 2016&echo.&echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&echo.&echo ============================================================================&echo Activating your Office...&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:WFG99 >nul&cscript //nologo ospp.vbs /unpkey:DRTFM >nul&cscript //nologo ospp.vbs /unpkey:BTDRB >nul&cscript //nologo ospp.vbs /unpkey:CPQVG >nul&cscript //nologo ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 >nul&set i=1
:server
if %i%==1 set KMS=kms7.MSGuides.com
if %i%==2 set KMS=kms8.MSGuides.com
if %i%==3 set KMS=kms9.MSGuides.com
if %i%==4 goto notsupported
cscript //nologo ospp.vbs /sethst:%KMS% >nul&echo ============================================================================&echo.&echo.
cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo ============================================================================&echo.&echo #My official blog: MSGuides.com&echo.&echo #How it works: bit.ly/kms-server&echo.&echo #Please feel free to contact me at msguides.com@gmail.com if you have any questions or concerns.&echo.&echo #Please consider supporting this project: donate.msguides.com&echo #Your support is helping me keep my servers running everyday!&echo.&echo ============================================================================&choice /n /c YN /m "Would you like to visit my blog [Y,N]?" & if errorlevel 2 exit) || (echo The connection to my KMS server failed! Trying to connect to another one... & echo Please wait... & echo. & echo. & set /a i+=1 & goto server)
explorer "http://MSGuides.com"&goto halt
:notsupported
echo.&echo ============================================================================&echo Sorry! Your version is not supported.&echo Please try installing the latest version here: bit.ly/downloadmsp
:stop
pause> null
```

Hoặc:

Mở Command Prompt bằng quyền Run Administrator (Tức bấm tìm kiếm Command Prompt, bấm chuột phải chọn Run Administrator)

![1](https://user-images.githubusercontent.com/82578024/168939483-8b9dd175-677f-4493-9ed0-f6731ce10b40.gif)

Dán đoạn mã sau vào:

```php
set ver=16
if exist “%ProgramFiles%\Microsoft Office\Office%ver%\ospp.vbs” cd /d “%ProgramFiles%\Microsoft Office\Office%ver%”
if exist “%ProgramFiles(x86)%\Microsoft Office\Office%ver%\ospp.vbs” cd /d “%ProgramFiles(x86)%\Microsoft Office\Office%ver%”
for /f “tokens=8” %b in (‘cscript //nologo OSPP.VBS /dstatus ^| findstr /b /c:”Last 5″‘) do (cscript //nologo ospp.vbs /unpkey:%b)
for /f %i in (‘dir /b ..\root\Licenses%ver%\ProPlusVL_KMS*.xrm-ms’) do cscript ospp.vbs /inslic:”..\root\Licenses%ver%\%i”
cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
cscript ospp.vbs /sethst:kms.lotro.cc
cscript ospp.vbs /act
Start winword
@
```

Hoặc đoạn mã:

```php
cscript slmgr.vbs /skms kms.digiboy.ir
cscript slmgr.vbs /ipk XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
cscript slmgr.vbs /ato d450596f-894d-49e0-966a-fd39ed4c4c64
timeout 2&start winword&exit
```

# 7. Office 2013 #

Mở Notepad lên dán đoạn mã dưới đây vào, bấm save as và lưu tên **kichhoatoffice2013.cmd** sau đó run file này bằng quyền administrator là OK.

```php
@echo off
title Activate Microsoft Office 2013 Volume for FREE!&cls&echo ============================================================================&echo #Project: Activating Microsoft software products for FREE without software&echo ============================================================================&echo.&echo #Supported products:&echo - Microsoft Office 2013 Standard Volume&echo - Microsoft Office 2013 Professional Plus Volume&echo.&echo.&(if exist "%ProgramFiles%\Microsoft Office\Office15\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office15")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office15")&echo.&echo ============================================================================&echo Activating your Office...&cscript //nologo ospp.vbs /unpkey:92CD4 >nul&cscript //nologo ospp.vbs /unpkey:GVGXT >nul&cscript //nologo ospp.vbs /inpkey:KBKQT-2NMXY-JJWGP-M62JB-92CD4 >nul&cscript //nologo ospp.vbs /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT >nul&set i=1
:server
if %i%==1 set KMS_Sev=kms7.MSGuides.com
if %i%==2 set KMS_Sev=kms8.MSGuides.com
if %i%==3 set KMS_Sev=kms9.MSGuides.com
if %i%==4 goto notsupported
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul&echo ============================================================================&echo.&echo.
cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo ============================================================================&echo.&echo #My official blog: MSGuides.com&echo.&echo #How it works: bit.ly/kms-server&echo.&echo #Please feel free to contact me at msguides.com@gmail.com if you have any questions or concerns.&echo.&echo #Please consider supporting this project: donate.msguides.com&echo #Your support is helping me keep my servers running everyday!&echo.&echo ============================================================================&choice /n /c YN /m "Would you like to visit my blog [Y,N]?" & if errorlevel 2 exit) || (echo The connection to my KMS server failed! Trying to connect to another one... & echo Please wait... & echo. & echo. & set /a i+=1 & goto server)
explorer "http://MSGuides.com"&goto halt
:notsupported
echo.&echo ============================================================================&echo Sorry! Your version is not supported.
:halt
pause >nul
```

Mở Command Prompt bằng quyền Run Administrator (Tức bấm tìm kiếm Command Prompt, bấm chuột phải chọn Run Administrator)

![1](https://user-images.githubusercontent.com/82578024/168939483-8b9dd175-677f-4493-9ed0-f6731ce10b40.gif)

Dán đoạn mã sau vào:

```php
if exist “%ProgramFiles%\Microsoft Office\Office15\ospp.vbs” cd /d “%ProgramFiles%\Microsoft Office\Office15”
if exist “%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs” cd /d “%ProgramFiles(x86)%\Microsoft Office\Office15”
cscript OSPP.VBS /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT
cscript OSPP.VBS /inpkey:FN8TT-7WMH6-2D4X9-M337T-2342K
cscript OSPP.VBS /inpkey:KBKQT-2NMXY-JJWGP-M62JB-92CD4
cscript OSPP.VBS /inpkey:C2FG9-N6J68-H8BTJ-BW3QX-RM3B3
cscript ospp.vbs /sethst:kms.lotro.cc
cscript ospp.vbs /act
```

Hoặc đoạn mã:

```php
if exist “%ProgramFiles%\Microsoft Office\Office15\ospp.vbs” cd /d “%ProgramFiles%\Microsoft Office\Office15”
if exist “%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs” cd /d “%ProgramFiles(x86)%\Microsoft Office\Office15”
cscript //nologo OSPP.VBS /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT
cscript //nologo ospp.vbs /sethst:kms.lotro.cc&cscript //nologo ospp.vbs /act&timeout 5&start winword&exit
@
```

# 8.Office 2010 #

Mở Notepad lên dán đoạn mã dưới đây vào, bấm save as và lưu tên **kichhoatoffice2010.cmd** sau đó run file này bằng quyền administrator là OK.

```php
@echo off title Activate Microsoft Office 2010 Volume for FREE!&cls&echo ============================================================================&echo #Project: Activating Microsoft software products for FREE without software&echo ============================================================================&echo.&echo #Supported products:&echo – Microsoft Office 2010 Standard Volume&echo – Microsoft Office 2010 Professional Plus Volume&echo.&echo.&(if exist “%ProgramFiles%\Microsoft Office\Office14\ospp.vbs” cd /d “%ProgramFiles%\Microsoft Office\Office14”)&(if exist “%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs” cd /d “%ProgramFiles(x86)%\Microsoft Office\Office14”)&echo.&echo ============================================================================&echo Activating your Office…&cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:8R6BM >nul&cscript //nologo ospp.vbs /unpkey:H3GVB >nul&cscript //nologo ospp.vbs /inpkey:V7QKV-4XVVR-XYV4D-F7DFM-8R6BM >nul&cscript //nologo ospp.vbs /inpkey:VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB >nul&set i=1 :server if %i%==1 set KMS_Sev=kms7.MSGuides.com if %i%==2 set KMS_Sev=kms8.MSGuides.com if %i%==3 set KMS_Sev=kms9.MSGuides.com if %i%==4 goto notsupported cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul&echo ============================================================================&echo.&echo. cscript //nologo ospp.vbs /act | find /i “successful” && (echo.&echo ============================================================================&echo.&echo #My official blog: MSGuides.com&echo.&echo #How it works: bit.ly/kms-server&echo.&echo #Please feel free to contact me at msguides.com@gmail.com if you have any questions or concerns.&echo.&echo #Please consider supporting this project: donate.msguides.com&echo #Your support is helping me keep my servers running everyday!&echo.&echo ============================================================================&choice /n /c YN /m “Would you like to visit my blog [Y,N]?” & if errorlevel 2 exit) || (echo The connection to my KMS server failed! Trying to connect to another one… & echo Please wait… & echo. & echo. & set /a i+=1 & goto server) explorer “http://MSGuides.com”&goto halt :notsupported echo.&echo ============================================================================&echo Sorry! Your version is not supported. :halt pause >nul
```

Hoặc:

Mở Command Prompt bằng quyền Run Administrator (Tức bấm tìm kiếm Command Prompt, bấm chuột phải chọn Run Administrator)

![1](https://user-images.githubusercontent.com/82578024/168939483-8b9dd175-677f-4493-9ed0-f6731ce10b40.gif)

Dán đoạn mã sau vào:

```php
if exist “%ProgramFiles%\Microsoft Office\Office14\ospp.vbs” cd /d “%ProgramFiles%\Microsoft Office\Office14”
if exist “%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs” cd /d “%ProgramFiles(x86)%\Microsoft Office\Office14”
cscript //Nologo OSPP.VBS /inpkey:VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB
cscript //Nologo ospp.vbs /sethst:kms.lotro.cc&cscript //nologo ospp.vbs /act&timeout 5&start winword&exit
@
```

# Đặc biệt Office 2010 đến 2019 bạn dùng **AIO Tools V3.1.3** kích hoạt rất OK #

**Kích hoạt Office bằng AIO Tools V3.1.3 [bấm vào đây để download](https://bit.ly/3O70Xnk)**

**Cũng có thể kích hoạt bằng MAS_1.5_AIO đoạn mã: [MAS_1.5_AIO_CRC32_21D20776.txt](https://github.com/BsNgChiThanh/Kich-hoat-Office/files/8711448/MAS_1.5_AIO_CRC32_21D20776.txt), trang chủ [bấm vào đây](https://massgrave.dev/). Hình ảnh khi chạy kích hoạt:**

 ![1](https://user-images.githubusercontent.com/82578024/168907463-2f726e32-dd9a-434f-a547-e99b84b80ae6.gif)

**Ngoài ra chúng ta có thể Download, cài đặt và kích hoạt Office từ [Office Tool Plus!](https://otp.landian.vip/en-us/)**

![1](https://user-images.githubusercontent.com/82578024/163676849-0c17b2f4-0316-4e02-a712-cb48914046e6.jpg)
Chọn Office sau đó intall licenses, bấm Yes
![2](https://user-images.githubusercontent.com/82578024/163676923-384d2e00-6f0d-4585-aeec-cdb22e5b08cd.jpg)
Quá trình Intall sẽ diễn ra, khi xong chúng ta bấm nút active!
![3](https://user-images.githubusercontent.com/82578024/163676979-a2c41195-a9ce-4ac9-a309-e38046730837.jpg)
![4](https://user-images.githubusercontent.com/82578024/163677053-a066a590-5f64-4890-a236-f0971909cfba.jpg)

**Hoàn thành Kích hoạt!**

**[Chia sẽ địa điểm trên Google map](https://goo.gl/maps/ZAzVMCgx4S4X4A55A)**
