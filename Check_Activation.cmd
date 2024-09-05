@echo off
cd /d %~dp0
color 1F
setLocal EnableDelayedExpansion
if exist "%Windir%\Sysnative\sppsvc.exe" set SysPath=%Windir%\Sysnative
if exist "%Windir%\System32\sppsvc.exe"  set SysPath=%Windir%\System32

echo Windows Status:
echo ===============
ver
cscript //nologo %SysPath%\slmgr.vbs /dli
cscript //nologo %SysPath%\slmgr.vbs /xpr

echo.
echo.
echo Office 2016 Status:
echo ===================
set office=
set installed=0
FOR /F "tokens=2*" %%a IN ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot" /v "Path" 2^>nul') do (SET office=%%b)
if exist "%office%\OSPP.VBS" (
	set installed=1
	cd /d "%office%"
	cscript //nologo ospp.vbs /dstatus
	cd /d %~dp0
)
set office=
FOR /F "tokens=2*" %%a IN ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Common\InstallRoot" /v "Path" 2^>nul') do (SET office=%%b)
if exist "%office%\OSPP.VBS" (
	set installed=1
	cd /d "%office%"
	cscript //nologo ospp.vbs /dstatus
	cd /d %~dp0
)
if %installed%==1 goto end2016
if exist "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" (
	set installed=1
	cd /d "C:\Program Files\Microsoft Office\Office16"
	cscript //nologo ospp.vbs /dstatus
	cd /d %~dp0
)
if exist "C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS" (
	set installed=1
	cd /d "C:\Program Files (x86)\Microsoft Office\Office16"
	cscript //nologo ospp.vbs /dstatus
	cd /d %~dp0
)
:end2016
if %installed%==0 echo Not installed

echo.
echo.
echo Office 2013 Status:
echo ===================
set office=
set installed=0
FOR /F "tokens=2*" %%a IN ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot" /v "Path" 2^>nul') do (SET office=%%b)
if exist "%office%\OSPP.VBS" (
	set installed=1
	cd /d "%office%"
	cscript //nologo ospp.vbs /dstatus
	cd /d %~dp0
)
set office=
FOR /F "tokens=2*" %%a IN ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\Common\InstallRoot" /v "Path" 2^>nul') do (SET office=%%b)
if exist "%office%\OSPP.VBS" (
	set installed=1
	cd /d "%office%"
	cscript //nologo ospp.vbs /dstatus
	cd /d %~dp0
)
if %installed%==0 echo Not installed

echo.
echo.
echo Office 2010 Status:
echo ===================
set office=
set installed=0
FOR /F "tokens=2*" %%a IN ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\Common\InstallRoot" /v "Path" 2^>nul') do (SET office=%%b)
if exist "%office%\OSPP.VBS" (
	set installed=1
	cd /d "%office%"
	cscript //nologo ospp.vbs /dstatus
	cd /d %~dp0
)
set office=
FOR /F "tokens=2*" %%a IN ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\14.0\Common\InstallRoot" /v "Path" 2^>nul') do (SET office=%%b)
if exist "%office%\OSPP.VBS" (
	set installed=1
	cd /d "%office%"
	cscript //nologo ospp.vbs /dstatus
	cd /d %~dp0
)
if %installed%==0 echo Not installed

echo.
pause
