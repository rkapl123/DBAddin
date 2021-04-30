rem copy Addin and settings...
@echo off
set /p answer="Enter Y to stop Excel (if running) and continue deployment of DB-Addin:"
if "%answer:~,1%" NEQ "Y" exit /b
taskkill /IM "Excel.exe" /F
if exist "C:\Program Files\Microsoft Office\root\" (
	echo 64bit office
	copy /Y DBaddin64.xll "%appdata%\Microsoft\AddIns\DBaddin.xll"
	copy /Y DBaddin.xll.config "%appdata%\Microsoft\AddIns\DBaddin.xll.config"
	copy /Y DBaddinCentral.config "%appdata%\Microsoft\AddIns\DBaddinCentral.config"
	copy /Y DBaddinUser.config "%appdata%\Microsoft\AddIns\DBaddinUser.config"
) else (
	echo 32bit office
	copy /Y DBaddin32.xll "%appdata%\Microsoft\AddIns\DBaddin.xll"
	copy /Y DBaddin.xll.config "%appdata%\Microsoft\AddIns\DBaddin.xll.config"
	copy /Y DBaddinCentral.config "%appdata%\Microsoft\AddIns\DBaddinCentral.config"
	copy /Y DBaddinUser.config "%appdata%\Microsoft\AddIns\DBaddinUser.config"
)
rem start Excel and install Addin there..
cscript //nologo switchToDBAddin.vbs
pause
