rem copy Addin and settings...
@echo off
if exist "C:\Program Files\Microsoft Office\root\" (
	echo 64bit office
	copy /Y DBaddin64.xll "%appdata%\Microsoft\AddIns\DBaddin.xll"
	copy /Y DBaddin.xll.config "%appdata%\Microsoft\AddIns\DBaddin.xll.config"
	copy /Y DBaddinCentral.config "%appdata%\Microsoft\AddIns\DBaddinCentral.config"
) else (
	echo 32bit office
	copy /Y DBaddin32.xll "%appdata%\Microsoft\AddIns\DBaddin.xll"
	copy /Y DBaddin.xll.config "%appdata%\Microsoft\AddIns\DBaddin.xll.config"
	copy /Y DBaddinCentral.config "%appdata%\Microsoft\AddIns\DBaddinCentral.config"
)
rem start Excel and install Addin there..
cscript //nologo switchToDBAddin.vbs
pause
