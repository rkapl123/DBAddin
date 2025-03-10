Set /P answr=deploy (r)elease (empty for debug)? 
@echo off
set source=bin\Debug
If "%answr%"=="r" (
	set source=bin\Release
)
if exist "C:\Program Files\Microsoft Office\root\" (
	echo 64bit office
	copy /Y %source%\DBaddin-AddIn64-packed.xll "%appdata%\Microsoft\AddIns\DBaddin.xll"
	copy /Y %source%\DBaddin.pdb "%appdata%\Microsoft\AddIns"
	rem copy /Y %source%\DBaddin.dll.config "%appdata%\Microsoft\AddIns\DBaddin.xll.config"
	rem copy /Y DBaddinCentral.config "%appdata%\Microsoft\AddIns\DBaddinCentral.config"
	rem copy /Y DBAddinUser.config "%appdata%\Microsoft\AddIns\DBaddinUser.config"
) else (
	echo 32bit office
	copy /Y %source%\DBaddin-AddIn-packed.xll "%appdata%\Microsoft\AddIns\DBaddin.xll"
	copy /Y %source%\DBaddin.pdb "%appdata%\Microsoft\AddIns"
	rem copy /Y %source%\DBaddin.dll.config "%appdata%\Microsoft\AddIns\DBaddin.xll.config"
	rem copy /Y DBaddinCentral.config "%appdata%\Microsoft\AddIns\DBaddinCentral.config"
	rem copy /Y DBAddinUser.config "%appdata%\Microsoft\AddIns\DBaddinUser.config"
)
if "%source%"=="bin\Release" (
	echo copying Distribution
	copy /Y %source%\DBaddin-AddIn64-packed.xll "..\Distribution\DBaddin64.xll"
	copy /Y %source%\DBaddin-AddIn-packed.xll "..\Distribution\DBaddin32.xll"
	rem copy /Y %source%\DBaddin.dll.config "..\Distribution\DBaddin.xll.config"
	rem copy /Y DBaddinCentral.config "..\Distribution\DBaddinCentral.config"
	rem copy /Y DBAddinUser.config "..\Distribution\DBaddinUser.config"
)
cscript replaceXMLdocFaults.vbs
pause
