Set /P answr=deploy (r)elease (empty for debug)? 
set source=bin\Debug
If "%answr%"=="r" (
set source=bin\Release
)
if exist "C:\Program Files\Microsoft Office\root\" (
	echo 64bit office
	copy /Y %source%\DBaddin-AddIn64-packed.xll "%appdata%\Microsoft\AddIns\DBaddin.xll"
	copy /Y %source%\DBaddin.pdb "%appdata%\Microsoft\AddIns"
	copy /Y %source%\DBaddin.dll.config "%appdata%\Microsoft\AddIns\DBaddin.xll.config"
) else (
	echo 32bit office
	copy /Y %source%\DBaddin-AddIn-packed.xll "%appdata%\Microsoft\AddIns\DBaddin.xll"
	copy /Y %source%\DBaddin.pdb "%appdata%\Microsoft\AddIns"
	copy /Y %source%\DBaddin.dll.config "%appdata%\Microsoft\AddIns\DBaddin.xll.config"
)
pause
