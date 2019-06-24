Set /P answr=deploy (r)elease (empty for debug)? 
set source=bin\Debug
If "%answr%"=="r" (
set source=bin\Release
)
copy /Y %source%\DBaddin-AddIn-packed.xll "%appdata%\Microsoft\AddIns"
copy /Y %source%\DBaddin.pdb "%appdata%\Microsoft\AddIns"
copy /Y %source%\DBaddin.dll.config "%appdata%\Microsoft\AddIns\DBaddin-AddIn-packed.xll.config"
rem regsvr32.exe "%appdata%\Microsoft\DBaddin-AddIn-packed.xll"
pause

