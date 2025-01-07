rmdir /S /Q devdoc
cd C:\dev\livedocumenter\exporter\net471\
exporter.exe C:\dev\DBAddin.NET\docs\DBaddin.ldproj -to C:\dev\DBAddin.NET\docs\ -filters "public|protected|internalprotected|private"
SETLOCAL EnableDelayedExpansion
set fname=""
cd C:\dev\DBAddin.NET\docs 
for /d %%a in (*) do (
	set fname=%%~na
	set res=!fname:~0,9!
	if /i "!res!"=="LD Export" set "FolderPath=%%a"
)
echo !FolderPath!
rename "!FolderPath!" devdoc
del devdoc\8589934*
copy /Y index.htm.bak devdoc\index.htm
pause