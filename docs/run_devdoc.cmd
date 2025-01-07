: need to cd to livedocumenter for access to dlls
cd C:\dev\livedocumenter\exporter\net471\
exporter.exe C:\dev\DBAddin.NET\docs\DBaddin.ldproj -to C:\dev\DBAddin.NET\docs\ -filters "public|protected|internalprotected|private"

: now get produced folder to rename into devdoc
SETLOCAL EnableDelayedExpansion
cd C:\dev\DBAddin.NET\docs
rmdir /S /Q devdoc
for /d %%a in (*) do (
	set fname=%%~na
	set res=!fname:~0,9!
	if /i "!res!"=="LD Export" set "FolderPath=%%a"
)
echo !FolderPath!
rename "!FolderPath!" devdoc

: remove unnecessary files
del devdoc\8589934*
: enriched cover-page copied over standard index.htm
copy /Y index.htm.bak devdoc\index.htm
pause