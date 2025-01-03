rmdir /S /Q devdoc
cd C:\dev\livedocumenter\exporter\net471\
exporter.exe C:\dev\DBAddin.NET\docs\DBaddin.ldproj -to C:\dev\DBAddin.NET\docs\ -filters "public|protected|internalprotected|private"
pause