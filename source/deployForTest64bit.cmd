Set /P answr=deploy (r)elease (empty for debug)? 
set source=bin\Debug
If "%answr%"=="r" (
set source=bin\Release
)
copy /Y %source%\DBaddin-AddIn64-packed.xll "%appdata%\Microsoft\AddIns"
copy /Y %source%\DBaddin.pdb "%appdata%\Microsoft\AddIns"
copy /Y %source%\DBaddin.dll.config "%appdata%\Microsoft\AddIns\DBaddin-AddIn64-packed.xll.config"
REM Create .tlb file
REM Setting up environment vairables
call "C:\Program Files (x86)\Microsoft Visual Studio 14.0\VC\vcvarsall.bat" x86
REM Temporarily copy ExcelDna.Integration.dll into output
REM Note: Might need to change depending on where packages directory is
copy "packages\ExcelDna.Integration.0.34.6\lib\ExcelDna.Integration.dll" "%source%"
tlbexp.exe "%source%\DBaddin.dll" /out:"%source%\DBaddin.tlb"
"packages\ExcelDna.AddIn.0.34.6\tools\ExcelDnaPack.exe" "%source%\DBaddin-AddIn64.dna" /Y  /O "%source%\DBaddin-AddIn64-packed.xll"
regsvr32.exe /s "%appdata%\Microsoft\AddIns\DBaddin-AddIn64-packed.xll"
pause

