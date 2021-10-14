Set WshShell = CreateObject("WScript.Shell")
' disable old addin
WshShell.RegWrite "HKCU\Software\Microsoft\Office\Excel\Addins\DBAddin.Connection\LoadBehavior",2,"REG_DWORD"

AddinName = "DBAddin"

dim OfficeVersion
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists("C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE") Or fso.FileExists("C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE") Then OfficeVersion = "16.0"
If fso.FileExists("C:\Program Files (x86)\Microsoft Office\root\Office15\EXCEL.EXE") Or fso.FileExists("C:\Program Files\Microsoft Office\root\Office15\EXCEL.EXE") Then OfficeVersion = "15.0"
If fso.FileExists("C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE") Or fso.FileExists("C:\Program Files\Microsoft Office\Office14\EXCEL.EXE") Then OfficeVersion = "14.0"
If fso.FileExists("C:\Program Files (x86)\Microsoft Office\Office12\EXCEL.EXE") Or fso.FileExists("C:\Program Files\Microsoft Office\Office12\EXCEL.EXE") Then OfficeVersion = "12.0"

addToReg = True
' write new Addin into OPEN key
i = 0
Do
	on error resume next
	bKey = WshShell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Office\" & OfficeVersion & "\Excel\Options\OPEN" & iif(i=0,"",i))
	if Lcase(bKey) = """" & Lcase(AddinName) & """" then addToReg = False
	if err <> 0 then
		if addToReg then
			WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\" & OfficeVersion &"\Excel\Options\OPEN" & iif(i=0,"",i),"""" & AddinName & """","REG_SZ"
		end if
		exitMe = True
	else
		i = i + 1
	end if
Loop Until exitMe

Function IIf(expr, truepart, falsepart)
	If expr Then 
		IIf = truepart
	Else
		IIf = falsepart
	End if
End Function
