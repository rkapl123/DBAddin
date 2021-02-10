On Error Resume Next
Set XLApp = GetObject(,"Excel.Application")
If err <> 0 Then
	Set XLApp = CreateObject("Excel.Application")
	WScript.Sleep 1000
End If
On Error goto 0
XLApp.Visible = true
For each ai in XLApp.AddIns
	If ai.name = "DBaddin.xll" then
		On Error Resume Next
		XLApp.AddIns("DBAddin.Functions").Installed = False ' deactivate old DB Addin ...
		If err <> 0 Then 
			WScript.Echo "no legacy DB-Addin installed."
		Else
			WScript.Echo "legacy DB-Addin uninstalled."
		End If
		On Error goto 0
		' install new add-in
		ai.Installed = True
	end if
next
' deactivate old DB Addin Startup (needs restart of excel)
Set WshShell = CreateObject("WScript.Shell")
WshShell.RegWrite "HKCU\Software\Microsoft\Office\Excel\Addins\DBAddin.Connection\LoadBehavior",2,"REG_DWORD"
Wscript.Echo ("Please restart Excel to make Installation effective ...")