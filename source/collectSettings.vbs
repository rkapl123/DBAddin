' develop support script to collect all settings from fetchSetting calls within DBAddni codebase
' results are put into resource file Settings.txt which is used for the insert dropdown in EditDBModifDef (user/global settings).
Set fso = CreateObject("Scripting.FileSystemObject")
Dim obj_datadict
Set obj_datadict = CreateObject("Scripting.Dictionary")
For Each inFile In fso.GetFolder(".").Files
  If UCase(fso.GetExtensionName(inFile.Name)) = "VB" Then
	Set ifile = fso.OpenTextFile(inFile.path, 1)
	theText = ifile.readAll()
	Set regEx = New RegExp
	regEx.Pattern = "fetchSetting\(""(.*?),"
	regEx.Global = True
	regEx.IgnoreCase = True
	Set myMatches = regEx.Execute(theText)
	For Each myMatch in myMatches 
		For Each mySubMatch in myMatch.SubMatches
			setting = Replace(mySubMatch, """", "")
			setting = Replace(setting, "Globals.", "")
			setting = Replace(setting, "()", "")
			setting = Replace(setting, ".ToString", "")
			setting = Replace(setting, "DBenv", "env")
			setting = Replace(setting, "myDBConnHelper.", "")
			setting = Replace(setting, "ConfigName + i", "ConfigName + env")
			if not obj_datadict.exists(setting) then
				if setting = "ConfigSelect + fetchSetting(ConfigSelectPreference" then
					obj_datadict.add "ConfigSelect + ConfigSelectPreference", "ConfigSelect + ConfigSelectPreference"
					setting = Replace(setting, "ConfigSelect + fetchSetting(", "")
				end if
				obj_datadict.add setting, setting
			end if
		next
	Next
	ifile.Close
	Set ifile = Nothing
  End if
Next
sortedArray = SortDictToArray(obj_datadict)
Set ofile = fso.CreateTextFile("Settings.txt", True)
For i=0 to Ubound(sortedArray)-1
	ofile.writeline(sortedArray(i)) 
Next
ofile.Close
Set ofile = Nothing
Set fso = Nothing

Function SortDictToArray(ByVal dict)
   arrKeys = dict.keys
   For i=0 To UBound(arrKeys)-1
        For j=i+1 To UBound(arrKeys)
            If(arrKeys(i) >= arrKeys(j)) Then
                temp = arrKeys(i)
                arrKeys(i) = arrKeys(j)
                arrKeys(j) = temp
            End If
        Next
    Next
    SortDictToArray = arrKeys
End Function
