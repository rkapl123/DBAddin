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
			if not obj_datadict.exists(setting) then
				if setting = "ConfigSelect + Globals.fetchSetting(ConfigSelectPreference" then
					obj_datadict.add "ConfigSelect + Globals.fetchSetting(ConfigSelectPreference)", "ConfigSelect + Globals.fetchSetting(ConfigSelectPreference)"
					setting = Replace(setting, "ConfigSelect + Globals.fetchSetting(", "")
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
			'wscript.echo "before:" & printDict(arrKeys)
            If(arrKeys(i) >= arrKeys(j)) Then
                temp = arrKeys(i)
                arrKeys(i) = arrKeys(j)
                arrKeys(j) = temp
				'wscript.echo "after:" & printDict(arrKeys)
            End If
        Next
    Next
    SortDictToArray = arrKeys
End Function

Function printDict(ByVal arrKeys)
	For i=0 to Ubound(arrKeys)-1
		printDict = printDict & arrKeys(i) & chr(13)
	Next
End Function
