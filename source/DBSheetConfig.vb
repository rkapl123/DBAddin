Imports Microsoft.Office.Interop
'''<summary>Helper class for DBSheetHandler and DBSheetConnection for easier manipulation of DBSheet definition / Connection configuration data</summary> 
Public Class DBSheetConfig
    ''' <summary>current Sheet configuration as XML string</summary>
    Public curConfig As String
    ''
    ' current SheetParams (contains reference to configuration file) as XML string
    Public SheetParams As String
    ''
    ' current connections file content as XML string (loaded with getDBConnectionsFile)
    Public DBConnFile As String

    ''
    ' init procedure for DBSheetHandler Config
    ' @param Sh the sheet which needs to be initialised as a DBSheet
    ' @remarks gets DBSheet definition from current sheet's top/leftmost cell's comment
    Public Sub initDBSheetConfig(Sh As Excel.Worksheet)
        Dim thePath As String

        On Error GoTo err1
        SheetParams = Sh.Range("A1").Comment.Text
        thePath = getEntry("DBSheetConfigPath", SheetParams)
        curConfig = readDBSheetConfig(thePath)
        Exit Sub
err1:
        LogError("Error: " & Err.Description & ", line " & Erl() & " in DBSheetConfig.initDBSheetConfig")
    End Sub

    ''
    ' reads DBSheet definition from thePath
    ' @param thePath
    ' @return the config
    Public Function readDBSheetConfig(thePath As String) As String
        On Error GoTo err1
        If Len(thePath) = 0 Then
            LogWarn("Error: no DBSheet config file given (need to set DBSheetConfigPath) !")
            Exit Function
        End If
        If Left$(thePath, 2) <> "\\" And Mid$(thePath, 2, 2) <> ":\" Then
            thePath = DBSheetDefinitionsFolder & "\" & thePath
        End If

        If Len(Dir(thePath)) > 0 Then
            On Error Resume Next
            Dim fileReader As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(thePath)
            readDBSheetConfig = fileReader.ReadToEnd()
            If Err().Number <> 0 Then LogWarn("Error: " & Err.Description & " while reading DBSheet config file in initDBSheetConfig")
            On Error GoTo err1
        Else
            LogWarn("Error: couldn't find DBSheet config file '" & thePath & "' !")
        End If
        Exit Function
err1:
        readDBSheetConfig = vbNullString
        LogError("Error: " & Err.Description & ", line " & Erl() & " in DBSheetConfig.readDBSheetConfig")
    End Function

    ''
    ' creates markup with setting value content in entryMarkup
    ' @param entryMarkup
    ' @param content
    ' @return the markup
    Public Function setEntry(ByVal entryMarkup As String, ByVal content As String) As String
        setEntry = "<" & entryMarkup & ">" & content & "</" & entryMarkup & ">"
    End Function

    ''
    ' fetches value in entryMarkup within paramText, search starts at position startSearch
    ' @param entryMarkup
    ' @param XMLString
    ' @param startSearch
    ' @return the value
    Public Function getEntry(ByVal entryMarkup As String, Optional ByVal XMLString As String = vbNullString, Optional startSearch As Integer = 1) As String
        Dim markStart As String, markEnd As String
        Dim fetchBeg, fetchEnd As Integer

        On Error GoTo getEntry_Err
        If Len(XMLString) = 0 Then
            If Len(DBConnFile) > 0 Then
                XMLString = curConfig
            Else
                ' if we're in a DBSheet combine sheet params with central DBSheet config
                XMLString = SheetParams & curConfig
            End If
        End If
        If Len(XMLString) = 0 Then
            getEntry = vbNullString
            Exit Function
        End If

        markStart = "<" & entryMarkup & ">"
        markEnd = "</" & entryMarkup & ">"

        fetchEnd = startSearch
        fetchBeg = InStr(fetchEnd, XMLString, markStart)
        If fetchBeg = 0 Then
            getEntry = vbNullString
            Exit Function
        End If
        fetchEnd = InStr(fetchBeg, XMLString, markEnd)
        startSearch = fetchEnd
        getEntry = Mid$(XMLString, fetchBeg + Len(markStart), fetchEnd - (fetchBeg + Len(markStart)))
        Exit Function

getEntry_Err:
        LogError("Error: " & Err.Description & ", line " & Erl() & " in DBSheetConfig.getEntry")
    End Function

    ''
    ' fetches entryMarkup parts contained within lists demarked by listMarkup within parentMarkup
    ' @param parentMarkup
    ' @param listMarkup
    ' @param entryMarkup
    ' @param XMLString
    ' @return list containing parts
    ' @remark
    ' returns list containing parts, if entryMarkup = vbNullString then list contains parts demarked by listMarkup
    Public Function getEntryList(ByVal parentMarkup As String, ByVal listMarkup As String, ByVal entryMarkup As String, Optional XMLString As String = vbNullString) As Object
        Dim list() As String = Nothing
        Dim i As Long, posEnd As Long
        Dim isFilled As Boolean
        Dim parentEntry As String, ListItem As String, part As String

        On Error GoTo getEntryList_Err
        If Len(XMLString) = 0 Then
            If Len(DBConnFile) > 0 Then
                XMLString = curConfig
            Else
                ' if we're in a DBSheet combine sheet params with central DBSheet config
                XMLString = SheetParams & curConfig
            End If
        End If
        If Len(XMLString) = 0 Then
            getEntryList = Nothing
            Exit Function
        End If

        i = 0 : posEnd = 1 : isFilled = False
        parentEntry = getEntry(parentMarkup, XMLString)
        Do
            ListItem = getEntry(listMarkup, XMLString, posEnd)
            If Len(entryMarkup) = 0 Then
                part = ListItem
            Else
                part = getEntry(entryMarkup, ListItem)
            End If
            If Len(part) > 0 Then
                isFilled = True
                ReDim Preserve list(i)
                list(i) = part
                i += 1
            End If
        Loop Until ListItem = ""
        If isFilled Then
            getEntryList = list
        Else
            getEntryList = Nothing
        End If
        Exit Function

getEntryList_Err:
        LogError("Error: " & Err.Description & ", line " & Erl() & " in DBSheetConfig.getEntryList")
    End Function

    ''
    ' sets markup denoted by entryMarkup to content (xml parameters contained in sheet ws)
    ' @param entryMarkup
    ' @param content
    Public Sub changeMarkup(ByVal entryMarkup As String, ByVal content As String)
        Dim tryConfigChange As String, oldConfig As String, newEntry As String, newConfig As String

        On Error GoTo changeMarkup_Err
        newEntry = setEntry(entryMarkup, content)
        oldConfig = SheetParams
        tryConfigChange = Change(oldConfig, "<" & entryMarkup & ">", content, "</" & entryMarkup & ">")
        If Len(tryConfigChange) = 0 Then
            newConfig = Replace(oldConfig, "</DBsheetParams>", newEntry & vbLf & "</DBsheetParams>")
        Else
            newConfig = tryConfigChange
        End If
        SheetParams = newConfig
        Exit Sub

changeMarkup_Err:

        LogError("Error: " & Err.Description & ", line " & Erl() & " in DBSheetConfig.changeMarkup")
    End Sub

End Class

