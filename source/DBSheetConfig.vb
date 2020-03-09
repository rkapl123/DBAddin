Imports System.IO
Imports System.Windows.Forms
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop


'''<summary>Helper class for DBSheetHandler and DBSheetConnection for easier manipulation of DBSheet definition / Connection configuration data</summary> 
Public Module DBSheetConfig
    ''' <summary>current Sheet configuration as XML string</summary>
    Public curConfig As String

    Public Sub createDBSheet()
        'MsgBox("not yet implemented..")
        'Exit Sub
        Dim openFileDialog1 = New OpenFileDialog With {
            .InitialDirectory = fetchSetting("DBSheetDefinitions", ""),
            .Filter = "XML files (*.xml)|*.xml",
            .RestoreDirectory = True
        }
        Dim result As DialogResult = openFileDialog1.ShowDialog()
        If result = Windows.Forms.DialogResult.OK Then
            ' Get the DBSheet Definition file name and read into curConfig
            Dim dsdPath As String = openFileDialog1.FileName
            curConfig = File.ReadAllText(dsdPath)
            ' get query
            Dim queryStr As String = getEntry("query")
            ' get lookup fields
            If Not IsNothing(getEntryList("columns", "field", "lookup")) Then
                Dim columnslist() As String       ' the list of the column infos (including lookups)
                Dim LookupDef
                columnslist = getEntryList("columns", "field", vbNullString)
                ' add lookup Queries in separate sheet
                For Each LookupDef In columnslist
                    MsgBox(LookupDef)
                Next
            End If
            ' add DBMapper with query and vlookup function fields for resolution of lookups

        End If
    End Sub

    ''' <summary>fetches value in entryMarkup within XMLString (if not given takes curConfig), search starts optionally at position startSearch (default 1)</summary>
    ''' <param name="entryMarkup"></param>
    ''' <param name="XMLString"></param>
    ''' <param name="startSearch"></param>
    ''' <returns>the fetched value</returns>
    Public Function getEntry(ByVal entryMarkup As String, Optional ByVal XMLString As String = vbNullString, Optional ByRef startSearch As Integer = 1) As String
        Dim markStart As String, markEnd As String
        Dim fetchBeg, fetchEnd As Integer

        On Error GoTo getEntry_Err
        If IsNothing(XMLString) Then XMLString = curConfig
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

    ''' <summary>fetches entryMarkup parts contained within lists demarked by listMarkup within parentMarkup inside XMLString (if not given takes curConfig)</summary>
    ''' <param name="parentMarkup"></param>
    ''' <param name="listMarkup"></param>
    ''' <param name="entryMarkup"></param>
    ''' <param name="XMLString"></param>
    ''' <returns>list containing parts, if entryMarkup = vbNullString then list contains parts demarked by listMarkup</returns>
    Public Function getEntryList(ByVal parentMarkup As String, ByVal listMarkup As String, ByVal entryMarkup As String, Optional XMLString As String = vbNullString) As Object
        Dim list() As String = Nothing
        Dim i As Long, posEnd As Long
        Dim isFilled As Boolean
        Dim parentEntry As String, ListItem As String, part As String

        On Error GoTo getEntryList_Err
        If IsNothing(XMLString) Then XMLString = curConfig
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


End Module

