Imports Microsoft.Office.Interop
Imports ExcelDna.Integration

''' <summary>Class handling the manipulation Of DBSheets (reading data, storing To DB). Also Handles some worksheet events (context menu, deactivating)</summary>
'Public Class DBSheetHandler
'    Private config As DBSheetConfig

'    ''
'    ' for filling the header collection and header row in DBsheet
'    Private rst As ADODB.Recordset
'    ''
'    ' for checking field properties in the underlying table
'    Private fieldPropCheck As ADODB.Recordset
'    ''
'    ' the table name for the current DBSheet
'    Private tableName As String
'    ''
'    ' the query for the current DBSheet
'    Private queryStr As String
'    ''
'    ' how many primary key columns are in the DBSHeet
'    Private primKeyColumns As Long
'    ''
'    ' column where calculation columns begin
'    Private calcColumnsStart As Long
'    ''
'    ' current amount of loaded rows+1, needed to determine whether we enter new data (insert flag)...
'    Private maxRows As Long
'    ''
'    ' where we keep the formulas from the calculated columns before overwriting data in sheet
'    Private formulaCells() As String
'    ''
'    ' used for setting formula1 for restricting bool values (xlValidateList is language dependent !!)
'    Private TrueFalseSelection As String
'    ''
'    ' (orange) header color for test environment
'    Private testHeaderColor As Integer
'    ''
'    ' header color for prod environment (vbBlack)
'    Private prodHeaderColor As Integer
'    ''
'    ' tab color if no errs happened (last color before DBSheet was saved)
'    Private noErrColor As Integer
'    ''
'    ' tab color if internal Errs
'    Private internalErrColor As Integer
'    ''
'    ' tab color if data Errors
'    Private dataErrColor As Integer
'    ''
'    ' color for required fields in the header
'    Private reqFieldsColor As Integer
'    ''
'    ' color for primary columns
'    Private primColFieldsColor As Integer
'    ''
'    ' color for calc columns
'    Private calcColFieldsColor As Integer
'    ''
'    ' color for displaying conflicting fields
'    Private conflictColor As Integer
'    ''
'    ' if user changes sheets, this disables ability to cancel refresh action
'    Private enforceRefresh As Boolean

'    Private Sub Class_Terminate()
'        config = Nothing
'    End Sub


'    ''
'    ' initializes the DBSheetHandler with settings from the DBSheetConfig and registry defaults (fetchSetting)
'    ' @param Sh
'    ' @param mandatoryDBsheetRefresh
'    ' @remarks also removes last DBSheet's helper sheets if a DBsheet was removed (lastWsName is still existing) ...
'    Public Sub Initialize(Sh As Excel.Worksheet)
'        On Error GoTo err1
'        config = New DBSheetConfig
'        config.initDBSheetConfig(Sh)
'        Exit Sub
'err1:
'        LogError("Error: " & Err.Description & ", line " & Erl() & " in DBSheetHandler.Initialize")
'    End Sub

'    ''
'    ' refreshes DBSHeet definition, which is needed in all cases when switching to a DBsheet, even
'    ' when refresh of data is not done (enforceRefresh -> False)...
'    ' @return refresh was succesful
'    ' @remarks
'    Public Function readDBSheetDefs() As Boolean
'        Dim theField As ADODB.Field
'        Dim addtlErrInfo As String

'        On Error GoTo readDBSheetDefs_Err
'        readDBSheetDefs = False

'        tableName = config.getEntry("table")
'        If Len(tableName) = 0 Then
'            LogWarn("table name not defined in DBSheet configuration of " & ws.name)
'            Exit Function
'        End If
'        ' force refresh of connection definition before reading basic definitions !!
'        If Not theDBSheetConnection.config.initConnectionConfig(config.getEntry("connID"), False, True) Then
'            LogError("readDBsheet: couldn't retrieve connDef for connection id: '" & config.getEntry("connID") & "'")
'            Exit Function
'        End If
'        primKeyColumns = CLng(config.getEntry("primcols"))
'        If primKeyColumns = 0 Then
'            LogWarn("readHeaders: primcols entry is 0 (not allowed)!!")
'            Exit Function
'        End If
'        calcColumnsStart = CLng(config.getEntry("calcedcols"))

'        If Not theDBSheetConnection.openConnection(resetDB:=True) Then
'            LogError("readDBSheetDefs: couldn't open connection for connid: '" & config.getEntry("connID") & "'")
'            Exit Function
'        End If

'        queryStr = config.getEntry("query")
'        If Not replaceQueryParameters(queryStr) Then
'            LogWarn("couldn't resolve parameters given in DBSheet's restriction, please check Parameter settings for this sheet !")
'            Exit Function
'        End If

'        rst = New ADODB.Recordset
'        fieldPropCheck = New ADODB.Recordset

'        addtlErrInfo = ", opening fieldPropCheck on table: " & tableName & " and recordset on query: " & queryStr
'        fieldPropCheck.Open(tableName, dbcnn, adOpenForwardOnly, adLockReadOnly)
'        rst.Open(queryStr, dbcnn, adOpenForwardOnly, adLockReadOnly, -1)

'        ' add fields and their column order to the fields collection
'        Dim j As Long
'        j = 1
'        For Each theField In rst.Fields
'            On Error Resume Next
'            Dim dummy As String
'            dummy = fieldPropCheck.Fields(theField.Name).Name
'            If Err().Number <> 0 Then
'                LogWarn("Error: " & Err.Description & ", probably field '" & theField.Name & "' was not found in data table (" & tableName & ") ! " & vbLf _
'                   & "Maybe you didn't define the alias in the lookup query correctly (has to be exactly the name of the field in the table) ?")
'                Err.Clear()
'                GoTo final
'            End If
'            On Error GoTo readDBSheetDefs_Err
'            theFields.Add theField.Name, CStr(j)
'         j = j + 1
'        Next

'        ' read lookup columns and assign to corresponding auxsheets
'        If Not IsNothing(config.getEntryList("columns", "field", "lookup")) Then
'            Dim columnslist() As String       ' the list of the column infos (including lookups)
'            Dim LookupDef
'            Dim lookupwsname As String, lookupName As String

'            columnslist = config.getEntryList("columns", "field", vbNullString)
'            ' add/clear lookup and lookup value sheets
'            lookupwsname = "L" & IIf(Len(ws.CodeName) > 30, Left$(ws.CodeName, 30), ws.CodeName)
'            theFields.lookUpsFilled = False
'            For Each LookupDef In columnslist
'                lookupName = correctNonNull(config.getEntry("name", LookupDef))
'                ' now add the lookup to field collection
'                If checkLookupRange(lookupwsname, lookupName) Then
'                    theFields.Add lookupName, ws.Parent.Worksheets(lookupwsname).Range(lookupName), LenB(config.getEntry("ftable", LookupDef)) > 0
'                 theFields.lookUpsFilled = True
'                End If
'            Next
'        End If

'        ' as we may only rely on the data in the sheet, maxRows is not taken from the Database
'        maxRows = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
'        readDBSheetDefs = True
'        Exit Function

'readDBSheetDefs_Err:
'        LogError("Error: " & Err.Description & ", line " & Erl() & " in DBSheetHandler.readDBSheetDefs" & addtlErrInfo)
'final:
'        cleanUpRecordsets()
'    End Function

'    ''
'    ' @param lookupwsname
'    ' @param lookupName
'    ' @return checkLookupRange was successfull
'    ' @remarks
'    Private Function checkLookupRange(lookupwsname As String, lookupName As String) As Boolean
'        Dim checkWhetherLookupEmpty As String

'        checkLookupRange = False
'        On Error Resume Next
'        checkWhetherLookupEmpty = CStr(ws.Parent.Worksheets(lookupwsname).Range(lookupName).Cells(1, 1))
'        If Len(checkWhetherLookupEmpty) > 0 Then checkLookupRange = True
'        Err.Clear()
'    End Function


'    Private Sub cleanUpRecordsets()
'100     Set rst = Nothing
'102     Set fieldPropCheck = Nothing
'End Sub

'    ''
'    ' read lookup information from database
'    ' @return function was successfull
'    ' @remarks
'    Private Function readLookups() As Boolean
'        Dim forTableRst As ADODB.Recordset
'        Dim columnslist() As String       ' the list of the column infos (including lookups)
'        Dim adhocVals() As String       ' ad hoc values as lookups (spearated by ||)
'        Dim queryStr As String, lookupwsname As String, vlookupwsname As String, lookupName As String, tmpname As String, addtlErrInfo As String
'        Dim retRecords As Long
'        Dim LookupDef

'        On Error GoTo err1
'100     readLookups = False
'102     If IsEmpty(config.getEntryList("columns", "field", "lookup")) Then
'104         readLookups = True
'            Exit Function
'        End If
'106     columnslist = config.getEntryList("columns", "field", vbNullString)

'        ' add/clear lookup and lookup value sheets
'108     lookupwsname = makeAuxSheet("L", True)
'110     vlookupwsname = makeAuxSheet("V", True)

'112     If Not theDBSheetConnection.openConnection Then Exit Function

'        ' first add the lookup queries to sheet via querytables
'114     Dim tableCount As Long: Dim place As Long
'116     tableCount = 1: place = 1
'118     For Each LookupDef In columnslist
'120         lookupName = correctNonNull(config.getEntry("name", LookupDef))
'122         place = theFields.Column(lookupName)
'124         queryStr = config.getEntry("lookup", LookupDef)
'126         Dim IsForeignLookup: IsForeignLookup = True
'128         If LenB(queryStr) > 0 Then
'130             If InStr(1, UCase$(queryStr), "SELECT") > 0 Then
'132                 Dim theForTable As String: theForTable = quotedReplace(queryStr, "T" & tableCount)
'134                 addtlErrInfo = "opening lookup info for " & theForTable
'136                 Set forTableRst = New ADODB.Recordset
'138                 forTableRst.CursorLocation = adUseClient
'140                 forTableRst.Open theForTable, dbcnn, adOpenForwardOnly, adLockReadOnly
'142                 With ws.Parent.Worksheets(lookupwsname).QueryTables.Add(Connection:=forTableRst, Destination:=ws.Parent.Worksheets(lookupwsname).Cells(1, place))
'144                     .FieldNames = False
'146                     .PreserveFormatting = True
'148                     .AdjustColumnWidth = False
'150                     tmpname = .name
'152                     .Refresh
'154                     retRecords = forTableRst.RecordCount
'156                     .Delete
'                    End With
'        ' sometimes excel doesn't delete the querytable given name
'        On Error Resume Next
'158                 ws.Parent.Worksheets(lookupwsname).Names(tmpname).Delete
'160                 ws.Parent.Names(tmpname).Delete
'                    On Error GoTo err1
'162                 forTableRst.Close

'                    ' copy lookup values to value sheet...
'164                 ws.Parent.Worksheets(lookupwsname).Columns(place + 1).Copy
'166                 ws.Parent.Worksheets(vlookupwsname).Visible = xlSheetVisible
'168                 ws.Parent.Worksheets(vlookupwsname).Select
'170                 ws.Parent.Worksheets(vlookupwsname).Cells(1, place).Select
'172                 ws.Parent.Worksheets(vlookupwsname).Paste
'174                 ws.Parent.Worksheets(lookupwsname).Columns(place + 1).Clear

'                    ' fill adhoc values...
'                Else
'        ' think about a nicer representation (xml?)
'176                 parse queryStr, adhocVals, "||"
'                    Dim i As Long
'178                 For i = 0 To UBound(adhocVals)
'180                     ws.Parent.Worksheets(lookupwsname).Cells(i + 1, place).Value = adhocVals(i)
'182                     ws.Parent.Worksheets(vlookupwsname).Cells(i + 1, place).Value = adhocVals(i)
'                    Next
'184                 retRecords = i
'                End If
'186             With ws.Parent.Worksheets(lookupwsname)
'188                 .Range(.Cells(1, place), .Cells(retRecords + IIf(retRecords < 1, 1, 0), place)).name = lookupName
'                    ' now add the lookup to field collection if not done so already in readDBSheetDefs
'190                 If Not theFields.lookUpsFilled Then _
'                       theFields.Add lookupName, .Range(lookupName), LenB(config.getEntry("ftable", LookupDef)) > 0
'192                 .Range(.Cells(1, place + 1), .Cells(retRecords + IIf(retRecords < 1, 1, 0), place + 1)).ClearContents
'                End With
'194             Dim r As Long: Dim V As Variant
'196             If retRecords > 0 Then
'198                 With ws.Parent.Worksheets(lookupwsname).Range(ws.Parent.Worksheets(lookupwsname).Cells(1, place), ws.Parent.Worksheets(lookupwsname).Cells(retRecords, place))
'200                     For r = retRecords To 1 Step -1
'202                         V = Replace$(Replace$(Replace$(.Cells(r, 1).Value, "*", "~*"), "?", "~?"), "~", "~~")
'204                         If V = vbNullString Then
'206                             LogError "Error: empty lookup: " & lookupwsname & "!" & .Address & "!"
'                            Else
'        On Error Resume Next
'208                             If theApp.WorksheetFunction.CountIf(.Columns(1), V) > 1 Then
'210                                 If Err = 0 Then LogError "Error: duplicate lookup: " & V & " !"
'                                End If
'        On Error GoTo err1
'        End If
'212                     Next r
'                    End With
'214                 With ws.Parent.Worksheets(vlookupwsname).Range(ws.Parent.Worksheets(vlookupwsname).Cells(1, place), ws.Parent.Worksheets(vlookupwsname).Cells(retRecords, place))
'216                     For r = retRecords To 1 Step -1
'218                         V = Replace$(Replace$(Replace$(.Cells(r, 1).Value, "*", "~*"), "?", "~?"), "~", "~~")
'220                         If V = vbNullString Then
'222                             LogError "Error: empty lookup value: " & vlookupwsname & "!" & .Address & "!"
'                            Else
'        On Error Resume Next
'224                             If theApp.WorksheetFunction.CountIf(.Columns(1), V) > 1 Then
'226                                 If Err = 0 Then LogError "Error: duplicate lookup value: " & V & " !"
'                                End If
'        On Error GoTo err1
'        End If
'228                     Next r
'                    End With
'        End If
'        End If
'        Next
'230     readLookups = True
'232     GoTo final
'err1:
'234     If VBDEBUG Then Debug.Print Err.Description: Stop : Resume Next
'236     readLookups = False
'238     LogError "Error: " & Err.Description & ", line " & Erl & " in DBSheetHandler.readLookups " & addtlErrInfo
'final:
'240     ws.Parent.Worksheets(lookupwsname).Visible = xlSheetVeryHidden
'242     ws.Parent.Worksheets(vlookupwsname).Visible = xlSheetVeryHidden
'244     ws.Select
'End Function

'    ''
'    ' reads the data as defined in the db sheet definition (comment in range A1 of DB sheet)<br>
'    ' @return function was successfull
'    ' @remarks
'    ' with locking header row and primary columns, also copying stored formulas for calculated cells and<br>
'    ' entering validation lists for foreign key lookup columns<br>
'    ' produces copy of DB sheet for checking stale data (multiuser editing !) into hidden sheet D<wsname><br>
'    Private Function readData() As Boolean
'        Dim rst As Recordset
'        Dim queryStr As String, tmpname As String, olddatawsname As String
'        Dim copyFormat() As String
'        Dim curWin As Excel.Window, otherWin As Excel.Window
'        Dim otherWinWs As Excel.Worksheet
'        Dim freezeHeader As Boolean, freezePrimCols As Boolean
'        Dim recCountCheck As ADODB.Recordset 'for checking record count in the underlying table

'        On Error GoTo err1
'100     readData = False
'102     If Not readHeaders Then Exit Function
'104     If ws.Parent.Windows.Count = 2 Then
'106         If ws.Parent.Windows(1).ActiveSheet.name = ws.name Then
'108             Set otherWin = ws.Parent.Windows(2)
'110             Set otherWinWs = ws.Parent.Windows(2).ActiveSheet
'            Else
'112             Set otherWin = ws.Parent.Windows(1)
'114             Set otherWinWs = ws.Parent.Windows(1).ActiveSheet
'            End If
'        End If
'116     With ws
'118         freezeHeader = (fetchSetting(DBSheetID & "." & "freezeHeader", "OK", True) = "OK")
'120         freezePrimCols = (fetchSetting(DBSheetID & "." & "freezePrimCols", "OK", True) = "OK")
'122         queryStr = config.getEntry("query")
'124         If Not replaceQueryParameters(queryStr) Then
'126             LogWarn "couldn't parse parameters in paramstring", True
'128             readData = False
'                Exit Function
'        End If
'130         If Not readLookups Then Exit Function
'132         If Not theDBSheetConnection.openConnection Then Exit Function
'134         .Activate
'            Dim previousSplitRow As Long, previousSplitColumn As Long
'136         previousSplitRow = theHostApp.ActiveWindow.SplitRow
'138         previousSplitColumn = theHostApp.ActiveWindow.SplitColumn
'140         theHostApp.ActiveWindow.FreezePanes = False
'142         .Range(.Cells(2, 1), .Cells(.Rows.Count, .Columns.Count)).ClearContents
'            ' prepare autoformatting text formatting down
'144         If (fetchSetting(DBSheetID & "." & "autoformatCellsDown", vbNullString, True) = "OK") Then
'                Dim i As Long
'146             For i = 1 To theFields.Count
'148                 ReDim Preserve copyFormat(i)
'                    ' either fetch formatting (datum, etc.) from stored setting or (if not existing) take from first data row
'150                 copyFormat(i) = fetchSetting(DBSheetID & "." & .Cells(1, i).Text, .Cells(2, i).NumberFormat, True)
'                Next
'        End If
'        ' now read data into DBSheet using query given in queryStr
'152         Set rst = New Recordset
'154         rst.CursorLocation = adUseClient
'156         rst.Open queryStr, dbcnn, adOpenForwardOnly, adLockReadOnly, -1
'158         With .QueryTables.Add(Connection:=rst, Destination:=.Range("A2"))
'160             .FieldNames = False
'162             .PreserveFormatting = True
'164             .FillAdjacentFormulas = False
'166             .AdjustColumnWidth = False
'168             .RefreshStyle = xlOverwriteCells  ' this is required to prevent "right" shifting of cells!
'170             .RefreshPeriod = 0
'172             .PreserveColumnInfo = True
'174             tmpname = .name
'176             .Refresh
'178             .Delete
'            End With
'        ' sometimes excel doesn't delete the querytable given name
'        On Error Resume Next
'180         .Names(tmpname).Delete
'182         .Parent.Names(tmpname).Delete
'            On Error GoTo err1
'184         maxRows = 1 + rst.RecordCount
'            ' autoformatting down
'186         If (fetchSetting(DBSheetID & "." & "autoformatCellsDown", vbNullString, True) = "OK") Then
'188             .Rows("3:" & .Rows.Count).ClearFormats
'                ' autoformat: restore format of 1st row...
'190             For i = 1 To theFields.Count
'192                 .Cells(2, i).NumberFormat = copyFormat(i)
'                Next
'        'fill a few rows (10) more down, to have formatting also for new data entered below...
'194             .Rows(2).AutoFill Destination:=.Rows("2:" & IIf(maxRows + 10 > .Rows.Count, .Rows.Count, maxRows + 10)), Type:=xlFillFormats
'            End If
'        ' other general set formats:
'196         .Rows("2:" & .Rows.Count).Interior.ColorIndex = xlNone
'198         .Rows("2:" & .Rows.Count).Font.Color = 0
'            ' this format is not possible as it is reserved for deletions
'200         .Rows("2:" & .Rows.Count).Font.Strikethrough = False
'            ' this format is not possible as it is reserved for data changes
'202         .Rows("2:" & .Rows.Count).Font.Italic = False

'            ' enable all cells for editing
'204         .Cells.Locked = False
'            ' except header row  !!
'206         .Rows(1).Locked = True
        
'            ' create action storage (for i/c/d flags)
'208         Dim actionwsname As String: actionwsname = makeAuxSheet("A", True)
'210         .Parent.Worksheets(actionwsname).Visible = xlSheetVeryHidden
'            ' produce copy of DB sheet for checking stale data (multiuser editing !) into sheet D<wsname> (hidden)
'212         olddatawsname = makeAuxSheet("D", True)
'214         .Rows("2:" & IIf(maxRows + 1 > ws.Rows.Count, ws.Rows.Count, maxRows + 1)).Copy
'216         .Parent.Worksheets(olddatawsname).Visible = xlSheetVisible
'218         .Parent.Worksheets(olddatawsname).Select
'220         .Parent.Worksheets(olddatawsname).Range("A2").Select
'            ' avoid warning messages about names already contained in workbook
'            ' (when renaming workbook, old names still refer to ranges and cause this)
'222         theHostApp.DisplayAlerts = False
'224         .Parent.Worksheets(olddatawsname).Paste
'226         theHostApp.CutCopyMode = False
'228         theHostApp.DisplayAlerts = True
'230         .Parent.Worksheets(olddatawsname).Visible = xlSheetVeryHidden

'232         .Select
'            ' copy stored formulas down for calculated cells
'234         If calcColumnsStart <> 0 Then
'236             For i = 0 To UBound(formulaCells)
'238                 .Range(.Cells(2, calcColumnsStart + i), .Cells(IIf(maxRows + 1 > .Rows.Count, .Rows.Count, maxRows + 1), calcColumnsStart + i)).Formula = formulaCells(i)
'                Next
'        End If

'        ' enter validation lists for foreign key lookup columns
'        Dim theField
'240         For Each theField In theFields.Fields
'242             If Not theFields.LookupSheet(theField) Is Nothing Then
'                    On Error Resume Next  ' as there are also empty lookup restrictions possible !!
'244                 With .Range(.Cells(2, theFields.Column(theField)), .Cells(.Rows.Count, theFields.Column(theField))).Validation
'246                     .Delete
'                        ' avoid errors when adding reference to lookup
'248                     .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & theField
'250                     .IgnoreBlank = True
'252                     .InCellDropdown = True
'254                     .InputTitle = vbNullString
'256                     .ErrorTitle = "Ungültig !!"
'258                     .InputMessage = vbNullString
'260                     .ErrorMessage = "Nur Auswahlwerte (siehe Dropdown) für Feld: " & ws.Cells(1, theFields.Column(theField))
'262                     .ShowInput = False
'264                     .ShowError = True
'                    End With
'        On Error GoTo err1
'        End If
'        Next
'        ' coulour and lock primary column(s)...
'266         If maxRows > 1 Then  ' special case: no data -> no locking!!
'268             If maxRows > .Rows.Count Then maxRows = .Rows.Count
'                ' lock/colour primary key columns only for existing data (new rows are free !)
'270             .Range(.Cells(2, 1), .Cells(maxRows, primKeyColumns)).Locked = True
'272             .Range(.Cells(2, 1), .Cells(maxRows, primKeyColumns)).Font.ColorIndex = primColFieldsColor
'            End If
'        ' split/freeze header row and primary column(s)
'274         theHostApp.ActiveWindow.SplitRow = 0
'276         theHostApp.ActiveWindow.SplitColumn = 0
        
'            ' this is necessary for the true row/column to be split, splitrow/splitcolumn refer to the currently visible rows/columns
'278         theHostApp.ScreenUpdating = True
'280         .Cells(1, 1).Activate
'282         theHostApp.ScreenUpdating = False
'284         If freezeHeader Then
'286             theHostApp.ActiveWindow.SplitRow = 1
'            Else
'288             theHostApp.ActiveWindow.SplitRow = previousSplitRow
'            End If
'290         If freezePrimCols Then
'292             theHostApp.ActiveWindow.SplitColumn = primKeyColumns
'            Else
'294             theHostApp.ActiveWindow.SplitColumn = previousSplitColumn
'            End If
'296         If theHostApp.ActiveWindow.SplitColumn <> 0 Or theHostApp.ActiveWindow.SplitRow <> 0 Then
'298             theHostApp.ActiveWindow.FreezePanes = True
'            End If
'        End With
'300     If theHostApp.ActiveWorkbook.Windows.Count = 2 Then
'302         Set curWin = theHostApp.ActiveWindow
'304         otherWin.Activate
'306         otherWinWs.Activate
'308         curWin.Activate
'        End If
'310     readData = True
'        ' check whether query retrieved less recrds than table contains and warn user !!
'312     Set recCountCheck = New ADODB.Recordset
'314     recCountCheck.Open "SELECT count(*) FROM " & config.getEntry("table"), dbcnn, adOpenStatic, adLockReadOnly
'        Dim recCount As Long
'315     recCountCheck.MoveFirst
'316     recCount = recCountCheck.Fields(0)
'317     If recCount <> maxRows - 1 Then
'318         Call confirmOK("Did not retrieve everything from '" & config.getEntry("table") & "' (retrieved:" & maxRows - 1 & ", actual: " & recCount & "), please check parameter settings and DBSheet query if this is not what you expected !", "dontShowRecCount")
'        End If
'320     recCountCheck.Close
'322     Set recCountCheck = Nothing
'        Exit Function
'err1:
'324     If VBDEBUG Then Debug.Print Err.Description: Stop : Resume
'326     LogError "Error: " & Err.Description & ", line " & Erl & " in DBSheetHandler.readData"
'End Function

'    ''
'    ' when we enter/modify any data we need to "mark" this row as edited (either (i)nserted or (c)hanged)....
'    ' @param Sh
'    ' @param Target
'    ' @remarks
'    Private Sub theapp_SheetChange(ByVal Sh As Object, ByVal Target As Excel.Range)
'        Dim theCell As Excel.Range
'        Dim actionwsname As String

'        On Error GoTo err1
'        ' as there are two instances of DBSheethandlers in foreignSheetMode, we need to filter event calls to get just the active window's dbsheet handler !
'100     If Not noDBSheetEvent And active And Sh Is ws Then
'            ' warn if we're about to insert/remove a whole column resulting in lots of change marks....
'102         If Target.Address = Target.Columns.EntireColumn.Address Then
'104             If Not confirmOK("Mark changes for column modification? This will result in marks being set for EVERY ROW !", "doWholeColumn", False, False, False, True) Then Exit Sub
'            End If
'106         theHostApp.EnableEvents = False
'108         actionwsname = makeAuxSheet("A", False)
'110         For Each theCell In Sh.Range(Sh.Cells(Target.row, 1), Sh.Cells(Target.row + Target.Rows.Count - 1, 1))
'112             If theCell.row <> 1 And IsEmpty(theHostApp.Worksheets(actionwsname).Cells(theCell.row, 1)) Or _
'                   theHostApp.Worksheets(actionwsname).Cells(theCell.row, 1) = "p" Then
'114                 If theCell.row > maxRows Or _
'                       checkIsEmptyRow(Sh.Range(Sh.Cells(theCell.row, 1), Sh.Cells(theCell.row, primKeyColumns))) Or _
'                       theHostApp.Worksheets(actionwsname).Cells(theCell.row, 1) = "p" Then
'116                         theHostApp.Worksheets(actionwsname).Cells(theCell.row, 1) = "i"
'                    Else
'118                     theHostApp.Worksheets(actionwsname).Cells(theCell.row, 1) = "c"
'                    End If
'        ' Änderungen blau
'120                 theCell.EntireRow.Font.Color = 16711680
'122                 DbsheetChanged = True
'124                 theHostApp.StatusBar = "Marking row " & theCell.row
'                End If
'126             Target.Font.Italic = True
'            Next
'128         theHostApp.EnableEvents = True
'130         theHostApp.StatusBar = False
'        End If
'        Exit Sub
'err1:
'        theHostApp.EnableEvents = True
'132     If VBDEBUG Then Debug.Print Err.Description: Stop : Resume
'134     LogError "Error: " & Err.Description & ", line " & Erl & " in DBSheetHandler.theapp_SheetChange"
'End Sub

'    ''
'    ' called after right clicking row and selecting "insert Row"
'    Public Sub insertRow()
'        Dim fromCell As Long, toCell As Long, i As Long
'        Dim actionwsname As String
'        Dim olddatawsname As String

'        On Error Resume Next
'        ' don't insert for rows where not necessary
'100     If theHostApp.Selection.row > maxRows Then Exit Sub
'102     theHostApp.EnableEvents = False
'        ' expected error: when having an open validation dropdown, this (and the rest of procedure) fails !!
'104     If Err <> 0 Then
'106         LogError "Can't insert row while dropdown is open !!"
'108         GoTo cleanup
'        End If

'        On Error GoTo err1
'110     If confirmOK("Really insert rows ?", "dontShowinsert", , False) Then
'112         With ws
'114             actionwsname = makeAuxSheet("A", False)
'116             olddatawsname = makeAuxSheet("D", False)
'118             primKeyColumns = CLng(config.getEntry("primcols"))
'120             .Unprotect
'122             fromCell = theHostApp.Selection.row
'124             toCell = fromCell + theHostApp.Selection.Rows.Count - 1
'126             For i = fromCell To toCell
'128                 .Rows(i).Insert
'130                 .Parent.Worksheets(olddatawsname).Rows(i).Insert
'132                 .Parent.Worksheets(actionwsname).Rows(i).Insert
'                    ' always allow edit of new primary keys (either it's ignored with identity fields or it get's written...)
'134                 .Range(.Cells(i, 1), .Cells(i, primKeyColumns)).Locked = False
'                    ' mark as potential insert !
'136                 .Parent.Worksheets(actionwsname).Cells(i, 1) = "p"
'138                 maxRows = maxRows + 1
'                Next
'140             DbsheetChanged = True
'142             protectTheSheet
'            End With
'        End If
'144     GoTo cleanup
'        Exit Sub
'err1:
'146     If VBDEBUG Then Debug.Print Err.Description: Stop : Resume
'148     LogError "Error: " & Err.Description & ", line " & Erl & " in DBSheetHandler.insertRow"
'cleanup:
'150     theHostApp.EnableEvents = True
'End Sub

'    ''
'    ' stores all data marked  as i(nsert),d(elete),c(hange) into database
'    ' @return whether function was successful
'    ' @remarks class member ws = the CURRENT worksheet to be stored
'    Public Function storeData() As Boolean
'        Dim i As Long, j As Long, k As Long
'        Dim exitMe As Boolean, enforceSaveAllWhenCalcColumns As Boolean
'        Dim actionwsname As String

'        On Error GoTo err1
'100     ws.Unprotect
'102     storeData = True
'104     noDBSheetEvent = True

'106     With ws
'108         tableName = config.getEntry("table")
'110         actionwsname = makeAuxSheet("A", False)
'            On Error Resume Next
'        Dim maxNumberMassChange As Long
'112         maxNumberMassChange = CLng(fetchSetting("maxNumberMassChange", "50"))
'114         If theApp.WorksheetFunction.CountA(theHostApp.Worksheets(actionwsname).Columns(1)) > maxNumberMassChange Then
'116             If Err = 0 Then
'118                 If Not confirmOK("There are more than " & CStr(maxNumberMassChange) & " changes, really apply ?", "chkMassChange", , False, , , True) Then
'120                     MsgBox "Cancelling Updates and refreshing DBSheet !!!", vbOKOnly + vbExclamation, "DBAddin: Cancelling Updates"
'122                     GoTo cleanupOK
'                    End If
'        End If
'        End If
'        On Error GoTo err1
'124         .Select
'126         enforceSaveAllWhenCalcColumns = (config.getEntry("enforceSaveAllWhenCalcColumns") = "OK")
'            ' find first startpoint, when calculating and forcing save, start from row 2
'128         theHostApp.Worksheets(actionwsname).Cells(1, 1).Clear

'130         If enforceSaveAllWhenCalcColumns And calcColumnsStart > 0 Then
'132             j = 2
'            Else
'134             j = theHostApp.Worksheets(actionwsname).Cells(1, 1).End(xlDown).row
'            End If

'136         Do While j < .Rows.Count
'                ' then next endpoint
'138             If enforceSaveAllWhenCalcColumns And calcColumnsStart > 0 Then
'                    ' if calculate and enforce save all then until the end (of all primary keys)
'140                 k = .Cells(j, 1).End(xlDown).row
'142             ElseIf IsEmpty(theHostApp.Worksheets(actionwsname).Cells(j + 1, 1)) Then
'                    'if only one row, equal to startpoint
'144                 k = j
'                Else
'        On Error Resume Next
'146                 k = theHostApp.Worksheets(actionwsname).Cells(j, 1).End(xlDown).row

'148                 If Err = 6 Then k = j + 1
'                    On Error GoTo err1
'        End If

'        ' all rows between the 2 points
'150             For i = j To k
'152                 If Not (checkIsEmptyRow(.Range(.Cells(i, 1), .Cells(i, .Columns.Count))) Or theHostApp.Worksheets(actionwsname).Cells(i, 1) = "p") Then
'154                     If Not modifyInDatabase(theHostApp.Worksheets(actionwsname).Cells(i, 1), i, primKeyColumns, exitMe) Then storeData = False
'156                     If exitMe Then GoTo cleanup
'                    End If
'        Next

'158             If enforceSaveAllWhenCalcColumns And calcColumnsStart > 0 Then
'                    ' finish in case of calcCol enforcement as we have already worked on all rows in the for loop !
'                    Exit Do
'        Else
'        ' find next startpoint
'160                 j = theHostApp.Worksheets(actionwsname).Cells(k, 1).End(xlDown).row
'                End If
'        Loop
'162         noDBSheetEvent = False
'        End With
'164     theHostApp.StatusBar = False
'        Exit Function
'err1:

'166     If VBDEBUG Then Debug.Print Err.Description: Stop : Resume
'168     LogError "Error: " & Err.Description & ", line " & Erl & " in DBSheetHandler.storeData"
'cleanup:
'170     storeData = False
'cleanupOK:
'172     noDBSheetEvent = False
'174     theHostApp.StatusBar = False
'End Function

'    ''
'    ' modifies the data in row modifRow into database as denoted by action (i(nsert),d(elete),c(hange)) having primKeyColumns primary key columns
'    ' @param actionCell cell containing action indo (i(nsert),d(elete),c(hange))
'    ' @param modifRow row that is modified
'    ' @param primKeyColumns how many primary key columns do we have
'    ' @param exitMe reference flag to indicate to calling function whether whole storing process should be exited...
'    ' @return whether function was successful
'    ' @remarks
'    ' data modification is first checked using ADO with datarst, this is undone, the actual modification<br>
'    ' is done by creating INSERT/UPDATE/DELETE statements. This lets the underlying database return the<br>
'    ' full & correct error message for the row...<br>
'    Private Function modifyInDatabase(actionCell As Excel.Range,
'                                  ByVal modifRow As Long,
'                                  ByVal primKeyColumns As Long,
'                                  ByRef exitMe As Boolean) As Boolean
'        Dim datarst As ADODB.Recordset  ' recordset where data is test-modified (real modification done in SQL statements))
'        Dim checkrst As ADODB.Recordset  ' recordset for checking datatypes
'        Dim theVal As String, theOldVal As String, displayVal As String, action As String, actionName As String, olddatawsname As String, dataStatement As String, theCriteria As String, theHead As String
'        Dim columnError As Boolean, autoIncrement As Boolean
'        Dim startColumn As Long

'        On Error GoTo err1
'100     exitMe = False
'102     theHostApp.StatusBar = "Storing Row " & modifRow & " to database..."
'104     modifyInDatabase = True
'106     startColumn = primKeyColumns + 1  ' column where we start to get values from (usually after the primary keys)
'108     With ws
'110         action = actionCell.Value
'112         actionName = Switch(action = "i", "insert", action = "d", "delete", action = "c", "update")
'            ' checkrst is opened to get information about table schema (field types)
'114         Set checkrst = New ADODB.Recordset
'116         checkrst.Open tableName, dbcnn, adOpenForwardOnly, adLockReadOnly
'            ' place where original data is stored...
'118         olddatawsname = "D" & IIf(Len(.CodeName) > 30, Left$(.CodeName, 30), .CodeName)
'120         columnError = False  ' any errors in columns ?

'            ' FIRST: produce dataStatement used for retrieving the record to be updated or deleted
'            ' for insert this is simply the whole table
'122         If action = "i" Then
'124             autoIncrement = False
'                ' two cases of primary key filling:
'                ' a) DB engine autoincrements primary key: leave primary key and just insert record...
'126             If checkrst.Fields(.Cells(1, 1).Value).Properties("ISAUTOINCREMENT") And primKeyColumns = 1 Then
'128                 .Cells(modifRow, 1).Value = vbNullString: startColumn = 2 : autoIncrement = True
'        ' b) more than one primary key or selectable prim key, ALL fields (of primary keys) need to be filled...
'        Else
'130                 startColumn = 1
'                End If
'132             dataStatement = "SELECT * FROM " & tableName
'            Else
'        'for delete or update produce primary column criteria for selecting record ..
'134             theCriteria = vbNullString
'                Dim c As Long
'136             For c = 1 To primKeyColumns
'138                 theVal = .Parent.Worksheets(olddatawsname).Cells(modifRow, c).Value
'140                 theHead = .Cells(1, c).Value
'                    ' resolve foreign key lookups in primary key columns...
'142                 If Not theFields.LookupSheet(theHead) Is Nothing Then
'144                     columnError = Not resolveID(theHead, theVal, exitMe, True)
'                        ' if the primary column contains values that are not found, take the direct values instead
'146                     If columnError And Not IsEmpty(.Parent.Worksheets(olddatawsname).Cells(modifRow, c).Value) Then
'148                         theVal = .Parent.Worksheets(olddatawsname).Cells(modifRow, c).Value
'150                         columnError = False
'                            ' set start column to non-found key to enable update of primary key as well..
'152                         If startColumn > c Then startColumn = c
'                        Else
'154                         modifyInDatabase = Not columnError
'156                         If exitMe Then
'158                             checkrst.Close
'                                Exit Function
'        End If
'        End If
'        End If
'160                 theCriteria = theCriteria & theHead & " = " & dbformat(theVal, theHead, checkrst) & " AND "
'                Next
'162             theCriteria = Left$(theCriteria, Len(theCriteria) - Len(" AND "))
'164             dataStatement = "SELECT * FROM " & tableName & " WHERE " & theCriteria
'            End If

'166         If Not columnError Then
'                ' SECOND: check with datarst if
'                ' 1) single cells are being invalid (and mark them red)
'                ' 2) data has been changed between initial refresh and store (stale data warning)
'168             Set datarst = New ADODB.Recordset
'170             datarst.Open dataStatement, dbcnn, adOpenKeyset, adLockOptimistic
                       
'                ' row inserted or updated
'                Dim fieldAssignList As String, fieldList As String, valueList As String
'172             fieldAssignList = vbNullString: fieldList = vbNullString : valueList = vbNullString
'174             If action = "i" Then datarst.AddNew
'                Dim col As Long
'176             For col = startColumn To theFields.Count
'178                 theOldVal = .Parent.Worksheets(olddatawsname).Cells(modifRow, col).Value
'180                 theVal = .Cells(modifRow, col).Value
'182                 theHead = .Cells(1, col).Value
'                    Dim fieldIsBoolean As Boolean
'184                 fieldIsBoolean = False
'186                 If TypeName(.Cells(modifRow, col).Value) = "Boolean" Then fieldIsBoolean = True
'188                 If checkIsDateTime(checkrst.Fields(theHead).Type) Then
'                        On Error Resume Next
'190                     theOldVal = CDate(theOldVal)
'192                     Err.Clear
'194                     If LenB(theVal) > 0 Then theVal = CDate(theVal)
'196                     If Err <> 0 Then
'198                         modifyInDatabase = False
'200                         exitMe = True
'202                         LogWarn "date type expected !!", exitMe
'204                         If exitMe Then
'206                             cleanUpAfterErrs datarst, checkrst
'                                Exit Function
'        End If
'        End If
'        On Error GoTo err1
'        End If
'208                 If Not theFields.LookupSheet(theHead) Is Nothing Then
'210                     If action = "i" Or action = "c" Then
'212                         columnError = Not resolveID(theHead, theVal, exitMe)
'                        End If
'        ' not found old IDs are no problem for deletions (need to be fetched but not to be checked)
'        ' also ignore in case of previous column error
'        Dim returnResolve As Boolean
'214                     returnResolve = resolveID(theHead, theOldVal, exitMe)
'216                     If LenB(theVal) > 0 And Not columnError Then columnError = Not returnResolve
'218                     modifyInDatabase = Not columnError
'220                     If exitMe Then
'222                         cleanUpAfterErrs datarst, checkrst
'                            Exit Function
'        End If
'        End If
'224                 If (action = "d" Or action = "c") And Not columnError Then
'                        ' did somebody change the field's data in the meantime ?
'226                     If Trim(datarst.Fields(theHead).Value) <> theOldVal Then ' don't do Trim$ here as this should be variant (NULL, etc. !!)
'                            ' undo delete ?
'228                         If action = "d" Then
'230                             If theOldVal <> theVal Then
'                                    Dim retval As Integer
'232                                 retval = MsgBox("Record to be deleted was edited by somebody else (" & .Cells(1, col).Value & ":" & datarst.Fields(theHead).Value & ")" & vbCrLf & "Do you want to delete the record definitely ('No' keeps those edits) ?", vbYesNo + vbInformation, "DBAddin: Data Edit Conflict")
'234                                 If retval = vbNo Then
'                                        ' revert both deleted and old data to data in DB...
'236                                     .Cells(modifRow, col).EntireRow.Font.Strikethrough = False
'238                                     For c = 1 To theFields.Count
'240                                         displayVal = CStr(datarst(.Cells(1, c).Value).Value)
'242                                         If Not theFields.LookupSheet(.Cells(1, c).Value) Is Nothing Then
'244                                             columnError = Not resolveLookup(.Cells(1, c), displayVal, exitMe)
'246                                             If columnError Then
'248                                                 modifyInDatabase = False
'250                                                 If exitMe Then
'252                                                     cleanUpAfterErrs datarst, checkrst
'                                                        Exit Function
'        End If
'        End If
'        End If
'254                                         If Not columnError Then
'256                                             .Parent.Worksheets(olddatawsname).Cells(modifRow, c).Value = displayVal
'258                                             .Cells(modifRow, c).Value = displayVal
'                                            End If
'        Next
'260                                     datarst.Close
'262                                     checkrst.Close
'264                                     modifyInDatabase = True
'                                        Exit Function
'        End If
'        End If
'        Else
'        ' undo single field changes ?
'266                             If theOldVal <> theVal Then
'                                    ' did we also change the field's data (conflict resolution needed) ?
'268                                 .Cells(modifRow, 1).Interior.ColorIndex = conflictColor
'270                                 .Cells(modifRow, col).Select
'272                                 displayVal = CStr(datarst.Fields(theHead).Value)
'274                                 If Not theFields.LookupSheet(theHead) Is Nothing Then
'276                                     columnError = Not resolveLookup(theHead, displayVal, exitMe)
'278                                     If columnError Then
'280                                         modifyInDatabase = False
'282                                         If exitMe Then
'284                                             cleanUpAfterErrs datarst, checkrst
'                                                Exit Function
'        End If
'        End If
'        End If
'286                                 retval = MsgBox("Field:'" & .Cells(1, col).Value & "' to be stored (" & .Cells(modifRow, col).Value & ") was also edited by somebody else (" & displayVal & ")." & vbLf & "Do you want to overwrite this with your changes ('No' keeps those edits) ?", vbYesNo + vbInformation, "DBAddin: Data Edit Conflict")
'288                                 If retval = vbNo Then
'290                                     If Not columnError Then
'292                                         .Parent.Worksheets(olddatawsname).Cells(modifRow, col).Value = displayVal
'294                                         .Cells(modifRow, col).Value = displayVal
'296                                         theVal = datarst.Fields(theHead).Value
'                                        End If
'        End If
'298                                 .Cells(modifRow, 1).Interior.ColorIndex = -4142 '(automatic)
'                                    ' resume updating and don't care about old lookup errors..
'300                                 modifyInDatabase = True
'302                                 columnError = False
'                                    ' check for new lookup errors introduced ..
'304                                 If Not theFields.LookupSheet(theHead) Is Nothing Then
'306                                     columnError = Not resolveLookup(theHead, theVal, exitMe)
'308                                     If columnError Then
'310                                         modifyInDatabase = False
'312                                         If exitMe Then
'314                                             cleanUpAfterErrs datarst, checkrst
'                                                Exit Function
'        End If
'        End If
'        End If
'        Else
'316                                 theVal = datarst.Fields(theHead).Value                                     ' don't care if we didn't change the field
'                                End If
'        End If
'        End If
'        End If
'        ' convert booleans in binary fields to 1 or 0
'318                 If fieldIsBoolean Then theVal = IIf(theVal, 1, 0)
                
'320                 If action = "i" And Not columnError Then
'                        ' check for empty primary columns...
'322                     If LenB(theVal) = 0 And col <= primKeyColumns And Not autoIncrement Then
'324                         exitMe = True: modifyInDatabase = False
'326                         LogWarn "inserting row for multiple primary keys: not all primary keys are given !!", exitMe
'328                         If exitMe Then
'330                             cleanUpAfterErrs datarst, checkrst
'                                Exit Function
'        End If
'332                         columnError = True
'                        End If
'        ' leave a chance to fill in column default values for inserts,
'        ' nulls are inserted anyway by default...
'334                     If LenB(theVal) > 0 Then
'336                         fieldList = fieldList & theHead & ","
'338                         valueList = valueList & dbformat(theVal, theHead, checkrst) & ","
'                        End If
'340                 ElseIf action = "c" Then
'342                     If LenB(theVal) = 0 Then
'344                         fieldAssignList = fieldAssignList & theHead & "=NULL,"
'                        Else
'346                         fieldAssignList = fieldAssignList & theHead & "=" & dbformat(theVal, theHead, checkrst) & ","
'                        End If
'        End If

'        ' try to insert the values in the recordset for marking single cells as errors
'348                 If action <> "d" Then
'                        Dim testVal As Variant
'350                     testVal = theVal
'352                     If LenB(theVal) = 0 Then testVal = Empty
'                        On Error Resume Next
'354                     datarst.Fields(theHead).Value = testVal
'356                     If Err <> 0 Then
'358                         columnError = True: modifyInDatabase = False
'360                         .Cells(modifRow, col).Select
'362                         .Cells(modifRow, col).Interior.ColorIndex = dataErrColor
'364                         .Cells(modifRow, 1).Interior.ColorIndex = dataErrColor
'366                         exitMe = True
'368                         LogWarn "Error in field: " & .Cells(1, col).Value & ", value: " & .Cells(modifRow, col).Value & ", error: " & Err.Description, exitMe
'370                         If exitMe Then
'372                             cleanUpAfterErrs datarst, checkrst
'                                Exit Function
'        End If
'        Else
'374                         .Cells(modifRow, col).Interior.ColorIndex = xlColorIndexNone
'                        End If
'        On Error GoTo err1
'        End If
'        Next
'376             datarst.CancelUpdate
'            End If

'378         If Not columnError Then
'                ' THIRD: construct & execute the final data insert/set/delete statement
'                Dim statement As String
'380             If action = "d" Then
'382                 statement = "DELETE FROM " & tableName & " WHERE " & theCriteria
'                Else
'384                 If action = "i" Then
'                        ' prepend the primary key name and value before the insert list
'386                     fieldList = Left$(fieldList, Len(fieldList) - 1)
'388                     valueList = Left$(valueList, Len(valueList) - 1)
'390                     statement = "INSERT INTO " & tableName & " (" & fieldList & ") VALUES (" & valueList & ")"
'392                 ElseIf action = "c" Then
'394                     fieldAssignList = Left$(fieldAssignList, Len(fieldAssignList) - 1)
'396                     statement = "UPDATE " & tableName & " SET " & fieldAssignList & " WHERE " & theCriteria
'                    End If
'        End If
'        On Error Resume Next
'        Dim recordsAffected As Long
'398             dbcnn.Execute statement, recordsAffected
'400             If Err <> 0 Or recordsAffected > 1 Then
'402                 modifyInDatabase = False
'404                 exitMe = True ' allow user to exit after error warning...
'406                 .Rows(modifRow).Select
'408                 If recordsAffected > 1 Then
'410                     LogWarn "more than one row (" & recordsAffected & " rows) were affected with: " & statement, exitMe
'                    Else
'412                     LogWarn "Error in data " & actionName & ": " & Err.Description & ", statement: " & statement, exitMe
'                    End If
'414                 If exitMe Then
'416                     cleanUpAfterErrs datarst, checkrst
'                        Exit Function
'        End If
'        Else
'418                 .Rows(modifRow).Interior.ColorIndex = xlColorIndexNone
'420                 .Rows(modifRow).Font.Bold = False
'422                 actionCell = Empty
'                    #If DEBUGME = 1 Then
'424                     LogInfo actionName & " row, statement: " & statement
'                    #End If
'                End If
'        End If
'        On Error GoTo err1
'        End With

'426     datarst.Close
'428     checkrst.Close
'        Exit Function
'err1:
'430     modifyInDatabase = False
'432     If VBDEBUG Then Debug.Print Err.Description: Stop : Resume
'434     LogError "Error: " & Err.Description & ", line " & Erl & " in DBSheetHandler.modifyInDatabase, statement: " & statement & ", dataStatement: " & dataStatement
'End Function


'    ''
'    ' replaces the parameter placeholders ("?") in queryStr with the contents of the assigned ranges
'    ' using the parameter list stored in dbsheet definition of ws.
'    ' returns false on failure, queryStr contains updated query.
'    ' @param queryStr
'    ' @return whether function was successful
'    ' @remarks
'    Private Function replaceQueryParameters(ByRef queryStr As String) As Boolean
'        Dim ParamsList, teststr

'        On Error GoTo replaceQueryParameters_Err
'        replaceQueryParameters = True
'        If InStr(1, queryStr, "?") = 0 And Len(config.getEntry("params")) = 0 Then Exit Function

'        ' read params list, format: <params><p><rng>param1RangeName</rng></p><p><rng>param2RangeName</rng></quote></p>....</params>
'        ParamsList = config.getEntryList("params", "p", vbNullString)
'        If IsNothing(ParamsList) Then
'            replaceQueryParameters = False
'            Exit Function
'        End If
'        ' quoted replace of "?" needs splitting of queryStr by quotes !
'        teststr = Split(queryStr, "'")
'        queryStr = vbNullString
'        ' walk through splitted parts and replace "?" one by one in even (unquoted) ones
'        Dim j As Long
'        j = 0
'        Dim i As Long
'        Dim subresult As String
'        For i = 0 To UBound(teststr)
'            If i Mod 2 = 0 Then
'                Dim replacedStr As String
'                replacedStr = teststr(i)
'                While InStr(1, replacedStr, "?") > 0
'                    If j = UBound(ParamsList) + 1 Then
'                        replaceQueryParameters = False
'                        Exit Function
'                    End If
'                    Dim paramRange As String, paramVal As String
'                    Dim paramQuote As Boolean, questionMarkLoc As Long
'                    paramRange = config.getEntry("rng", ParamsList(j))
'                    paramVal = ExcelDnaUtil.Application.ActiveWorkbook.Names(paramRange).RefersToRange.Value
'                    If Len(paramVal) = 0 Then
'                        LogWarn("No entry in range " & paramRange & " for parameter " & j & " of DBsheet Query !!")
'                        replaceQueryParameters = False
'                        Exit Function
'                    End If
'                    paramQuote = InStr(1, ParamsList(j), "</quote>") > 0
'                    questionMarkLoc = InStr(1, replacedStr, "?")
'                    replacedStr = Mid$(replacedStr, 1, questionMarkLoc - 1) & IIf(paramQuote, "'", vbNullString) & paramVal & IIf(paramQuote, "'", vbNullString) & Mid$(replacedStr, questionMarkLoc + 1)
'                    j = j + 1
'                End While
'                subresult = replacedStr
'            Else
'                subresult = teststr(i)
'            End If
'            queryStr = queryStr & subresult & IIf(i < UBound(teststr), "'", vbNullString)
'        Next
'        Exit Function

'replaceQueryParameters_Err:
'        LogError("Error: " & Err.Description & ", line " & Erl() & " in DBSheetHandler.replaceQueryParameters")
'final:
'    End Function

'    ''
'    ' parses myStr for substrings separated by separator string sep
'    ' result is returned in array of strings list
'    ' @param myStr
'    ' @param list values are returned here !
'    ' @param sep
'    ' @remarks
'    Private Sub parse(ByVal myStr As String, ByRef list() As String, sep As String)
'        Dim last, start, leng, i As Integer
'        Dim exitloop As Boolean
'        Dim part As String

'        On Error GoTo parse_Err
'        last = 1
'        start = 1
'        i = 0
'        Do
'            leng = InStr(start + 1, myStr, sep) - start
'            If leng <= 0 Then
'                leng = Len(myStr) - start + 1
'                exitloop = True
'            End If
'            part = Mid$(myStr, start, leng)
'            start = InStr(start + leng, myStr, sep) + Len(sep)
'            ReDim Preserve list(i)
'            list(i) = part
'            i = i + 1
'        Loop Until exitloop
'        Exit Sub
'parse_Err:
'        LogError("Error: " & Err.Description & ", line " & Erl() & " in DBSheetHandler.parse")
'    End Sub

'End Class