Imports Microsoft.Office.Interop.Excel
Imports ExcelDna.Integration
Imports ADODB
Imports System.Timers

''' <summary>main calculation event handling and Data retrieving</summary>
Public Class DBFuncEventHandler

    ''' <summary>connection string can be changed for calls with different connection strings</summary>
    Public CurrConnString As String
    ''' <summary>to work around a silly Excel bug with Dirty Method we have to select the sheet with the "dirtied" cell to actually do the dirtification.
    ''' to return to the target, need the original worksheet here.</summary>
    Public origWS As Worksheet
    ''' <summary>cnn object always the same (only open/close)</summary>
    Public cnn As ADODB.Connection
    ''' <summary>the app object needed for excel event handling (most of this class is decdicated to that)</summary>
    Private WithEvents Application As Application

    ''' <summary>which error occurred?</summary>
    Private errorReason As String

    Private ODBCconnString As String
    ''' <summary>query cache for avoiding unnecessary recalculations/data retrievals</summary>
    Public queryCache As Collection

    Public Sub New()
        Application = ExcelDnaUtil.Application
        queryCache = New Collection
    End Sub

    ''' <summary>necessary to asynchronously start refresh of db functions after save event</summary>
    Private aTimer As System.Timers.Timer

    Private Sub App_WorkbookOpen(ByVal Wb As Workbook) Handles Application.WorkbookOpen
        If Not Wb.IsAddin Then
            Dim refreshDBFuncs As Boolean
            ' when opening, force recalculation of DB functions in workbook.
            ' this is required as there is no recalculation if no dependencies have changed (usually when opening workbooks)
            ' however the most important dependency for DB functions is the database data....
            Try
                refreshDBFuncs = Not Wb.CustomDocumentProperties("DBFskip")
            Catch ex As Exception
                refreshDBFuncs = True
            End Try
            If refreshDBFuncs Then refreshDBFunctions(Wb)
            repairLegacyFunctions()
        End If
    End Sub

    ''' <summary>catch the save event, used to remove contents of DBListfunction results (data safety/space consumption)
    ''' choosing functions for removal of target data is done with custom docproperties</summary>
    ''' <param name="Wb"></param>
    ''' <param name="SaveAsUI"></param>
    ''' <param name="Cancel"></param>
    Private Sub App_WorkbookBeforeSave(ByVal Wb As Workbook, SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeSave
        Dim refreshDBFuncs As Boolean
        Dim docproperty
        Dim DBFCContentColl As Collection, DBFCAllColl As Collection
        Dim theFunc
        Dim ws As Worksheet, lastWs As Worksheet = Nothing
        Dim searchCell As Range
        Dim firstAddress As String

        On Error GoTo App_WorkbookBeforeSave_Err
        DBFCContentColl = New Collection
        DBFCAllColl = New Collection
        refreshDBFuncs = True
        For Each docproperty In Wb.CustomDocumentProperties
            If TypeName(docproperty.Value) = "Boolean" Then
                If Left$(docproperty.name, 5) = "DBFCC" And docproperty.Value Then DBFCContentColl.Add(True, Mid$(docproperty.name, 6))
                If Left$(docproperty.name, 5) = "DBFCA" And docproperty.Value Then DBFCAllColl.Add(True, Mid$(docproperty.name, 6))
                If docproperty.name = "DBFskip" Then refreshDBFuncs = Not docproperty.Value
            End If
        Next
        dontCalcWhileClearing = True
        For Each ws In Wb.Worksheets
            For Each theFunc In {"DBListFetch(", "DBRowFetch("}
                searchCell = ws.Cells.Find(What:=theFunc, After:=ws.Range("A1"), LookIn:=XlFindLookIn.xlFormulas, LookAt:=XlLookAt.xlPart, SearchOrder:=XlSearchOrder.xlByRows, SearchDirection:=XlSearchDirection.xlNext, MatchCase:=False)
                If Not (searchCell Is Nothing) Then
                    firstAddress = searchCell.Address
                    Do
                        Dim targetName As String
                        targetName = getDBRangeName(searchCell).Name
                        targetName = Replace(targetName, "DBFsource", "DBFtarget", 1, , vbTextCompare)
                        Dim DBFCC As Boolean : Dim DBFCA As Boolean
                        DBFCC = False : DBFCA = False
                        On Error Resume Next
                        DBFCC = DBFCContentColl("*")
                        DBFCC = DBFCContentColl(searchCell.Parent.name & "!" & Replace(searchCell.Address, "$", String.Empty)) Or DBFCC
                        DBFCA = DBFCAllColl("*")
                        DBFCA = DBFCAllColl(searchCell.Parent.name & "!" & Replace(searchCell.Address, "$", String.Empty)) Or DBFCA
                        Err.Clear()
                        Dim theTargetRange As Range
                        theTargetRange = hostApp.Range(targetName)
                        If DBFCC Then
                            theTargetRange.Parent.Range(theTargetRange.Parent.Cells(theTargetRange.Row, theTargetRange.Column), theTargetRange.Parent.Cells(theTargetRange.Row + theTargetRange.Rows.Count - 1, theTargetRange.Column + theTargetRange.Columns.Count - 1)).ClearContents
                            LogInfo("App_WorkbookSave/DBFCC cleared")
                        End If
                        If DBFCA Then
                            theTargetRange.Parent.Range(theTargetRange.Parent.Cells(theTargetRange.Row + 2, theTargetRange.Column), theTargetRange.Parent.Cells(theTargetRange.Row + theTargetRange.Rows.Count - 1, theTargetRange.Column + theTargetRange.Columns.Count - 1)).Clear
                            theTargetRange.Parent.Range(theTargetRange.Parent.Cells(theTargetRange.Row, theTargetRange.Column), theTargetRange.Parent.Cells(theTargetRange.Row + 2, theTargetRange.Column + theTargetRange.Columns.Count - 1)).ClearContents
                            LogInfo("App_WorkbookSave/DBFCA cleared")
                        End If
                        searchCell = ws.Cells.FindNext(searchCell)
                    Loop While Not searchCell Is Nothing And searchCell.Address <> firstAddress
                End If
            Next
            lastWs = ws
        Next
        dontCalcWhileClearing = False
        ' reset the cell find dialog....
        searchCell = Nothing
        searchCell = lastWs.Cells.Find(What:="", After:=lastWs.Range("A1"), LookIn:=XlFindLookIn.xlFormulas, LookAt:=XlLookAt.xlPart, SearchOrder:=XlSearchOrder.xlByRows, SearchDirection:=XlSearchDirection.xlNext, MatchCase:=False)
        lastWs = Nothing
        ' refresh after save event
        If refreshDBFuncs And (DBFCContentColl.Count > 0 Or DBFCAllColl.Count > 0) Then
            aTimer = New Timers.Timer(100)
            AddHandler aTimer.Elapsed, New ElapsedEventHandler(AddressOf refreshDBFuncLater)
            aTimer.Enabled = True
        End If
        Exit Sub

App_WorkbookBeforeSave_Err:
        WriteToLog("DBFuncEventHandler.App_WorkbookBeforeSave Error: " & Wb.Name & Err.Description & ", in line " & Erl(), EventLogEntryType.Error)
    End Sub

    ''' <summary>"OnTime" event function to "escape" workbook_save: event procedure to refetch DB functions results after saving</summary>
    ''' <param name="sender">the sending object (ourselves)</param>
    ''' <param name="e">Data for the Timer.Elapsed event</param>
    Shared Sub refreshDBFuncLater(ByVal sender As Object, ByVal e As ElapsedEventArgs)
        Dim previouslySaved As Boolean

        If Not hostApp.ActiveWorkbook Is Nothing Then
            previouslySaved = hostApp.ActiveWorkbook.Saved
            refreshDBFunctions(hostApp.ActiveWorkbook, True)
            hostApp.ActiveWorkbook.Saved = previouslySaved
        End If
    End Sub


    ''' <summary>catch the calculation event: this is the technical basis to separate actions not usually allowed in UDFs</summary>
    ''' <param name="Sh">the invoking Sheet</param>
    Private Sub App_SheetCalculate(ByVal Sh As Object) Handles Application.SheetCalculate
        Dim calcCont As ContainerCalcMsgs
        Dim statusCont As ContainerStatusMsgs
        Dim callID As String, callerText As String
        Dim xlcalcmode As Long

        If allCalcContainers Is Nothing Then Exit Sub
        If allStatusContainers Is Nothing Then allStatusContainers = New Collection

        'hostApp.StatusBar = "Number of calcContainers: " & allStatusContainers.Count
        For Each calcCont In allCalcContainers

            With calcCont
                On Error Resume Next
                Err.Clear()

                ' fetch each container just once for working on (use calcCont.working and removal...)
                ' do not compare Sh with .callsheet, as excel sometimes doesn't invoke the calcevent for a sheet
                If Not .caller Is Nothing Then
                    callID = calcCont.callID

                    Dim doFetching As Boolean = True
                    ' To avoid unneccessary queries (volatile funcions, autofilter set, etc.) , only run data fetching if ConnString/query is either not yet cached or has changed !
                    If Not existsQueryCache(callID) Then
                        queryCache.Add(calcCont.ConnString & calcCont.Query, callID)
                        doFetching = True
                    Else
                        doFetching = (calcCont.ConnString & calcCont.Query <> queryCache(callID))
                        ' refresh the query cache...
                        queryCache.Remove(callID)
                        queryCache.Add(calcCont.ConnString & calcCont.Query, callID)
                    End If

                    If doFetching Then
                        ' avoid (infinite loop) processing if the event procedure invoked the calling DB function again (indirectly by changing target cells)
                        If Not (allCalcContainers(callID).working Or allCalcContainers(callID).errOccured) Then ' either an error occured or working flag was not reset...
                            allCalcContainers(callID).working = True

                            'get status container or add new one to global collection of all status msg containers
                            If existsStatusColl(callID) Then
                                statusCont = allStatusContainers(callID)
                            Else
                                statusCont = New ContainerStatusMsgs
                                allStatusContainers.Add(statusCont, callID)
                            End If

                            callerText = .caller.Formula

                            If Err.Number <> 0 Then
                                WriteToLog("App_SheetCalculate: ERROR with retrieving .caller.Formula: " & Err.Description, EventLogEntryType.Error)
                                errorReason = "App_SheetCalculate: ERROR with .caller.Formula: " & Err.Description
                                allCalcContainers(callID).errOccured = True
                            End If

                            ' create/reconnect database connection, except for setting query/connstring with DBSETQUERY !
                            If InStr(1, UCase$(callerText), "DBSETQUERY(") = 0 Then
                                ODBCconnString = String.Empty

                                If InStr(1, UCase$(.ConnString), ";ODBC;") Then
                                    ODBCconnString = Mid$(.ConnString, InStr(1, UCase$(.ConnString), ";ODBC;") + 1)
                                    .ConnString = Left$(.ConnString, InStr(1, UCase$(.ConnString), ";ODBC;") - 1)
                                End If

                                If cnn Is Nothing Then cnn = New ADODB.Connection
                                If CurrConnString <> .ConnString And cnn.State <> 0 Then cnn.Close()

                                If cnn.State <> ADODB.ObjectStateEnum.adStateOpen And Not dontTryConnection Then
                                    cnn.ConnectionTimeout = CnnTimeout
                                    cnn.CommandTimeout = CmdTimeout
                                    cnn.CursorLocation = CursorLocationEnum.adUseClient
                                    hostApp.StatusBar = "Trying " & CnnTimeout & " sec. with connstring: " & .ConnString
                                    Err.Clear()
                                    cnn.Open(.ConnString)

                                    If Err.Number <> 0 Then
                                        WriteToLog("App_SheetCalculate Connection error: " & Err.Description, EventLogEntryType.Error)
                                        ' prevent multiple reconnecting if connection errors present...
                                        dontTryConnection = True
                                        LogError("Connection Error: " & Err.Description)
                                        errorReason = "Connection Error: " & Err.Description
                                        statusCont.statusMsg = errorReason
                                        allCalcContainers(callID).errOccured = True
                                    End If
                                    CurrConnString = .ConnString
                                End If
                            End If

                            ' Do the work !!
                            If cnn.State = 1 And Not allCalcContainers(callID).errOccured Or UCase$(Left$(.ConnString, 5)) = "ODBC;" Then    ' only try database functions for open connection  and no previous errors!!
                                xlcalcmode = hostApp.Calculation
                                hostApp.EnableEvents = False
                                hostApp.Cursor = XlMousePointer.xlWait  ' To show the hourglass
                                Interrupted = False

                                If InStr(1, UCase$(callerText), "DBLISTFETCH(") > 0 Then
                                    DBListQuery(calcCont, statusCont)
                                ElseIf InStr(1, UCase$(callerText), "DBROWFETCH(") > 0 Then
                                    DBRowQuery(calcCont, statusCont)
                                ElseIf InStr(1, UCase$(callerText), "DBSETQUERY(") > 0 Then
                                    DBSetQueryParams(calcCont, statusCont)
                                End If

                                ' Clean up settings...
                                hostApp.Cursor = XlMousePointer.xlDefault  ' To return cursor to normal
                                hostApp.StatusBar = False
                                'this is NOT done here, otherwise we have a problem with print preview !!
                                'Instead, enable events at the end of the calling db function
                                'hostApp.EnableEvents = True
                            Else
                                statusCont.statusMsg = "No open connection for DB function, reason: " & errorReason
                            End If

                            ' to work around silly Excel bug with Dirty Method (in refresh for selected target area)
                            ' we have to select the sheet with the "dirtied" cell. Here we return to the (calling) target
                            If Not origWS Is Nothing Then
                                origWS.Select()
                                origWS = Nothing
                                hostApp.ScreenUpdating = True
                            End If
                            hostApp.Calculation = xlcalcmode

                            ' in manual calculation no recalc of own results is done so we do this now:
                            If xlcalcmode = XlCalculation.xlCalculationManual Then calcCont.targetRange.Parent.Calculate
                        End If
                    Else
                        ' still set this (even for disabled fetching due to unchanged query) to true as it is needed in the calling function to determine a worked on calc container...
                        allCalcContainers(callID).working = True
                    End If
                End If
            End With
nextCalcCont:
        Next

        ' remove all worked containers
        For Each calcCont In allCalcContainers
            If calcCont.working Then
                allCalcContainers.Remove(calcCont.callID)
                calcCont = Nothing
            End If
        Next

        If allCalcContainers.Count = 0 Then
            allCalcContainers = Nothing
            allStatusContainers = Nothing
        End If
    End Sub

    ''' <summary>set Query parameters (query text and connection string) of Query List or pivot table (incl. chart)</summary>
    ''' <param name="calcCont"><see cref="ContainerCalcMsgs"/></param>
    ''' <param name="statusCont"><see cref="ContainerStatusMsgs"/></param>
    Public Sub DBSetQueryParams(calcCont As ContainerCalcMsgs, statusCont As ContainerStatusMsgs)
        Dim TargetCell As Range
        Dim targetSH As Worksheet
        Dim targetWB As Workbook
        Dim callID As String, Query As String, errMsg As String, ConnString As String
        Dim thePivotTable As PivotTable
        Dim theListObject As ListObject

        hostApp.Calculation = XlCalculation.xlCalculationManual
        ' this works around the data validation input bug
        ' when selecting a value from a list of validated field, excel won't react to
        ' Application.Calculation changes, so just leave here...
        If hostApp.Calculation <> XlCalculation.xlCalculationManual Then Exit Sub

        callID = calcCont.callID
        targetSH = calcCont.targetRange.Parent
        targetWB = calcCont.targetRange.Parent.Parent
        TargetCell = calcCont.targetRange
        Query = calcCont.Query
        ConnString = calcCont.ConnString

        On Error Resume Next
        thePivotTable = TargetCell.PivotTable
        theListObject = TargetCell.ListObject
        Err.Clear()

        Dim connType As String
        Dim bgQuery As Boolean
        On Error GoTo DBSetQueryParams_Error
        If Not thePivotTable Is Nothing Then
            bgQuery = thePivotTable.PivotCache.BackgroundQuery
            connType = Left$(thePivotTable.PivotCache.Connection, InStr(1, thePivotTable.PivotCache.Connection, ";"))
            thePivotTable.PivotCache.Connection = connType & ConnString
            thePivotTable.PivotCache.CommandType = XlCmdType.xlCmdSql
            thePivotTable.PivotCache.CommandText = Query
            thePivotTable.PivotCache.BackgroundQuery = False
            thePivotTable.PivotCache.Refresh()
            statusCont.statusMsg = "Set " & connType & " PivotTable to (bgQuery= " & bgQuery & "): " & Query
            thePivotTable.PivotCache.BackgroundQuery = bgQuery
        End If

        If Not theListObject Is Nothing Then
            bgQuery = theListObject.QueryTable.BackgroundQuery
            connType = Left$(theListObject.QueryTable.Connection, InStr(1, theListObject.QueryTable.Connection, ";"))
            ' Attention Dirty Hack ! This works only for SQLOLEDB driver to ODBC driver setting change...
            theListObject.QueryTable.Connection = connType & Replace(ConnString, "provider=SQLOLEDB", "driver=SQL SERVER")
            theListObject.QueryTable.CommandType = XlCmdType.xlCmdSql
            theListObject.QueryTable.CommandText = Query
            theListObject.QueryTable.BackgroundQuery = False
            theListObject.QueryTable.Refresh()
            statusCont.statusMsg = "Set " & connType & " ListObject to (bgQuery= " & bgQuery & "): " & Query
            theListObject.QueryTable.BackgroundQuery = bgQuery
        End If
        Exit Sub

DBSetQueryParams_Error:
        errMsg = Err.Description & " in query: " & Query
        WriteToLog("DBFuncEventHandler.DBSetQueryParams Error: " & errMsg & ", caller: " & callID, EventLogEntryType.Error)

        statusCont.statusMsg = errMsg
        ' need to mark calc container here as excel won't return to main event proc in case of error
        ' calc container is then removed in calling function
        allCalcContainers(callID).errOccured = True
        allCalcContainers(callID).callsheet.Range(allCalcContainers(callID).caller.Address).Dirty
    End Sub

    ''' <summary>Query list of data delimited by maxRows and maxCols, write it into targetCells
    '''             additionally copy formulas contained in formulaRange and extend list depending on extendArea</summary>
    ''' <param name="calcCont"><see cref="ContainerCalcMsgs"/></param>
    ''' <param name="statusCont"><see cref="ContainerStatusMsgs"/></param>
    Public Sub DBListQuery(calcCont As ContainerCalcMsgs, statusCont As ContainerStatusMsgs)
        Dim tableRst As ADODB.Recordset
        Dim targetCells As Range, formulaRange As Range, formulaFilledRange As Range = Nothing
        Dim targetSH As Worksheet, formulaSH As Worksheet = Nothing
        Dim copyFormat() As String = Nothing, copyFormatF() As String = Nothing
        Dim headingOffset As Long, rowDataStart As Long, startRow As Long, startCol As Long, arrayCols As Long, arrayRows As Long, copyDown As Long
        Dim oldRows As Long, oldCols As Long, oldFRows As Long, oldFCols As Long, retrievedRows As Long, targetColumns As Long, formulaStart As Long
        Dim callID As String, Query As String, warning As String, errMsg As String, targetRangeName As String, formulaRangeName As String, tmpname As String
        Dim extendArea As Integer, headingsPresent As Boolean, ShowRowNumbers As Boolean
        Dim storedNames() As String

        hostApp.Calculation = XlCalculation.xlCalculationManual
        ' this works around the data validation input bug
        ' when selecting a value from a list of validated field, excel won't react to
        ' Application.Calculation changes, so just leave here...
        If hostApp.Calculation <> XlCalculation.xlCalculationManual Then Exit Sub

        callID = calcCont.callID
        LogInfo("entering DBListQuery: callID:" + callID)
        formulaRange = calcCont.formulaRange
        targetSH = calcCont.targetRange.Parent
        targetCells = calcCont.targetRange
        Query = calcCont.Query
        extendArea = calcCont.extendArea
        ShowRowNumbers = calcCont.ShowRowNumbers
        headingsPresent = calcCont.HeaderInfo
        targetRangeName = calcCont.targetRangeName
        formulaRangeName = calcCont.formulaRangeName
        warning = String.Empty

        Dim srcExtentConnect As String, targetExtent As String, targetExtentF As String
        On Error Resume Next
        srcExtentConnect = calcCont.caller.Name.name
        If Err.Number <> 0 Or InStr(1, srcExtentConnect, "DBFsource") = 0 Then
            Err.Clear()
            srcExtentConnect = "DBFsource" & Replace(Replace(CDbl(Now.ToOADate()), ",", String.Empty), ".", String.Empty)
            calcCont.caller.Name = srcExtentConnect
            calcCont.callsheet.Parent.Names(srcExtentConnect).Visible = False
            If Err.Number <> 0 Then
                errMsg = "Error in setting srcExtentConnect name: " & Err.Description & " in query: " & Query
                GoTo err_0
            End If
        End If
        targetExtent = Replace(srcExtentConnect, "DBFsource", "DBFtarget")
        targetExtentF = Replace(srcExtentConnect, "DBFsource", "DBFtargetF")

        If Not formulaRange Is Nothing Then
            formulaSH = formulaRange.Parent
            ' only first row of formulaRange is important, rest will be autofilled down (actually this is needed to make the autoformat work)
            formulaRange = formulaRange.Rows(1)
        End If
        Err.Clear()

        startRow = targetCells.Cells.Row : startCol = targetCells.Cells.Column
        If Err.Number <> 0 Then
            errMsg = "Error in setting startRow/startCol: " & Err.Description & " in query: " & Query
            GoTo err_0
        End If

        oldRows = 0 : oldCols = 0 : oldFRows = 0 : oldFCols = 0
        On Error Resume Next
        oldRows = targetSH.Parent.Names(targetExtent).RefersToRange.Rows.Count
        oldCols = targetSH.Parent.Names(targetExtent).RefersToRange.Columns.Count
        If Err.Number = 0 Then
            Err.Clear()
            ' clear old data area
            targetSH.Parent.Names(targetExtent).RefersToRange.ClearContents
            If Err.Number <> 0 Then
                errMsg = "Error in clearing old data for targetExtent: (" & Err.Description & ") in query: " & Query
                GoTo err_0
            End If
        End If
        Err.Clear()

        On Error Resume Next
        oldFRows = formulaSH.Parent.Names(targetExtentF).RefersToRange.Rows.Count
        oldFCols = formulaSH.Parent.Names(targetExtentF).RefersToRange.Columns.Count
        If Err.Number = 0 And oldFRows > 2 Then
            Err.Clear()
            ' clear old formulas
            formulaSH.Range(formulaSH.Cells(formulaRange.Row + 1, formulaRange.Column), formulaSH.Cells(formulaRange.Row + oldFRows - 1, formulaRange.Column + oldFCols - 1)).ClearContents()

            If Err.Number <> 0 Then
                errMsg = "Error in clearing old data for formulaSH: (" & Err.Description & ") in query: " & Query
                GoTo err_0
            End If
        End If
        Err.Clear()

        hostApp.StatusBar = "Retrieving data for DBList: " & IIf(targetRangeName.Length > 0, targetRangeName, targetSH.Name & "!" & targetCells.Address)
        tableRst = New ADODB.Recordset
        tableRst.Open(Query, cnn, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly, CommandTypeEnum.adCmdText)
        Dim dberr As String = String.Empty
        If cnn.Errors.Count > 0 Then
            Dim errcount As Integer
            For errcount = 0 To cnn.Errors.Count - 1
                If cnn.Errors.Item(errcount).Description <> Err.Description Then _
                   dberr = dberr & ";" & cnn.Errors.Item(errcount).Description
            Next
            If dberr.Length > 0 Then dberr = " (" & dberr & ")"
        End If
        If Err.Number <> 0 Then
            errMsg = "Error in retrieving data: " & Err.Description & dberr & " in query: " & Query
            GoTo err_1
        End If
        ' this fails in case of known issue with OLEDB driver...
        retrievedRows = tableRst.RecordCount
        If Err.Number <> 0 Then
            errMsg = "Error in opening recordset: " & Err.Description & dberr & " in query: " & Query
            GoTo err_1
        End If

        ' from now on we don't propagate any errors as data is modified in sheet....
        hostApp.StatusBar = "Displaying data for DBList: " & IIf(targetRangeName.Length > 0, targetRangeName, targetSH.Name & "!" & targetCells.Address)
        If tableRst.EOF Then warning = "Warning: No Data returned in query: " & Query
        ' set size for named range (size: arrayRows, arrayCols) used for resizing the data area (old extent)
        arrayCols = tableRst.Fields.Count
        arrayRows = retrievedRows
        ' need to shift down 1 row if headings are present
        arrayRows += IIf(headingsPresent, 1, 0)
        rowDataStart = 1 + IIf(headingsPresent, 1, 0)

        ' check whether retrieved data exceeds excel's limits and limit output (arrayRows/arrayCols) in case ...
        ' check rows
        If targetCells.Row + arrayRows > (targetCells.EntireColumn.Rows.Count + 1) Then
            warning = "row count" & " of returned data exceeds max row of excel: start row:" & targetCells.Row & " + row count:" & arrayRows & " > max row+1:" & targetCells.EntireColumn.Rows.Count + 1
            arrayRows = targetCells.EntireColumn.Rows.Count - targetCells.Row + 1
        End If
        ' check columns
        If targetCells.Column + arrayCols > (targetCells.EntireRow.Columns.Count + 1) Then
            warning = warning & ", column count" & " of returned data exceed max column of excel: start column:" & targetCells.Column & " + column count:" & arrayCols & " > max column+1:" & targetCells.EntireRow.Columns.Count + 1
            arrayCols = targetCells.EntireRow.Columns.Count - targetCells.Column + 1
        End If

        ' autoformat: copy 1st rows formats range to reinsert them afterwards
        targetColumns = arrayCols - IIf(ShowRowNumbers, 0, 1)
        If calcCont.autoformat Then
            arrayRows += IIf(headingsPresent And arrayRows = 1, 1, 0)  ' need special case for autoformat
            Dim i As Long
            For i = 0 To targetColumns
                ReDim Preserve copyFormat(i)
                copyFormat(i) = targetSH.Cells(targetCells.Row + rowDataStart - 1, targetCells.Column + i).NumberFormat
            Next
            ' now for the calculated data area
            If Not formulaRange Is Nothing Then
                For i = 0 To formulaRange.Columns.Count - 1
                    ReDim Preserve copyFormatF(i)
                    copyFormatF(i) = formulaSH.Cells(targetCells.Row + rowDataStart - 1, formulaRange.Column + i).NumberFormat
                Next
            End If
        End If
        If arrayRows = 0 Then arrayRows = 1  ' sane behavior of named range in case no data retrieved...

        ' check if formulaRange and targetRange overlap !
        Dim possibleIntersection As Range = Nothing
        possibleIntersection = hostApp.Intersect(formulaRange, targetSH.Range(targetCells.Cells(1, 1), targetCells.Cells(1, 1).Offset(0, arrayCols - 1)))
        Err.Clear()
        If Not possibleIntersection Is Nothing Then
            warning = warning & ", formulaRange and targetRange intersect (" & targetSH.Name & "!" & possibleIntersection.Address & "), formula copying disabled !!"
            formulaRange = Nothing
        End If

        '''' data list and formula range extension (ignored in first call after creation -> no defined name is set -> oldRows=0)...
        Dim headingFirstRowPrevent As Long, headingLastRowPrevent As Long
        headingOffset = IIf(headingsPresent, 1, 0)  ' use that for generally regarding headings !!
        If oldRows > 0 Then
            ' either cells/rows are shifted down (old data area was smaller than current) ...
            If oldRows < arrayRows Then
                'prevent insertion from heading row if headings are present (to not get the header formats..)
                headingFirstRowPrevent = IIf(headingsPresent And oldRows = 1 And arrayRows > 2, 1, 0)
                '1: add cells (not whole rows)
                If extendArea = 1 Then
                    targetSH.Range(targetSH.Cells(startRow + oldRows + headingFirstRowPrevent, startCol), targetSH.Cells(startRow + arrayRows - 1, startCol + oldCols - 1)).Insert(Shift:=XlDirection.xlDown)
                    If Not formulaRange Is Nothing Then formulaSH.Range(formulaSH.Cells(startRow + oldFRows + headingOffset, formulaRange.Column), formulaSH.Cells(startRow + arrayRows - 1 - headingFirstRowPrevent, formulaRange.Column + oldFCols - 1)).Insert(Shift:=XlDirection.xlDown)
                    '2: add whole rows
                ElseIf extendArea = 2 Then
                    targetSH.Rows(startRow + oldRows + headingFirstRowPrevent & ":" & startRow + arrayRows - 1).Insert(Shift:=XlDirection.xlDown)
                    If Not formulaRange Is Nothing Then
                        ' take care not to insert twice (if we're having formulas in the same sheet)
                        If Not targetSH Is formulaSH Then formulaSH.Rows(startRow + oldFRows + headingOffset & ":" & startRow + arrayRows - 1 - headingFirstRowPrevent).Insert(Shift:=XlDirection.xlDown)
                    End If
                End If
                'else 0: just overwrite -> no special action

                ' .... or cells/rows are shifted up (old data area was larger than current)
            ElseIf oldRows > arrayRows Then
                'prevent deletion of last row if headings are present (to not get the header formats, lose formulas, etc..)
                headingLastRowPrevent = IIf(headingsPresent And arrayRows = 1 And oldRows > 2, 1, 0)
                '1: add cells (not whole rows)
                If extendArea = 1 Then
                    targetSH.Range(targetSH.Cells(startRow + arrayRows + headingLastRowPrevent, startCol), targetSH.Cells(startRow + oldRows - 1, startCol + oldCols - 1)).Delete(Shift:=XlDirection.xlUp)
                    If Not formulaRange Is Nothing Then formulaSH.Range(formulaSH.Cells(startRow + arrayRows + headingLastRowPrevent, formulaRange.Column), formulaSH.Cells(startRow + oldFRows - 1 + headingOffset, formulaRange.Column + oldFCols - 1)).Delete(Shift:=XlDirection.xlUp)
                    '2: add whole rows
                ElseIf extendArea = 2 Then
                    targetSH.Rows(startRow + arrayRows + headingLastRowPrevent & ":" & startRow + oldRows - 1).Delete(Shift:=XlDirection.xlUp)
                    If Not formulaRange Is Nothing Then
                        ' take care not to delete twice (if we're having formulas in the same sheet)
                        If Not targetSH Is formulaSH Then formulaSH.Rows(startRow + arrayRows + headingLastRowPrevent & ":" & startRow + oldFRows - 1 + headingOffset).Delete(Shift:=XlDirection.xlUp)
                    End If
                End If
                '0: just overwrite -> no special action
            End If
            If Err.Number <> 0 Then
                errMsg = "Error in resizing area: " & Err.Description & " in query: " & Query
                GoTo err_1
            End If
        End If
        ' now fill in the data from the query
        If ODBCconnString.Length > 0 Then
            With targetSH.QueryTables.Add(Connection:=ODBCconnString, Destination:=targetCells)
                .CommandText = Query
                .FieldNames = headingsPresent
                .RowNumbers = ShowRowNumbers
                .PreserveFormatting = True
                .AdjustColumnWidth = False
                .FillAdjacentFormulas = False
                .BackgroundQuery = True
                .RefreshStyle = XlCellInsertionMode.xlOverwriteCells
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .Refresh()
                tmpname = .Name
                .Delete()
            End With
        Else
            With targetSH.QueryTables.Add(Connection:=tableRst, Destination:=targetCells)
                .FieldNames = headingsPresent
                .RowNumbers = ShowRowNumbers
                .PreserveFormatting = True
                .AdjustColumnWidth = False
                .FillAdjacentFormulas = False
                .BackgroundQuery = True
                .RefreshStyle = XlCellInsertionMode.xlOverwriteCells   ' this is required to prevent "right" shifting of cells at the beginning !
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .Refresh()
                tmpname = .Name
                .Delete()
            End With
        End If
        If Err.Number <> 0 Then
            errMsg = "Error in adding QueryTable: " & Err.Description & " in query: " & Query
            GoTo err_2
        End If
        tableRst.Close()

        ' sometimes excel doesn't delete the querytable given name
        targetSH.Names(tmpname).Delete
        targetSH.Parent.Names(tmpname).Delete
        Err.Clear()

        '''' formulas recreation (removal and autofill new ones)
        If Not formulaRange Is Nothing Then
            formulaSH = formulaRange.Parent
            With formulaRange
                If .Row < startRow + rowDataStart - 1 Then
                    warning = "Error: formulaRange start above data-area, no formulas filled down !"
                Else
                    ' retrieve bottom of formula range
                    ' check for excels boundaries !!
                    If .Cells.Row + arrayRows > .EntireColumn.Rows.Count + 1 Then
                        warning = warning & ", formulas would exceed max row of excel: start row:" & formulaStart & " + row count:" & arrayRows & " > max row+1:" & .EntireColumn.Rows.Count + 1
                        copyDown = .EntireColumn.Rows.Count
                    Else
                        'the normal end of our autofilled rows = formula start + list size,
                        'reduced by offset of formula start and startRow if formulas start below data area top
                        copyDown = .Cells.Row + arrayRows - 1 - IIf(.Cells.Row > startRow, .Cells.Row - startRow, 0)
                    End If
                    ' sanity check not to fill upwards !
                    If copyDown > .Cells.Row Then .Cells.AutoFill(Destination:=formulaSH.Range(.Cells, formulaSH.Cells(copyDown, .Column + .Columns.Count - 1)))
                    ' restore filters in formulaSheet, calculate explicitly or we wouldn't filter correctly !
                    formulaFilledRange = formulaSH.Range(formulaSH.Cells(.Row, .Column), formulaSH.Cells(copyDown, .Column + .Columns.Count - 1))
                    formulaFilledRange.Calculate()

                    ' reassign internal name to changed formula area
                    ' delete the name to have a "clean" name area (otherwise visible = false setting wont work for dataTargetRange)
                    storedNames = removeRangeName(formulaFilledRange, targetExtentF)
                    formulaFilledRange.Name = targetExtentF
                    formulaFilledRange.Name.Visible = False
                    restoreRangeNames(formulaFilledRange, storedNames)
                    ' reassign visible defined name to changed formula area only if defined...
                    If formulaRangeName.Length > 0 Then
                        formulaFilledRange.Name = formulaRangeName    ' DO NOT use formulaFilledRange.Name.Visible = True, or hidden range will also be visible...
                    End If
                    If Err.Number <> 0 Then
                        errMsg = "Error in (re)assigning formula range name: " & Err.Description & " in query: " & Query
                        GoTo err_0
                    End If
                End If
            End With
        End If
        If Err.Number <> 0 Then
            errMsg = "Error in filling formulas: " & Err.Description & " in query: " & Query
            GoTo err_0
        End If

        Dim newTargetRange As Range
        ' reassign name to changed data area
        If targetRangeName.Length > 0 Then
            ' if formulas are adjacent to data extend name to formula range !
            Dim additionalFormulaColumns As Long : additionalFormulaColumns = 0
            ' need this as excel throws errors when comparing nonexistent formulaSH !!
            If Not formulaRange Is Nothing Then
                If targetSH Is formulaSH And formulaRange.Column = startCol + targetColumns + 1 Then additionalFormulaColumns = formulaRange.Columns.Count
            End If
            ' set the new hidden targetExtent name...
            newTargetRange = targetSH.Range(targetSH.Cells(startRow, startCol), targetSH.Cells(startRow + arrayRows - 1, startCol + targetColumns))
            storedNames = removeRangeName(newTargetRange, targetExtent)
            newTargetRange.Name = targetExtent
            newTargetRange.Name.Visible = False
            restoreRangeNames(newTargetRange, storedNames)
            ' now set the name for the total area
            targetSH.Range(targetSH.Cells(startRow, startCol), targetSH.Cells(startRow + arrayRows - 1, startCol + targetColumns + additionalFormulaColumns)).Name = targetRangeName
        Else
            ' set the new hidden targetExtent name...
            newTargetRange = targetSH.Range(targetSH.Cells(startRow, startCol), targetSH.Cells(startRow + arrayRows - 1, startCol + targetColumns))
            storedNames = removeRangeName(newTargetRange, targetExtent)
            newTargetRange.Name = targetExtent
            newTargetRange.Name.Visible = False
            restoreRangeNames(newTargetRange, storedNames)
        End If
        If Err.Number <> 0 Then
            errMsg = "Error in (re)assigning data target name: " & Err.Description & " (maybe known issue with 'cell like' sheetnames, e.g. 'C701 country' ?) in query: " & Query
            GoTo err_0
        End If

        '''' any warnings, errors ?
        If warning.Length > 0 Then
            If InStr(1, warning, "Error:") = 0 And InStr(1, warning, "No Data") = 0 Then
                If Left$(warning, 1) = "," Then
                    warning = Right$(warning, Len(warning) - 2)
                End If
                statusCont.statusMsg = "Retrieved " & retrievedRows & " record" & IIf(retrievedRows > 1, "s", String.Empty) & ", Warning: " & warning
            Else
                statusCont.statusMsg = warning
            End If
        Else
            statusCont.statusMsg = "Retrieved " & retrievedRows & " record" & IIf(retrievedRows > 1, "s", String.Empty) & " from: " & Query
        End If

        ' autoformat: restore format of 1st row...
        If calcCont.autoformat Then
            For i = 0 To UBound(copyFormat)
                newTargetRange.Rows(rowDataStart).Cells(i + 1).NumberFormat = copyFormat(i)
            Next
            ' now for the calculated cells...
            If Not formulaRange Is Nothing Then
                For i = 0 To UBound(copyFormatF)
                    formulaSH.Cells(targetCells.Row + rowDataStart - 1, formulaRange.Column + i).NumberFormat = copyFormatF(i)
                Next
            End If
            'auto format 1st rows down...
            If arrayRows > rowDataStart Then
                'This doesn't work anymore:
                'newTargetRange.Rows(rowDataStart).AutoFill(Destination:=newTargetRange.Rows(rowDataStart & ":" & arrayRows), Type:=XlAutoFillType.xlFillFormats)
                targetSH.Range(targetSH.Cells(targetCells.Row + rowDataStart - 1, newTargetRange.Column), targetSH.Cells(targetCells.Row + rowDataStart - 1, newTargetRange.Column + newTargetRange.Columns.Count - 1)).AutoFill(Destination:=targetSH.Range(targetSH.Cells(targetCells.Row + rowDataStart - 1, newTargetRange.Column), targetSH.Cells(targetCells.Row + arrayRows - 1, newTargetRange.Column + newTargetRange.Columns.Count - 1)), Type:=XlAutoFillType.xlFillFormats)
                If Not formulaRange Is Nothing Then
                    formulaSH.Range(formulaSH.Cells(targetCells.Row + rowDataStart - 1, formulaRange.Column), formulaSH.Cells(targetCells.Row + rowDataStart - 1, formulaRange.Column + formulaRange.Columns.Count - 1)).AutoFill(Destination:=formulaSH.Range(formulaSH.Cells(targetCells.Row + rowDataStart - 1, formulaRange.Column), formulaSH.Cells(targetCells.Row + arrayRows - 1, formulaRange.Column + formulaRange.Columns.Count - 1)), Type:=XlAutoFillType.xlFillFormats)
                End If
            End If
        End If

        If Err.Number <> 0 Then
            errMsg = "Error in restoring formats: " & Err.Description & " in query: " & Query
            GoTo err_0
        End If

        'auto fit columns AFTER autoformat so we don't have problems with applied formats visibility ...
        If calcCont.AutoFit Then
            newTargetRange.Columns.AutoFit()
            newTargetRange.Rows.AutoFit()

            If Not formulaRange Is Nothing And Not formulaFilledRange Is ExcelEmpty.Value Then
                If Not formulaFilledRange Is Nothing Then
                    formulaFilledRange.Columns.AutoFit()
                    formulaFilledRange.Rows.AutoFit()
                End If
            End If
        End If
        If Err.Number <> 0 Then
            errMsg = "Error in autofitting: " & Err.Description & " in query: " & Query
            GoTo err_0
        End If
        LogInfo("leaving DBListQuery: callID:" + callID)
        Exit Sub

err_2:
        targetSH.Names(tmpname).Delete
        targetSH.Parent.Names(tmpname).Delete
err_1:
        If tableRst.State <> 0 Then tableRst.Close()
err_0:
        Dim severity As EventLogEntryType
        If errMsg.Length = 0 Then
            errMsg = Err.Description & " in query: " & Query
            severity = EventLogEntryType.Error
        End If
        Err.Clear() ' this is important as otherwise the error propagates to App_SheetCalculate,
        ' which recalcs in case of errors there, leading to endless calc loops !!
        If severity = Nothing Then severity = EventLogEntryType.Warning
        WriteToLog("DBFuncEventHandler.DBListQuery Error: " & errMsg & ", caller: " & callID, severity)

        statusCont.statusMsg = errMsg
        ' need to mark calc container here as excel won't return to main event proc in case of error
        ' calc container is then removed in calling function
        allCalcContainers(callID).errOccured = True
    End Sub

    ''' <summary>Query (assumed) one row of data, write it into targetCells</summary>
    ''' <param name="calcCont"><see cref="ContainerCalcMsgs"/></param>
    ''' <param name="statusCont"><see cref="ContainerStatusMsgs"/></param>
    Public Sub DBRowQuery(calcCont As ContainerCalcMsgs, statusCont As ContainerStatusMsgs)
        Dim tableRst As ADODB.Recordset = Nothing
        Dim targetCells As Object
        Dim Query As String, callID As String, errMsg As String, refCollector As Range
        Dim headingsPresent As Boolean, headerFilled As Boolean, Delete As Boolean, fillByRows As Boolean
        Dim returnedRows As Long, fieldIter As Long, rangeIter As Long
        Dim theCell As Range, targetSlice As Range, targetSlices As Range
        Dim targetSH As Worksheet

        hostApp.Calculation = XlCalculation.xlCalculationManual
        ' this works around the data validation input bug
        ' when selecting a value from a list of validated field, excel won't react to
        ' Application.Calculation changes, so just leave here...
        If hostApp.Calculation <> XlCalculation.xlCalculationManual Then Exit Sub

        Query = calcCont.Query
        targetCells = calcCont.targetArray
        headingsPresent = calcCont.HeaderInfo
        callID = calcCont.callID
        targetSH = targetCells(0).Parent
        statusCont.statusMsg = ""
        On Error GoTo err_1
        hostApp.StatusBar = "Retrieving data for DBRows: " & targetSH.Name & "!" & targetCells(0).Address

        Dim srcExtentConnect As String, targetExtent As String
        On Error Resume Next
        srcExtentConnect = calcCont.caller.Name.name
        If Err.Number <> 0 Or InStr(1, UCase$(srcExtentConnect), "DBFSOURCE") = 0 Then
            Err.Clear()
            srcExtentConnect = "DBFsource" & Replace(Replace(CDbl(Now().ToOADate()), ",", String.Empty), ".", String.Empty)
            calcCont.caller.Name = srcExtentConnect
            calcCont.callsheet.Parent.Names(srcExtentConnect).Visible = False
            If Err.Number <> 0 Then
                errMsg = "Error in setting srcExtentConnect name: " & Err.Description & " in query: " & Query
                GoTo err_1
            End If
        End If
        targetExtent = Replace(srcExtentConnect, "DBFsource", "DBFtarget")
        ' remove old data in case we changed the target range array
        targetSH.Range(targetExtent).ClearContents()

        tableRst = New ADODB.Recordset
        tableRst.Open(Query, cnn, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly, CommandTypeEnum.adCmdText)
        On Error Resume Next
        Dim dberr As String = String.Empty
        If cnn.Errors.Count > 0 Then
            Dim errcount As Integer
            For errcount = 0 To cnn.Errors.Count - 1
                If cnn.Errors.Item(errcount).Description <> Err.Description Then _
                   dberr = dberr & ";" & cnn.Errors.Item(errcount).Description
            Next
            errMsg = "Error in retrieving data: " & dberr & " in query: " & Query
            GoTo err_1
        End If

        ' this fails in case of known issue with OLEDB driver...
        returnedRows = tableRst.RecordCount
        If Err.Number <> 0 Then
            errMsg = "Error in opening recordset: " & Err.Description & " in query: " & Query
            GoTo err_1
        End If

        On Error GoTo err_1

        ' check whether anything retrieved?
        Delete = tableRst.EOF
        If Delete Then statusCont.statusMsg = "Warning: No Data returned in query: " & Query

        ' if "heading range" is present then orientation of first range (header) defines layout of data: if "heading range" is column then data is returned columnwise, else row by row.
        ' if there is just one block of data then it is assumed that there are usually more rows than columns and orientation is set by row/column size
        fillByRows = IIf(UBound(targetCells) > 0, targetCells(0).Rows.Count < targetCells(0).Columns.Count, targetCells(0).Rows.Count > targetCells(0).Columns.Count)
        ' put values (single record) from Recordset into targetCells
        fieldIter = 0 : rangeIter = 0 : headerFilled = Not headingsPresent    ' if we don't need headers the assume they are filled already....
        refCollector = targetCells(0)
        Do
            If fillByRows Then
                targetSlices = targetCells(rangeIter).Rows
            Else
                targetSlices = targetCells(rangeIter).Columns
            End If
            For Each targetSlice In targetSlices
                If Interrupted Then
                    errMsg = "data fetching interrupted by user !"
                    GoTo err_1
                End If
                For Each theCell In targetSlice.Cells
                    If tableRst.EOF Then
                        theCell.Value = String.Empty
                    Else
                        If Not headerFilled Then
                            theCell.Value = tableRst.Fields(fieldIter).Name
                        ElseIf Delete Then
                            theCell.Value = String.Empty
                        Else
                            On Error Resume Next
                            theCell.Value = tableRst.Fields(fieldIter).Value
                            If Err.Number <> 0 Then theCell.Value = "Field '" & tableRst.Fields(fieldIter).Name & "' caused following error: '" & Err.Description & "'"
                            On Error GoTo err_1
                        End If
                        If fieldIter = tableRst.Fields.Count - 1 Then
                            If headerFilled Then
                                hostApp.StatusBar = "Displaying data for DBRows: " & targetSH.Name & "!" & targetCells(0).Address & ", record " & tableRst.AbsolutePosition & "/" & returnedRows
                                tableRst.MoveNext()
                            Else
                                headerFilled = True
                            End If
                            fieldIter = -1
                        End If
                    End If
                    fieldIter += 1
                Next
            Next
            rangeIter += 1
            If Not rangeIter > UBound(targetCells) Then refCollector = hostApp.Union(refCollector, targetCells(rangeIter))
        Loop Until rangeIter > UBound(targetCells)

        ' delete the name to have a "clean" name area (otherwise visible = false setting wont work for dataTargetRange)
        Dim storedNames() As String
        storedNames = removeRangeName(refCollector, targetExtent)
        refCollector.Name = targetExtent
        refCollector.Name.Visible = False
        restoreRangeNames(refCollector, storedNames)

        tableRst.Close()
        If statusCont.statusMsg.Length = 0 Then statusCont.statusMsg = "Retrieved " & returnedRows & " record" & IIf(returnedRows > 1, "s", String.Empty) & " from: " & Query
        Exit Sub

err_1:
        Dim severity As EventLogEntryType
        If errMsg.Length = 0 Then
            errMsg = Err.Description & " in query: " & Query
            severity = EventLogEntryType.Error
        End If
        'Err.Clear ' this is important as otherwise the error propagates to App_SheetCalculate,
        ' which recalcs in case of errors there, leading to endless calc loops !!
        If severity = Nothing Then severity = EventLogEntryType.Warning
        If tableRst.State <> 0 Then tableRst.Close()
        WriteToLog("DBFuncEventHandler.DBRowQuery Error: " & errMsg & ", caller: " & callID & ", in line " & Erl(), severity)
        statusCont.statusMsg = errMsg
        ' need to mark calc container here as excel won't return to main event proc in case of error
        ' calc container is then removed in calling function
        allCalcContainers(callID).errOccured = True
    End Sub

    ''' <summary>check whether a statusMsgContainer exists in allStatusContainers or not</summary>
    ''' <param name="theName">name of statusMsgContainer</param>
    ''' <returns>exists in allStatusContainers or not</returns>
    Private Function existsStatusColl(ByRef theName As String) As Boolean
        Dim dummy As String

        On Error GoTo err_1
        existsStatusColl = True
        dummy = allStatusContainers(theName).statusMsg
        Exit Function
err_1:
        Err.Clear()
        existsStatusColl = False
    End Function

    ''' <summary>check whether a dbfunction query for callID exists in queryCache or not</summary>
    ''' <param name="callID">callID of dbfunction in queryCache</param>
    ''' <returns>exists in queryCache or not</returns>
    Private Function existsQueryCache(ByRef callID As String) As Boolean
        Dim dummy As String

        On Error GoTo err_1
        existsQueryCache = True
        dummy = queryCache(callID)
        Exit Function
err_1:
        Err.Clear()
        existsQueryCache = False
    End Function

    ''' <summary>remove alle names from Range Target except the passed name (theName) and store them into list storedNames</summary>
    ''' <param name="Target"></param>
    ''' <param name="theName"></param>
    ''' <returns>the removed names as a string list for restoring them later (see restoreRangeNames)</returns>
    Private Function removeRangeName(Target As Range, theName As String) As String()
        Dim storedNames() As String = {}
        Dim i As Long
        Dim nextName As String

        i = 0
        On Error Resume Next
        nextName = Target.Name.name
        Do
            If Err.Number = 0 And nextName <> theName Then
                ReDim Preserve storedNames(i)
                storedNames(i) = nextName
                i += 1
            End If
            Target.Name.Delete
            nextName = Target.Name.name
        Loop Until Err.Number <> 0
        Err.Clear()
        removeRangeName = storedNames
    End Function

    ''' <summary>restore the passed storedNames into Range Target</summary>
    ''' <param name="Target"></param>
    ''' <param name="storedNames"></param>
    Private Sub restoreRangeNames(Target As Range, storedNames() As String)
        Dim theName
        If UBound(storedNames) > 0 Then
            For Each theName In storedNames
                If theName.Length > 0 Then Target.Name = theName
            Next
        End If
    End Sub

End Class