Imports Microsoft.Office.Interop.Excel
Imports ExcelDna.Integration
Imports ADODB

''' <summary>Provides a data structure for transporting information back from the calculation action procedure to the calling function</summary>
Public Class ContainedStatusMsg
    ''' <summary>any status msg used for displaying in the result of function</summary>
    Public statusMsg As String
End Class

''' <summary>Contains the public callable DB functions and helper functions</summary>
Public Module Functions
    ''' <summary>cnn object always the same (only open/close)</summary>
    Public conn As ADODB.Connection
    ''' <summary>connection string can be changed for calls with different connection strings</summary>
    Public CurrConnString As String
    ''' <summary>query cache for avoiding unnecessary recalculations/data retrievals by volatile inputs to DB Functions (now(), etc.)</summary>
    Public queryCache As Collection = New Collection
    ''' <summary>prevent multiple connection retries for each function in case of error</summary>
    Public dontTryConnection As Boolean
    ''' <summary>avoid entering dblistfetch function during clearing of listfetch areas (before saving)</summary>
    Public dontCalcWhileClearing As Boolean
    ' Global objects/variables for DBFuncs
    ''' <summary>store target filter in case of empty data lists</summary>
    Public targetFilterCont As Collection
    ''' <summary>because of a stupid excel behaviour (Range.Dirty only works if the parent sheet of Range is active)
    ''' we have to jump to the sheet containing the dbfunction and then activate back...</summary>
    Public origWS As Worksheet

    ''' <summary>global collection of information transport containers between function and calc event procedure</summary>
    Public StatusCollection As Collection = New Collection

    ''' <summary>Create database compliant date, time or datetime string from excel datetype value</summary>
    ''' <param name="datVal">date/time/datetime</param>
    ''' <param name="formatting">see remarks</param>
    ''' <returns>the DB compliant formatted date/time/datetime</returns>
    ''' <remarks>
    ''' formatting = 0: A simple datestring (format 'YYYYMMDD'), datetime values are converted to 'YYYYMMDD HH:MM:SS' and time values are converted to 'HH:MM:SS'.
    ''' formatting = 1: An ANSI compliant Date string (format date 'YYYY-MM-DD'), datetime values are converted to timestamp 'YYYY-MM-DD HH:MM:SS' and time values are converted to time time 'HH:MM:SS'.
    ''' formatting = 2: An ODBC compliant Date string (format {d 'YYYY-MM-DD'}), datetime values are converted to {ts 'YYYY-MM-DD HH:MM:SS'} and time values are converted to {t 'HH:MM:SS'}.
    ''' formatting = 3: An Access/JetDB compliant Date string (format #YYYY-MM-DD#), datetime values are converted to #YYYY-MM-DD HH:MM:SS# and time values are converted to #HH:MM:SS#.
    ''' formatting = 99 (default value): take the formatting option from setting DefaultDBDateFormatting (0 if not given)
    ''' </remarks>
    <ExcelFunction(Description:="Create database compliant date, time or datetime string from excel datetype value")>
    Public Function DBDate(<ExcelArgument(Description:="date/time/datetime")> ByVal datVal As Date,
                           <ExcelArgument(Description:="formatting option, 0: simple datestring (format 'YYYYMMDD'), 1: ANSI compliant Date string (format date 'YYYY-MM-DD'), 2: ODBC compliant Date string (format {d 'YYYY-MM-DD'}),3: Access/JetDB compliant Date string (format #DD/MM/YYYY#)")> Optional formatting As Integer = 99) As String
        Try
            Dim retval As String = String.Empty
            If formatting = 99 Then formatting = DefaultDBDateFormatting
            If Int(datVal.ToOADate()) = datVal.ToOADate() Then
                If formatting = 0 Then
                    retval = "'" & Format$(datVal, "yyyyMMdd") & "'"
                ElseIf formatting = 1 Then
                    retval = "DATE '" & Format$(datVal, "yyyy-MM-dd") & "'"
                ElseIf formatting = 2 Then
                    retval = "{d '" & Format$(datVal, "yyyy-MM-dd") & "'}"
                ElseIf formatting = 3 Then
                    retval = "#" & Format$(datVal, "yyyy-MM-dd") & "#"
                End If
            ElseIf CInt(datVal.ToOADate()) > 1 Then
                If formatting = 0 Then
                    retval = "'" & Format$(datVal, "yyyyMMdd hh:mm:ss") & "'"
                ElseIf formatting = 1 Then
                    retval = "timestamp '" & Format$(datVal, "yyyy-MM-dd hh:mm:ss") & "'"
                ElseIf formatting = 2 Then
                    retval = "{ts '" & Format$(datVal, "yyyy-MM-dd hh:mm:ss") & "'}"
                ElseIf formatting = 3 Then
                    retval = "#" & Format$(datVal, "yyyy-MM-dd hh:mm:ss") & "#"
                End If
            Else
                If formatting = 0 Then
                    retval = "'" & Format$(datVal, "hh:mm:ss") & "'"
                ElseIf formatting = 1 Then
                    retval = "time '" & Format$(datVal, "hh:mm:ss") & "'"
                ElseIf formatting = 2 Then
                    retval = "{t '" & Format$(datVal, "hh:mm:ss") & "'}"
                ElseIf formatting = 3 Then
                    retval = "#" & Format$(datVal, "hh:mm:ss") & "#"
                End If
            End If
            DBDate = retval
        Catch ex As Exception
            WriteToLog("Error (" & ex.Message & ") in Functions.DBDate", EventLogEntryType.Warning)
            DBDate = "Error (" & ex.Message & ") in function DBDate"
        End Try
    End Function

    ''' <summary>Create a database compliant string from cell values, potentially concatenating with other parts for easy inclusion of wildcards (%,_)</summary>
    ''' <param name="StringPart">array of strings/wildcards or ranges containing strings/wildcards</param>
    ''' <returns>database compliant string</returns>
    <ExcelFunction(Description:="Create a database compliant string from cell values, potentially concatenating with other parts for easy inclusion of wildcards (%,_)")>
    Public Function DBString(<ExcelArgument(Description:="array of strings/wildcards or ranges containing strings/wildcards")> ParamArray StringPart() As Object) As String
        Dim myRef, myCell
        Try
            Dim retval As String = String.Empty
            For Each myRef In StringPart
                If TypeName(myRef) = "Object(,)" Then
                    For Each myCell In myRef
                        If TypeName(myCell) = "ExcelEmpty" Then
                            ' do nothing here
                        Else
                            retval &= myCell.ToString()
                        End If
                    Next
                ElseIf TypeName(myRef) = "ExcelEmpty" Then
                    ' do nothing here
                Else
                    retval &= myRef.ToString()
                End If
            Next
            DBString = "'" & retval & "'"
        Catch ex As Exception
            WriteToLog("Error (" & ex.Message & ") in Functions.DBString", EventLogEntryType.Warning)
            DBString = "Error (" & ex.Message & ") in DBString"
        End Try
    End Function

    ''' <summary>Create an in clause from cell values, strings are created with quotation marks,
    '''             dates are created with DBDate</summary>
    ''' <param name="inPart">array of values or ranges containing values</param>
    ''' <returns>database compliant in-clause string</returns>
    <ExcelFunction(Description:="Create an in clause from cell values, strings are created with quotation marks, dates are created with DBDate")>
    Public Function DBinClause(<ExcelArgument(AllowReference:=True, Description:="array of values or ranges containing values")> ParamArray inPart As Object()) As String
        DBinClause = "in (" & DoConcatCellsSep(",", True, inPart) & ")"
    End Function

    ''' <summary>concatenates values contained in thetarget together (using .value attribute for cells)</summary>
    ''' <param name="thetarget">all cells/values which should be concatenated</param>
    ''' <returns>concatenated String</returns>
    <ExcelFunction(Description:="concatenates values contained in thetarget together (using .value attribute for cells)")>
    Public Function concatCells(<ExcelArgument(AllowReference:=True, Description:="all cells/values which should be concatenated")> ParamArray thetarget As Object()) As String
        concatCells = DoConcatCellsSep(String.Empty, False, thetarget)
    End Function

    ''' <summary>concatenates values contained in thetarget (using .value for cells) using a separator</summary>
    ''' <param name="separator">the separator</param>
    ''' <param name="thetarget">all cells/values which should be concatenated</param>
    ''' <returns>concatenated String</returns>
    <ExcelFunction(Description:="concatenates values contained in thetarget (using .value for cells) using a separator")>
    Public Function concatCellsSep(<ExcelArgument(AllowReference:=True, Description:="the separator")> separator As String,
                                   <ExcelArgument(AllowReference:=True, Description:="all cells/values which should be concatenated")> ParamArray thetarget As Object()) As String
        concatCellsSep = DoConcatCellsSep(separator, False, thetarget)
    End Function

    ''' <summary>chains values contained in thetarget together with commas, mainly used for creating select header</summary>
    ''' <param name="thetarget">range where values should be chained</param>
    ''' <returns>chained String</returns>
    <ExcelFunction(Description:="chains values contained in thetarget together with commas, mainly used for creating select header")>
    Public Function chainCells(<ExcelArgument(AllowReference:=True, Description:="range where values should be chained")> ParamArray thetarget As Object()) As String
        chainCells = DoConcatCellsSep(",", False, thetarget)
    End Function

    ''' <summary>private function that actually concatenates values contained in Object array myRange together (either using .text or .value for cells in myrange) using a separator</summary>
    ''' <param name="separator">the separator-string that is filled between values</param>
    ''' <param name="DBcompliant">should a potential string or date part be formatted database compliant (surrounded by quotes)?</param>
    ''' <param name="concatParts">Object array, whose values should be concatenated</param>
    ''' <returns>concatenated String</returns>
    Private Function DoConcatCellsSep(separator As String, DBcompliant As Boolean, ParamArray concatParts As Object()) As String
        Dim myRef, myCell

        Try
            Dim retval As String = String.Empty
            For Each myRef In concatParts
                Dim isMultiCell As Boolean = False
                If TypeName(myRef) = "ExcelReference" Then
                    ' pack into try to avoid exception for errors in passed cells in myRef (#NAME! etc.)
                    Try
                        isMultiCell = (TypeName(myRef.GetValue()) = "Object(,)")
                        If Not isMultiCell Then myRef = myRef.GetValue() ' single cell in reference -> convert to value
                    Catch ex As Exception
                        myRef = "ExcelError"
                    End Try
                End If
                If isMultiCell Then ' multiple cells in reference (range)
                    For Each myCell In myRef.GetValue()
                        If TypeName(myCell) = "ExcelEmpty" Then
                            ' do nothing here
                        ElseIf IsNumeric(myCell) Then
                            retval = retval & separator & Convert.ToString(myCell, System.Globalization.CultureInfo.InvariantCulture)
                        Else
                            ' avoid double quoting if passed string is already quoted (by using DBDate or DBString as input to this) and DBcompliant quoting is requested
                            retval = retval & separator & IIf(DBcompliant And Left(myCell, 1) <> "'", "'", "") & myCell & IIf(DBcompliant And Right(myCell, 1) <> "'", "'", "")
                        End If
                    Next
                Else
                    ' and other direct values in formulas..
                    If TypeName(myRef) = "ExcelEmpty" Then
                        ' do nothing here
                    ElseIf IsNumeric(myRef) Then ' no separate Date type for direct formula values
                        retval = retval & separator & Convert.ToString(myRef, System.Globalization.CultureInfo.InvariantCulture)
                    Else
                        ' avoid double quoting if passed string is already quoted (by using DBDate or DBString as input to this) and DBcompliant quoting is requested
                        retval = retval & separator & IIf(DBcompliant And Left(myRef, 1) <> "'", "'", "") & myRef & IIf(DBcompliant And Right(myRef, 1) <> "'", "'", "")
                    End If
                End If
            Next
            DoConcatCellsSep = Mid$(retval, Len(separator) + 1) ' skip first separator
            If DoConcatCellsSep = "" Then DoConcatCellsSep = "only empty arguments!"
        Catch ex As Exception
            WriteToLog("Error (" & ex.Message & ") in Functions.DoConcatCellsSep", EventLogEntryType.Warning)
            DoConcatCellsSep = "Error (" & ex.Message & ") in DoConcatCellsSep"
        End Try
    End Function

    ''' <summary>Stores a query into an Object defined in targetRange (an embedded MS Query/Listobject, Pivot table, etc.)</summary>
    ''' <param name="Query">query for getting data</param>
    ''' <param name="ConnString">connection string defining DB, user, etc...</param>
    ''' <param name="targetRange">Range with Object beneath to put the Query/ConnString into</param>
    ''' <returns>Status Message</returns>
    <ExcelFunction(Description:="Stores a query into an Object defined in targetRange (an embedded MS Query/Listobject, Pivot table, etc.)")>
    Public Function DBSetQuery(<ExcelArgument(Description:="query for getting data")> Query As Object,
                               <ExcelArgument(Description:="connection string defining DB, user, etc...")> ConnString As Object,
                               <ExcelArgument(Description:="Range with Object beneath to put the Query/ConnString into", AllowReference:=True)> targetRange As Object) As String
        Dim callID As String = ""
        Dim caller As Range
        Dim EnvPrefix As String = ""

        Try
            caller = ToRange(XlCall.Excel(XlCall.xlfCaller))
            resolveConnstring(ConnString, EnvPrefix)
            ' calcContainers are identified by wbname + Sheetname + function caller cell Address
            callID = "[" & caller.Parent.Parent.Name & "]" & caller.Parent.Name & "!" & caller.Address

            ' check query, also converts query to string (if it is a range)
            ' error message or cached status message is returned from checkParamsAndCache, if query OK and result was not already calculated (cached) then empty string
            DBSetQuery = checkParamsAndCache(Query, callID, ConnString)
            If DBSetQuery.Length > 0 Then
                DBSetQuery = EnvPrefix & ", " & DBSetQuery
                Exit Function
            End If

            ' first call: actually perform query
            If Not StatusCollection.Contains(callID) Then
                Dim statusCont As ContainedStatusMsg = New ContainedStatusMsg
                StatusCollection.Add(statusCont, callID)
                StatusCollection(callID).statusMsg = "" ' need this to prevent object not set errors in checkCache
                ExcelAsyncUtil.QueueAsMacro(Sub()
                                                DBSetQueryAction(callID, Query, targetRange, ConnString)
                                            End Sub)
            End If

        Catch ex As Exception
            WriteToLog("Error (" & ex.Message & "), callID : " & callID, EventLogEntryType.Warning)
            DBSetQuery = EnvPrefix & ", Error (" & ex.Message & ") in DBSetQueryAsync, callID : " & callID
        End Try
    End Function

    ''' <summary>set Query parameters (query text and connection string) of Query List or pivot table (incl. chart)</summary>
    ''' <param name="callID">the key for the statusMsg container</param>
    ''' <param name="Query"></param>
    ''' <param name="targetRange"></param>
    ''' <param name="ConnString"></param>
    Sub DBSetQueryAction(callID As String, Query As String, targetRange As ExcelReference, ConnString As String)
        Dim TargetCell As Range
        Dim targetSH As Worksheet
        Dim targetWB As Workbook
        Dim errMsg As String
        Dim thePivotTable As PivotTable = Nothing
        Dim theListObject As ListObject = Nothing

        Dim calcMode = hostApp.Calculation
        hostApp.Calculation = XlCalculation.xlCalculationManual
        TargetCell = ToRange(targetRange)
        targetSH = TargetCell.Parent
        targetWB = TargetCell.Parent.Parent

        ' try to get either a pivot table object or a list object from the target cell. What we have, is checked later...
        Try : thePivotTable = TargetCell.PivotTable : Catch ex As Exception : End Try
        Try : theListObject = TargetCell.ListObject : Catch ex As Exception : End Try

        Dim connType As String
        Dim bgQuery As Boolean
        Try
            StatusCollection(callID).statusMsg = "neither pivot table nor Listobject could be set..."
            If Not thePivotTable Is Nothing Then
                bgQuery = thePivotTable.PivotCache.BackgroundQuery
                Try
                    connType = Left$(thePivotTable.PivotCache.Connection, InStr(1, thePivotTable.PivotCache.Connection, ";"))
                Catch ex As Exception
                    MsgBox("couldn't get connection from Pivot Table, please create pivot table with external data source !", vbCritical + vbOKOnly, "Pivot Table assigning")
                    Throw ex
                End Try
                thePivotTable.PivotCache.Connection = connType & ConnString
                thePivotTable.PivotCache.CommandType = XlCmdType.xlCmdSql
                thePivotTable.PivotCache.CommandText = Query
                thePivotTable.PivotCache.BackgroundQuery = False
                thePivotTable.PivotCache.Refresh()
                StatusCollection(callID).statusMsg = "Set " & connType & " PivotTable to (bgQuery= " & bgQuery & "): " & Query
                thePivotTable.PivotCache.BackgroundQuery = bgQuery
            End If

            If Not theListObject Is Nothing Then
                bgQuery = theListObject.QueryTable.BackgroundQuery
                Try
                    connType = Left$(theListObject.QueryTable.Connection, InStr(1, theListObject.QueryTable.Connection, ";"))
                Catch ex As Exception
                    MsgBox("couldn't get connection from ListObject, please create ListObject with external data source !", vbCritical + vbOKOnly, "ListObject assigning")
                    Throw ex
                End Try
                ' Attention Dirty Hack ! This works only for SQLOLEDB driver to ODBC driver setting change...
                theListObject.QueryTable.Connection = connType & Replace(ConnString, "provider=SQLOLEDB", "driver=SQL SERVER")
                theListObject.QueryTable.CommandType = XlCmdType.xlCmdSql
                theListObject.QueryTable.CommandText = Query
                theListObject.QueryTable.BackgroundQuery = False
                theListObject.QueryTable.Refresh()
                StatusCollection(callID).statusMsg = "Set " & connType & " ListObject to (bgQuery= " & bgQuery & "): " & Query
                theListObject.QueryTable.BackgroundQuery = bgQuery
            End If
        Catch ex As Exception
            TargetCell.Cells(1, 1) = "" ' set first cell to ALWAYS trigger return of error messages to calling function
            errMsg = Err.Description & " in query: " & Query
            WriteToLog(errMsg & ", caller: " & callID, EventLogEntryType.Warning)
            StatusCollection(callID).statusMsg = errMsg
        End Try
        hostApp.Calculation = calcMode
    End Sub

    ''' <summary>
    ''' Fetches a list of data defined by query into TargetRange.
    ''' Optionally copy formulas contained in FormulaRange, extend list depending on ExtendDataArea (0(default) = overwrite, 1=insert Cells, 2=insert Rows)
    ''' and add field headers if HeaderInfo = TRUE
    ''' </summary>
    ''' <param name="Query">query for getting data</param>
    ''' <param name="ConnString">connection string defining DB, user, etc...</param>
    ''' <param name="targetRange">Range to put the data into</param>
    ''' <param name="formulaRange">Range to copy formulas down from</param>
    ''' <param name="extendDataArea">how to deal with extending List Area</param>
    ''' <param name="HeaderInfo">should headers be included in list</param>
    ''' <param name="AutoFit">should columns be autofitted ?</param>
    ''' <param name="autoformat">should 1st row formats be autofilled down?</param>
    ''' <param name="ShowRowNums">should row numbers be displayed in 1st column?</param>
    ''' <returns>Status Message, data values are returned outside of function cell (@see DBFuncEventHandler)</returns>
    <ExcelFunction(Description:="Fetches a list of data defined by query into TargetRange. Optionally copy formulas contained in FormulaRange, extend list depending on ExtendDataArea (0(default) = overwrite, 1=insert Cells, 2=insert Rows) and add field headers if HeaderInfo = TRUE")>
    Public Function DBListFetch(<ExcelArgument(Description:="query for getting data")> Query As Object,
                                <ExcelArgument(Description:="connection string defining DB, user, etc...")> ConnString As Object,
                                <ExcelArgument(Description:="Range to put the data into", AllowReference:=True)> targetRange As Object,
                                <ExcelArgument(Description:="Range to copy formulas down from", AllowReference:=True)> Optional formulaRange As Object = Nothing,
                                <ExcelArgument(Description:="how to deal with extending List Area")> Optional extendDataArea As Integer = 0,
                                <ExcelArgument(Description:="should headers be included in list")> Optional HeaderInfo As Boolean = False,
                                <ExcelArgument(Description:="should columns be autofitted ?")> Optional AutoFit As Boolean = False,
                                <ExcelArgument(Description:="should 1st row formats be autofilled down?")> Optional autoformat As Boolean = False,
                                <ExcelArgument(Description:="should row numbers be displayed in 1st column?")> Optional ShowRowNums As Boolean = False) As String
        Dim callID As String = ""
        Dim EnvPrefix As String = ""

        Try
            Dim caller As Range = ToRange(XlCall.Excel(XlCall.xlfCaller))
            resolveConnstring(ConnString, EnvPrefix)
            ' calcContainers are identified by wbname + Sheetname + function caller cell Address
            callID = "[" & caller.Parent.Parent.Name & "]" & caller.Parent.Name & "!" & caller.Address

            ' prepare information for action proc
            If dontCalcWhileClearing Then
                DBListFetch = EnvPrefix & ", dontCalcWhileClearing = True !"
                Exit Function
            End If
            If TypeName(targetRange) <> "ExcelReference" Then
                DBListFetch = EnvPrefix & ", Invalid targetRange or range name doesn't exist!"
                Exit Function
            End If
            If TypeName(formulaRange) <> "ExcelMissing" And TypeName(formulaRange) <> "ExcelReference" Then
                DBListFetch = EnvPrefix & ", Invalid FormulaRange or range name doesn't exist!"
                Exit Function
            End If

            ' check query, also converts query to string (if it is a range)
            ' error message or cached status message is returned from checkParamsAndCache, if query OK and result was not already calculated (cached) then empty string
            DBListFetch = checkParamsAndCache(Query, callID, ConnString)
            If DBListFetch.Length > 0 Then
                DBListFetch = EnvPrefix & ", " & DBListFetch
                Exit Function
            End If

            ' get target range name ...
            Dim functionArgs = functionSplit(caller.Formula, ",", """", "DBListFetch", "(", ")")
            Dim targetRangeName As String : targetRangeName = functionArgs(2)
            ' check if fetched argument targetRangeName is really a name or just a plain range address
            If Not existsNameInWb(targetRangeName, caller.Parent.Parent) And Not existsNameInSheet(targetRangeName, caller.Parent) Then targetRangeName = String.Empty
            ' get formula range name ...
            Dim formulaRangeName As String
            If UBound(functionArgs) > 2 Then
                formulaRangeName = functionArgs(3)
                If Not existsNameInWb(formulaRangeName, caller.Parent.Parent) And Not existsNameInSheet(formulaRangeName, caller.Parent) Then formulaRangeName = String.Empty
            Else
                formulaRangeName = String.Empty
            End If

            ' first call: Status Container not set, actually perform query
            If Not StatusCollection.Contains(callID) Then
                Dim statusCont As ContainedStatusMsg = New ContainedStatusMsg
                StatusCollection.Add(statusCont, callID)
                StatusCollection(callID).statusMsg = "" ' need this to prevent object not set errors in checkCache
                ExcelAsyncUtil.QueueAsMacro(Sub()
                                                DBListFetchAction(callID, CStr(Query), caller, ToRange(targetRange), CStr(ConnString), ToRange(formulaRange), extendDataArea, HeaderInfo, AutoFit, autoformat, ShowRowNums, targetRangeName, formulaRangeName)
                                            End Sub)
            End If
        Catch ex As Exception
            WriteToLog("Error (" & ex.Message & "), callID : " & callID, EventLogEntryType.Warning)
            DBListFetch = EnvPrefix & ", Error (" & ex.Message & "), callID : " & callID
        End Try
    End Function

    ''' <summary>Actually do the work for DBListFetch: Query list of data delimited by maxRows and maxCols, write it into targetCells
    '''             additionally copy formulas contained in formulaRange and extend list depending on extendArea</summary>
    ''' <param name="callID"></param>
    ''' <param name="Query"></param>
    ''' <param name="appCaller"></param>
    ''' <param name="targetRange"></param>
    ''' <param name="ConnString"></param>
    ''' <param name="formulaRange"></param>
    ''' <param name="extendArea"></param>
    ''' <param name="HeaderInfo"></param>
    ''' <param name="AutoFit"></param>
    ''' <param name="autoformat"></param>
    ''' <param name="ShowRowNumbers"></param>
    ''' <param name="targetRangeName"></param>
    ''' <param name="formulaRangeName"></param>
    Public Sub DBListFetchAction(callID As String, Query As String, appCaller As Object, targetRange As Range, ConnString As String, formulaRange As Object, extendArea As Integer, HeaderInfo As Boolean, AutoFit As Boolean, autoformat As Boolean, ShowRowNumbers As Boolean, targetRangeName As String, formulaRangeName As String)
        Dim tableRst As ADODB.Recordset
        Dim formulaFilledRange As Range = Nothing
        Dim targetSH As Worksheet, formulaSH As Worksheet = Nothing
        Dim copyFormat() As String = Nothing, copyFormatF() As String = Nothing
        Dim headingOffset As Long, rowDataStart As Long, startRow As Long, startCol As Long, arrayCols As Long, arrayRows As Long, copyDown As Long
        Dim oldRows As Long = 0, oldCols As Long = 0, oldFRows As Long = 0, oldFCols As Long = 0, retrievedRows As Long, targetColumns As Long, formulaStart As Long
        Dim warning As String, errMsg As String, tmpname As String
        Dim storedNames() As String

        LogInfo("Entering DBListFetchAction: callID " & callID)
        'If Not existsStatusCont(callID) Then Exit Sub
        Dim calcMode = hostApp.Calculation
        hostApp.Cursor = XlMousePointer.xlWait  ' To show the hourglass
        hostApp.Calculation = XlCalculation.xlCalculationManual
        ' this works around the data validation input bug
        ' when selecting a value from a list of validated field, excel won't react to
        ' Application.Calculation changes, so just leave here...
        If hostApp.Calculation <> XlCalculation.xlCalculationManual Then Exit Sub

        formulaRange = formulaRange
        targetSH = targetRange.Parent
        warning = String.Empty

        Dim srcExtentConnect As String, targetExtent As String, targetExtentF As String
        On Error Resume Next
        srcExtentConnect = appCaller.Name.Name
        If Err.Number <> 0 Or InStr(1, srcExtentConnect, "DBFsource") = 0 Then
            Err.Clear()
            srcExtentConnect = "DBFsource" & Replace(Replace(CDbl(Now.ToOADate()), ",", String.Empty), ".", String.Empty)
            appCaller.Name = srcExtentConnect
            appCaller.Parent.Parent.Names(srcExtentConnect).Visible = False
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

        startRow = targetRange.Cells.Row : startCol = targetRange.Cells.Column
        If Err.Number <> 0 Then
            errMsg = "Error in setting startRow/startCol: " & Err.Description & " in query: " & Query
            GoTo err_0
        End If

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

        Dim ODBCconnString As String = String.Empty
        If InStr(1, UCase$(ConnString), ";ODBC;") Then
            ODBCconnString = Mid$(ConnString, InStr(1, UCase$(ConnString), ";ODBC;") + 1)
            ConnString = Left$(ConnString, InStr(1, UCase$(ConnString), ";ODBC;") - 1)
        End If

        If conn Is Nothing Then conn = New ADODB.Connection
        If CurrConnString <> ConnString Then
            If conn.State <> 0 Then conn.Close()
            dontTryConnection = False
        End If

        If conn.State <> ADODB.ObjectStateEnum.adStateOpen And Not dontTryConnection Then
            conn.ConnectionTimeout = CnnTimeout
            conn.CommandTimeout = CmdTimeout
            conn.CursorLocation = CursorLocationEnum.adUseClient
            hostApp.StatusBar = "Trying " & CnnTimeout & " sec. with connstring: " & ConnString
            Err.Clear()
            conn.Open(ConnString)

            If Err.Number <> 0 Then
                WriteToLog("Connection Error: " & Err.Description, EventLogEntryType.Error)
                ' prevent multiple reconnecting if connection errors present...
                dontTryConnection = True
                StatusCollection(callID).statusMsg = "Connection Error: " & Err.Description
            End If
            CurrConnString = ConnString
        End If

        hostApp.StatusBar = "Retrieving data for DBList: " & IIf(targetRangeName.Length > 0, targetRangeName, targetSH.Name & "!" & targetRange.Address)
        tableRst = New ADODB.Recordset
        tableRst.Open(Query, conn, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly, CommandTypeEnum.adCmdText)
        Dim dberr As String = String.Empty
        If conn.Errors.Count > 0 Then
            Dim errcount As Integer
            For errcount = 0 To conn.Errors.Count - 1
                If conn.Errors.Item(errcount).Description <> Err.Description Then dberr = dberr & ";" & conn.Errors.Item(errcount).Description
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
        Dim aborted As Boolean = XlCall.Excel(XlCall.xlAbort) ' for long running actions, allow interruption
        If aborted Then
            errMsg = "data fetching interrupted by user !"
            GoTo err_1
        End If

        ' from now on we don't propagate any errors as data is modified in sheet....
        hostApp.StatusBar = "Displaying data for DBList: " & IIf(targetRangeName.Length > 0, targetRangeName, targetSH.Name & "!" & targetRange.Address)
        If tableRst.EOF Then warning = "Warning: No Data returned in query: " & Query
        ' set size for named range (size: arrayRows, arrayCols) used for resizing the data area (old extent)
        arrayCols = tableRst.Fields.Count
        arrayRows = retrievedRows
        ' need to shift down 1 row if headings are present
        arrayRows += IIf(HeaderInfo, 1, 0)
        rowDataStart = 1 + IIf(HeaderInfo, 1, 0)

        ' check whether retrieved data exceeds excel's limits and limit output (arrayRows/arrayCols) in case ...
        ' check rows
        If targetRange.Row + arrayRows > (targetRange.EntireColumn.Rows.Count + 1) Then
            warning = "row count" & " of returned data exceeds max row of excel: start row:" & targetRange.Row & " + row count:" & arrayRows & " > max row+1:" & targetRange.EntireColumn.Rows.Count + 1
            arrayRows = targetRange.EntireColumn.Rows.Count - targetRange.Row + 1
        End If
        ' check columns
        If targetRange.Column + arrayCols > (targetRange.EntireRow.Columns.Count + 1) Then
            warning = warning & ", column count" & " of returned data exceed max column of excel: start column:" & targetRange.Column & " + column count:" & arrayCols & " > max column+1:" & targetRange.EntireRow.Columns.Count + 1
            arrayCols = targetRange.EntireRow.Columns.Count - targetRange.Column + 1
        End If

        ' autoformat: copy 1st rows formats range to reinsert them afterwards
        targetColumns = arrayCols - IIf(ShowRowNumbers, 0, 1)
        If autoformat Then
            arrayRows += IIf(HeaderInfo And arrayRows = 1, 1, 0)  ' need special case for autoformat
            Dim i As Long
            For i = 0 To targetColumns
                ReDim Preserve copyFormat(i)
                copyFormat(i) = targetSH.Cells(targetRange.Row + rowDataStart - 1, targetRange.Column + i).NumberFormat
            Next
            ' now for the calculated data area
            If Not formulaRange Is Nothing Then
                For i = 0 To formulaRange.Columns.Count - 1
                    ReDim Preserve copyFormatF(i)
                    copyFormatF(i) = formulaSH.Cells(targetRange.Row + rowDataStart - 1, formulaRange.Column + i).NumberFormat
                Next
            End If
        End If
        If arrayRows = 0 Then arrayRows = 1  ' sane behavior of named range in case no data retrieved...

        ' check if formulaRange and targetRange overlap !
        Dim possibleIntersection As Range = hostApp.Intersect(formulaRange, targetSH.Range(targetRange.Cells(1, 1), targetRange.Cells(1, 1).Offset(0, arrayCols - 1)))
        Err.Clear()
        If Not possibleIntersection Is Nothing Then
            warning = warning & ", formulaRange and targetRange intersect (" & targetSH.Name & "!" & possibleIntersection.Address & "), formula copying disabled !!"
            formulaRange = Nothing
        End If

        '''' data list and formula range extension (ignored in first call after creation -> no defined name is set -> oldRows=0)...
        headingOffset = IIf(HeaderInfo, 1, 0)  ' use that for generally regarding headings !!
        If oldRows > 0 Then
            ' either cells/rows are shifted down (old data area was smaller than current) ...
            If oldRows < arrayRows Then
                'prevent insertion from heading row if headings are present (to not get the header formats..)
                Dim headingFirstRowPrevent As Long = IIf(HeaderInfo And oldRows = 1 And arrayRows > 2, 1, 0)
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
                Dim headingLastRowPrevent As Long = IIf(HeaderInfo And arrayRows = 1 And oldRows > 2, 1, 0)
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
            With targetSH.QueryTables.Add(Connection:=ODBCconnString, Destination:=targetRange)
                .CommandText = Query
                .FieldNames = HeaderInfo
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
            With targetSH.QueryTables.Add(Connection:=tableRst, Destination:=targetRange)
                .FieldNames = HeaderInfo
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
                StatusCollection(callID).statusMsg = "Retrieved " & retrievedRows & " record" & IIf(retrievedRows > 1, "s", String.Empty) & ", Warning: " & warning
            Else
                StatusCollection(callID).statusMsg = warning
            End If
        Else
            StatusCollection(callID).statusMsg = "Retrieved " & retrievedRows & " record" & IIf(retrievedRows > 1, "s", String.Empty) & " from: " & Query
        End If

        ' autoformat: restore format of 1st row...
        If autoformat Then
            For i = 0 To UBound(copyFormat)
                newTargetRange.Rows(rowDataStart).Cells(i + 1).NumberFormat = copyFormat(i)
            Next
            ' now for the calculated cells...
            If Not formulaRange Is Nothing Then
                For i = 0 To UBound(copyFormatF)
                    formulaSH.Cells(targetRange.Row + rowDataStart - 1, formulaRange.Column + i).NumberFormat = copyFormatF(i)
                Next
            End If
            'auto format 1st rows down...
            If arrayRows > rowDataStart Then
                'This doesn't work anymore:
                'newTargetRange.Rows(rowDataStart).AutoFill(Destination:=newTargetRange.Rows(rowDataStart & ":" & arrayRows), Type:=XlAutoFillType.xlFillFormats)
                targetSH.Range(targetSH.Cells(targetRange.Row + rowDataStart - 1, newTargetRange.Column), targetSH.Cells(targetRange.Row + rowDataStart - 1, newTargetRange.Column + newTargetRange.Columns.Count - 1)).AutoFill(Destination:=targetSH.Range(targetSH.Cells(targetRange.Row + rowDataStart - 1, newTargetRange.Column), targetSH.Cells(targetRange.Row + arrayRows - 1, newTargetRange.Column + newTargetRange.Columns.Count - 1)), Type:=XlAutoFillType.xlFillFormats)
                If Not formulaRange Is Nothing Then
                    formulaSH.Range(formulaSH.Cells(targetRange.Row + rowDataStart - 1, formulaRange.Column), formulaSH.Cells(targetRange.Row + rowDataStart - 1, formulaRange.Column + formulaRange.Columns.Count - 1)).AutoFill(Destination:=formulaSH.Range(formulaSH.Cells(targetRange.Row + rowDataStart - 1, formulaRange.Column), formulaSH.Cells(targetRange.Row + arrayRows - 1, formulaRange.Column + formulaRange.Columns.Count - 1)), Type:=XlAutoFillType.xlFillFormats)
                End If
            End If
        End If

        If Err.Number <> 0 Then
            errMsg = "Error in restoring formats: " & Err.Description & " in query: " & Query
            GoTo err_0
        End If

        'auto fit columns AFTER autoformat so we don't have problems with applied formats visibility ...
        If AutoFit Then
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
        finishAction(calcMode, callID)
        Exit Sub

err_2: ' errors where recordset was opened and QueryTables were already added, but temp names were not deleted
        targetSH.Names(tmpname).Delete
        targetSH.Parent.Names(tmpname).Delete
err_1: ' errors where recordset was opened
        If tableRst.State <> 0 Then tableRst.Close()
err_0: ' errors where recordset was not opened or is already closed
        targetRange.Cells(1, 1) = "" ' target to dirty to ALWAYS trigger return of error messages to calling function
        If errMsg.Length = 0 Then errMsg = Err.Description & " in query: " & Query
        WriteToLog(errMsg & ", caller: " & callID, EventLogEntryType.Warning)
        StatusCollection(callID).statusMsg = errMsg
        finishAction(calcMode, callID, "Error")
    End Sub

    ''' <summary>common sub to finish the action procedures, resetting anything that was set otherwise...</summary>
    ''' <param name="calcMode"></param>
    ''' <param name="callID"></param>
    ''' <param name="additionalLogInfo"></param>
    Private Sub finishAction(calcMode As XlCalculation, callID As String, Optional additionalLogInfo As String = "")
        hostApp.Cursor = XlMousePointer.xlDefault  ' To return cursor to normal
        hostApp.StatusBar = False
        LogInfo("Leaving DBListFetchAction: callID " & callID & IIf(additionalLogInfo <> "", ", additionalInfo: " & additionalLogInfo, ""))
        If Not IsNothing(origWS) Then origWS.Select()
        hostApp.ScreenUpdating = True ' coming from refresh, this might be off for dirtying "foreign" (being on a different sheet than the calling function) data targets 
        hostApp.Calculation = calcMode
    End Sub

    ''' <summary>Fetches a row (single record) queried (defined in query) from DB (defined in ConnString) into targetArray</summary>
    ''' <param name="Query">query for getting data</param>
    ''' <param name="ConnString">connection string defining DB, user, etc...</param>
    ''' <param name="targetArray">Range to put the data into</param>
    ''' <returns>Status Message, data values are returned outside of function cell (@see DBFuncEventHandler)</returns>
    <ExcelFunction(Description:="Fetches a row (single record) queried (defined in query) from DB (defined in ConnString) into targetArray")>
    Public Function DBRowFetch(<ExcelArgument(Description:="query for getting data")> Query As Object,
                               <ExcelArgument(Description:="connection string defining DB, user, etc...")> ConnString As Object,
                               <ExcelArgument(Description:="Range to put the data into", AllowReference:=True)> ParamArray targetArray() As Object) As String
        Dim tempArray() As Range = Nothing ' final target array that is passed to makeCalcMsgContainer (after removing header flag)
        Dim callID As String = ""
        Dim HeaderInfo As Boolean
        Dim EnvPrefix As String = ""

        Try
            Dim caller As Range = ToRange(XlCall.Excel(XlCall.xlfCaller))
            resolveConnstring(ConnString, EnvPrefix)
            ' calcContainers are identified by wbname + sheetname + function caller cell Address
            callID = "[" & caller.Parent.Parent.Name & "]" & caller.Parent.Name & "!" & caller.Address

            ' prepare information for action proc
            Dim i As Long
            If TypeName(targetArray(0)) = "Boolean" Then
                HeaderInfo = targetArray(0)
                For i = 1 To UBound(targetArray)
                    ReDim Preserve tempArray(i - 1)
                    tempArray(i - 1) = ToRange(targetArray(i))
                Next
            ElseIf TypeName(targetArray(0)) = "Error" Then
                DBRowFetch = EnvPrefix & ", Error: First argument empty or error !"
                Exit Function
            Else
                For i = 0 To UBound(targetArray)
                    ReDim Preserve tempArray(i)
                    tempArray(i) = ToRange(targetArray(i))
                Next
            End If
            ' check query, also converts query to string (if it is a range)
            ' error message or cached status message is returned from checkParamsAndCache, if query OK and result was not already calculated (cached) then empty string
            DBRowFetch = checkParamsAndCache(Query, callID, ConnString)
            If DBRowFetch.Length > 0 Then
                DBRowFetch = EnvPrefix & ", " & DBRowFetch
                Exit Function
            End If

            ' first call: actually perform query
            If Not StatusCollection.Contains(callID) Then
                Dim statusCont As ContainedStatusMsg = New ContainedStatusMsg
                StatusCollection.Add(statusCont, callID)
                StatusCollection(callID).statusMsg = "" ' need this to prevent object not set errors in checkCache
                ExcelAsyncUtil.QueueAsMacro(Sub()
                                                DBRowFetchAction(callID, CStr(Query), caller, tempArray, CStr(ConnString), HeaderInfo)
                                            End Sub)
            End If
        Catch ex As Exception
            WriteToLog("Error (" & ex.Message & "), callID : " & callID, EventLogEntryType.Warning)
            DBRowFetch = EnvPrefix & ", Error (" & ex.Message & "), callID : " & callID
        End Try
    End Function

    ''' <summary>Actually do the work for DBRowFetch: Query (assumed) one row of data, write it into targetCells</summary>
    ''' <param name="callID"></param>
    ''' <param name="Query"></param>
    ''' <param name="appCaller"></param>
    ''' <param name="targetArray"></param>
    ''' <param name="ConnString"></param>
    ''' <param name="HeaderInfo"></param>
    Public Sub DBRowFetchAction(callID As String, Query As String, appCaller As Object, targetArray As Object, ConnString As String, HeaderInfo As Boolean)
        Dim tableRst As ADODB.Recordset = Nothing
        Dim targetCells As Object
        Dim errMsg As String = String.Empty, refCollector As Range
        Dim headerFilled As Boolean, DeleteExistingContent As Boolean, fillByRows As Boolean
        Dim returnedRows As Long, fieldIter As Integer, rangeIter As Integer
        Dim theCell As Range, targetSlice As Range, targetSlices As Range
        Dim targetSH As Worksheet

        Dim calcMode = hostApp.Calculation
        hostApp.Calculation = XlCalculation.xlCalculationManual
        ' this works around the data validation input bug
        ' when selecting a value from a list of validated field, excel won't react to
        ' Application.Calculation changes, so just leave here...
        If hostApp.Calculation <> XlCalculation.xlCalculationManual Then Exit Sub

        hostApp.Cursor = XlMousePointer.xlWait  ' To show the hourglass
        targetCells = targetArray
        targetSH = targetCells(0).Parent
        StatusCollection(callID).statusMsg = ""
        On Error GoTo err_1
        hostApp.StatusBar = "Retrieving data for DBRows: " & targetSH.Name & "!" & targetCells(0).Address

        Dim srcExtentConnect As String, targetExtent As String
        On Error Resume Next
        srcExtentConnect = appCaller.Name.Name
        If Err.Number <> 0 Or InStr(1, UCase$(srcExtentConnect), "DBFSOURCE") = 0 Then
            Err.Clear()
            srcExtentConnect = "DBFsource" & Replace(Replace(CDbl(Now().ToOADate()), ",", String.Empty), ".", String.Empty)
            appCaller.Name = srcExtentConnect
            ' dbfsource is a workbook name
            appCaller.Parent.Parent.Names(srcExtentConnect).Visible = False
            If Err.Number <> 0 Then
                errMsg = "Error in setting srcExtentConnect name: " & Err.Description & " in query: " & Query
                GoTo err_1
            End If
        End If
        targetExtent = Replace(srcExtentConnect, "DBFsource", "DBFtarget")
        ' remove old data in case we changed the target range array
        targetSH.Range(targetExtent).ClearContents()

        If conn Is Nothing Then conn = New ADODB.Connection
        If CurrConnString <> ConnString Then
            If conn.State <> 0 Then conn.Close()
            dontTryConnection = False
        End If

        If conn.State <> ADODB.ObjectStateEnum.adStateOpen And Not dontTryConnection Then
            conn.ConnectionTimeout = CnnTimeout
            conn.CommandTimeout = CmdTimeout
            conn.CursorLocation = CursorLocationEnum.adUseClient
            hostApp.StatusBar = "Trying " & CnnTimeout & " sec. with connstring: " & ConnString
            Err.Clear()
            conn.Open(ConnString)

            If Err.Number <> 0 Then
                WriteToLog("Connection Error: " & Err.Description, EventLogEntryType.Error)
                ' prevent multiple reconnecting if connection errors present...
                dontTryConnection = True
                StatusCollection(callID).statusMsg = "Connection Error: " & Err.Description
            End If
            CurrConnString = ConnString
        End If

        tableRst = New ADODB.Recordset
        tableRst.Open(Query, conn, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly, CommandTypeEnum.adCmdText)
        On Error Resume Next
        Dim dberr As String = String.Empty
        If conn.Errors.Count > 0 Then
            Dim errcount As Integer
            For errcount = 0 To conn.Errors.Count - 1
                If conn.Errors.Item(errcount).Description <> Err.Description Then _
                   dberr = dberr & ";" & conn.Errors.Item(errcount).Description
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
        ' check whether anything retrieved? if not, delete possible existing content...
        DeleteExistingContent = tableRst.EOF
        If DeleteExistingContent Then StatusCollection(callID).statusMsg = "Warning: No Data returned in query: " & Query

        ' if "heading range" is present then orientation of first range (header) defines layout of data: if "heading range" is column then data is returned columnwise, else row by row.
        ' if there is just one block of data then it is assumed that there are usually more rows than columns and orientation is set by row/column size
        fillByRows = IIf(UBound(targetCells) > 0, targetCells(0).Rows.Count < targetCells(0).Columns.Count, targetCells(0).Rows.Count > targetCells(0).Columns.Count)
        ' put values (single record) from Recordset into targetCells
        fieldIter = 0 : rangeIter = 0 : headerFilled = Not HeaderInfo    ' if we don't need headers the assume they are filled already....
        refCollector = targetCells(0)
        Do
            If fillByRows Then
                targetSlices = targetCells(rangeIter).Rows
            Else
                targetSlices = targetCells(rangeIter).Columns
            End If
            For Each targetSlice In targetSlices
                Dim aborted As Boolean = XlCall.Excel(XlCall.xlAbort) ' for long running actions, allow interruption
                If aborted Then
                    errMsg = "data fetching interrupted by user !"
                    GoTo err_1
                End If
                For Each theCell In targetSlice.Cells
                    If tableRst.EOF Then
                        theCell.Value = String.Empty
                    Else
                        If Not headerFilled Then
                            theCell.Value = tableRst.Fields(fieldIter).Name
                        ElseIf DeleteExistingContent Then
                            theCell.Value = String.Empty
                        Else
                            On Error Resume Next
                            theCell.Value = tableRst.Fields(fieldIter).Value
                            If Err.Number <> 0 Then errMsg &= "Field '" & tableRst.Fields(fieldIter).Name & "' caused following error: '" & Err.Description & "'"
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
        If StatusCollection(callID).statusMsg.Length = 0 Then StatusCollection(callID).statusMsg = "Retrieved " & returnedRows & " record" & IIf(returnedRows > 1, "s", String.Empty) & " from: " & Query
        finishAction(calcMode, callID)
        Exit Sub

err_1:
        targetCells(0).Cells(1, 1) = "" ' target to dirty to ALWAYS trigger return of error messages to calling function
        If errMsg.Length = 0 Then errMsg = Err.Description & " in query: " & Query
        If tableRst.State <> 0 Then tableRst.Close()
        WriteToLog(errMsg & ", caller: " & callID, EventLogEntryType.Warning)
        StatusCollection(callID).statusMsg = errMsg
        finishAction(calcMode, callID, "Error")
    End Sub

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
        nextName = Target.Name.Name
        Do
            If Err.Number = 0 And nextName <> theName Then
                ReDim Preserve storedNames(i)
                storedNames(i) = nextName
                i += 1
            End If
            Target.Name.Delete
            nextName = Target.Name.Name
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

    ''' <summary>Get the current selected Environment for DB Functions</summary>
    ''' <returns>ConfigName of environment</returns>
    <ExcelFunction(Description:="Get the current selected Environment for DB Functions")>
    Public Function DBAddinEnvironment() As String
        hostApp.Volatile()
        DBAddinEnvironment = fetchSetting("ConfigName", String.Empty)
        If hostApp.Calculation = XlCalculation.xlCalculationManual Then DBAddinEnvironment = "calc Mode is manual, please press F9 to get current DBAddin environment !"
    End Function

    ''' <summary>Get the server settings for the currently selected Environment for DB Functions</summary>
    ''' <returns>Server part from connection string of environment</returns>
    <ExcelFunction(Description:="Get the server settings for the currently selected Environment for DB Functions")>
    Public Function DBAddinServerSetting() As String
        hostApp.Volatile()
        Try
            Dim theConnString As String = fetchSetting("ConstConnString", String.Empty)
            Dim keywordstart As Integer = InStr(1, theConnString, "Server=") + Len("Server=")
            DBAddinServerSetting = Mid$(theConnString, keywordstart, InStr(keywordstart, theConnString, ";") - keywordstart)
            If hostApp.Calculation = XlCalculation.xlCalculationManual Then DBAddinServerSetting = "calc Mode is manual, please press F9 to get current DBAddin server setting !"
        Catch ex As Exception
            DBAddinServerSetting = "Error happened: " & Err.Description
        End Try
    End Function

    ''' <summary>checks calculation mode, query and cached status message.</summary>
    ''' <param name="Query"></param>
    ''' <param name="callID"></param>
    ''' <param name="ConnString"></param>
    ''' <returns>Error String or cached status message (empty if OK)</returns>
    Private Function checkParamsAndCache(ByRef Query, callID, ConnString) As String
        checkParamsAndCache = ""
        If hostApp.Calculation = XlCalculation.xlCalculationManual Then
            checkParamsAndCache = "calc Mode is manual, please press F9 to trigger data fetching !"
        Else
            If TypeName(Query) = "ExcelEmpty" Then
                checkParamsAndCache = "empty query provided !"
            ElseIf Left(TypeName(Query), 10) = "ExcelError" Then
                If Query = ExcelError.ExcelErrorValue Then
                    checkParamsAndCache = "query contains: #Val! (in case query is an argument of a DBfunction, check if it's > 255 chars)"
                Else
                    checkParamsAndCache = "query contains: #" + Replace(Query.ToString(), "ExcelError", "") + "!"
                End If
            ElseIf TypeName(Query) = "Object(,)" Then
                ' if query is reference then get the query string out of it..
                Dim myCell
                Dim retval As String = ""
                For Each myCell In Query
                    If TypeName(myCell) = "ExcelEmpty" Then
                        'do nothing here
                    ElseIf Left(TypeName(myCell), 10) = "ExcelError" Then
                        If myCell = ExcelError.ExcelErrorValue Then
                            checkParamsAndCache = "query contains: #Val! (in case query is an argument of a DBfunction, check if it's > 255 chars)"
                        Else
                            checkParamsAndCache = "query contains: #" + Replace(myCell.ToString(), "ExcelError", "") + "!"
                        End If
                    ElseIf IsNumeric(myCell) Then
                        retval &= Convert.ToString(myCell, System.Globalization.CultureInfo.InvariantCulture) & " "
                    Else
                        retval &= myCell & " "
                    End If
                    Query = retval
                Next
                If retval.Length = 0 Then checkParamsAndCache = "empty query provided !"
            ElseIf TypeName(Query) = "String" Then
                If Query.Length = 0 Then checkParamsAndCache = "empty query provided !"
            Else
                checkParamsAndCache = "query parameter invalid (not a range and not a string) !"
            End If
        End If

        ' caching check mechanism to avoid unnecessary recalculations/refetching
        Dim doFetching As Boolean
        If queryCache.Contains(callID) Then
            doFetching = (ConnString & Query <> queryCache(callID))
            ' refresh the query cache with new query/connstring ...
            queryCache.Remove(callID)
            queryCache.Add(ConnString & Query, callID)
        Else
            queryCache.Add(ConnString & Query, callID)
            doFetching = True
        End If
        If doFetching Then
            ' remove Status Container to signal a new calculation request
            If StatusCollection.Contains(callID) Then StatusCollection.Remove(callID)
        Else
            ' return Status Containers Message as last result
            If StatusCollection.Contains(callID) Then checkParamsAndCache = StatusCollection(callID).statusMsg
        End If
    End Function


    ''' <summary>create a final connection string from passed String or number (environment), as well as a EnvPrefix for showing the environment (or set ConnString)</summary>
    ''' <param name="ConnString">passed connection string or environment number, resolved (=returned) to actual connection string</param>
    ''' <param name="EnvPrefix">prefix for showing environment (ConnString set if no environment)</param>
    Sub resolveConnstring(ByRef ConnString As Object, ByRef EnvPrefix As String)
        If Left(TypeName(ConnString), 10) = "ExcelError" Then Exit Sub
        If TypeName(ConnString) = "ExcelReference" Then ConnString = ConnString.Value
        If TypeName(ConnString) = "ExcelMissing" Then ConnString = ""
        ' in case ConnString is a number (set environment, retrieve ConnString from Setting ConstConnString<Number>
        If TypeName(ConnString) = "Double" Then
            EnvPrefix = "Env:" & fetchSetting("ConfigName" & ConnString.ToString(), String.Empty)
            ConnString = fetchSetting("ConstConnString" & ConnString.ToString(), String.Empty)
        ElseIf TypeName(ConnString) = "String" Then
            If ConnString = "" Then ' no ConnString or environment number set: get current set connection string
                EnvPrefix = "Env:" & fetchSetting("ConfigName", String.Empty)
                ConnString = fetchSetting("ConstConnString", String.Empty)
            Else
                EnvPrefix = "ConnString set"
            End If
        End If
    End Sub

    ''' <summary>checks whether theName exists as a name in Workbook theWb</summary>
    ''' <param name="theName"></param>
    ''' <param name="theWb"></param>
    ''' <returns>true if it exists</returns>
    Private Function existsNameInWb(ByRef theName As String, theWb As Workbook) As Boolean
        existsNameInWb = False
        For Each aName In theWb.Names()
            If aName.Name = theName Then
                existsNameInWb = True
                Exit Function
            End If
        Next
    End Function

    ''' <summary>checks whether theName exists as a name in Worksheet theWs</summary>
    ''' <param name="theName"></param>
    ''' <param name="theWs"></param>
    ''' <returns>true if it exists</returns>
    Private Function existsNameInSheet(ByRef theName As String, theWs As Worksheet) As Boolean
        existsNameInSheet = False
        For Each aName In theWs.Names()
            If aName.Name = theWs.Name & "!" & theName Then
                existsNameInSheet = True
                Exit Function
            End If
        Next
    End Function

    ''' <summary>converts ExcelDna (C API) reference to excel (COM Based) Range</summary>
    ''' <param name="reference">reference to be converted</param>
    ''' <returns>range for passed reference</returns>
    Private Function ToRange(reference As Object) As Range
        If TypeName(reference) <> "ExcelReference" Then Return Nothing

        Dim item As String = XlCall.Excel(XlCall.xlSheetNm, reference)
        Dim index As Integer = item.LastIndexOf("]")
        item = item.Substring(index + 1)
        Dim ws As Worksheet = ExcelDnaUtil.Application.Sheets(item)
        Dim target As Range = ws.Range(ws.Cells(reference.RowFirst + 1, reference.ColumnFirst + 1), ws.Cells(reference.RowLast + 1, reference.ColumnLast + 1))
        Return target
    End Function

End Module
