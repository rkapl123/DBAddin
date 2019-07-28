Imports Microsoft.Office.Interop.Excel
Imports ExcelDna.Integration

''' <summary>Contains the public callable DB functions and helper functions</summary>
Public Module Functions

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
            callID = "[" & caller.Parent.Parent.name & "]" & caller.Parent.name & "!" & caller.Address

            ' check query, also converts query to string (if it is a range)
            DBSetQuery = checkParams(Query)
            ' error message is returned from checkParams, if OK then returns nothing
            If DBSetQuery.Length > 0 Then
                DBSetQuery = EnvPrefix & ", checkParams error: " & DBSetQuery
                Exit Function
            End If

            ' second call (we're being set to dirty in calc event handler)
            If existsCalcCont(callID) Then
                If allCalcContainers(callID).errOccured Then
                    ' commented this to prevent endless loops !!
                    'allCalcContainers.Remove callID
                    ' special case for invocations from function wizard
                ElseIf Not allCalcContainers(callID).working Then
                    allCalcContainers.Remove(callID)
                    makeCalcMsgContainer(callID, CStr(Query), caller, Nothing, ToRange(targetRange), CStr(ConnString), Nothing, 0, False, False, False, False, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, False)
                End If
            Else
                ' reset status messages when starting new query...
                If existsStatusCont(callID) Then allStatusContainers(callID).statusMsg = String.Empty
                ' add transportation info for event proc
                makeCalcMsgContainer(callID, CStr(Query), caller, Nothing, ToRange(targetRange), CStr(ConnString), Nothing, 0, False, False, False, False, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, False)
            End If
            If existsStatusCont(callID) Then
                DBSetQuery = EnvPrefix & ", statusMsg: " & allStatusContainers(callID).statusMsg
            Else
                DBSetQuery = EnvPrefix & ", no recalculation done for unchanged query..."
            End If
            hostApp.EnableEvents = True
        Catch ex As Exception
            WriteToLog("Error (" & ex.Message & ") in Functions.DBSetQuery, callID : " & callID, EventLogEntryType.Warning)
            DBSetQuery = EnvPrefix & ", Error (" & ex.Message & ") in DBSetQuery, callID : " & callID
            hostApp.EnableEvents = True
        End Try
    End Function

    ''' <summary>Stores a query into an Object defined in targetRange (an embedded MS Query/Listobject, Pivot table, etc.)</summary>
    ''' <param name="Query">query for getting data</param>
    ''' <param name="ConnString">connection string defining DB, user, etc...</param>
    ''' <param name="targetRange">Range with Object beneath to put the Query/ConnString into</param>
    ''' <returns>Status Message</returns>
    <ExcelFunction(Description:="Stores a query into an Object defined in targetRange (an embedded MS Query/Listobject, Pivot table, etc.)")>
    Public Function DBSetQueryAsync(<ExcelArgument(Description:="query for getting data")> Query As Object,
                               <ExcelArgument(Description:="connection string defining DB, user, etc...")> ConnString As Object,
                               <ExcelArgument(Description:="Range with Object beneath to put the Query/ConnString into", AllowReference:=True)> targetRange As Object) As String
        Dim callID As String = ""
        Dim caller As Range
        Dim EnvPrefix As String = ""
        Try
            caller = ToRange(XlCall.Excel(XlCall.xlfCaller))
            resolveConnstring(ConnString, EnvPrefix)
            ' calcContainers are identified by wbname + Sheetname + function caller cell Address
            callID = "[" & caller.Parent.Parent.name & "]" & caller.Parent.name & "!" & caller.Address

            ' check query, also converts query to string (if it is a range)
            DBSetQueryAsync = checkParams(Query)
            ' error message is returned from checkParams, if OK then returns nothing
            If DBSetQueryAsync.Length > 0 Then
                DBSetQueryAsync = EnvPrefix & ", checkParams error: " & DBSetQueryAsync
                Exit Function
            End If

            ' second call (we're being set to dirty in calc event handler)
            If existsCalcCont(callID) Then
                If allCalcContainers(callID).errOccured Then
                    ' commented this to prevent endless loops !!
                    'allCalcContainers.Remove callID
                    ' special case for invocations from function wizard
                ElseIf Not allCalcContainers(callID).working Then
                    allCalcContainers.Remove(callID)
                    makeCalcMsgContainer(callID, CStr(Query), caller, Nothing, ToRange(targetRange), CStr(ConnString), Nothing, 0, False, False, False, False, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, False)
                End If
            Else
                ' reset status messages when starting new query...
                If existsStatusCont(callID) Then allStatusContainers(callID).statusMsg = String.Empty
                ' add transportation info for event proc
                makeCalcMsgContainer(callID, CStr(Query), caller, Nothing, ToRange(targetRange), CStr(ConnString), Nothing, 0, False, False, False, False, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, False)
            End If
            ExcelAsyncUtil.QueueAsMacro(Sub()
                                            DBSetQueryAction(allCalcContainers(callID), allStatusContainers(callID))
                                        End Sub
                                        )
            If existsStatusCont(callID) Then
                DBSetQueryAsync = EnvPrefix & ", statusMsg: " & allStatusContainers(callID).statusMsg
            Else
                DBSetQueryAsync = EnvPrefix & ", no recalculation done for unchanged query..."
            End If
            hostApp.EnableEvents = True
        Catch ex As Exception
            WriteToLog("Error (" & ex.Message & ") in Functions.DBSetQuery, callID : " & callID, EventLogEntryType.Warning)
            DBSetQueryAsync = EnvPrefix & ", Error (" & ex.Message & ") in DBSetQuery, callID : " & callID
            hostApp.EnableEvents = True
        End Try
    End Function

    ''' <summary>set Query parameters (query text and connection string) of Query List or pivot table (incl. chart)</summary>
    ''' <param name="calcCont"><see cref="ContainerCalcMsgs"/></param>
    ''' <param name="statusCont"><see cref="ContainerStatusMsgs"/></param>
    Public Sub DBSetQueryAction(calcCont As ContainerCalcMsgs, statusCont As ContainerStatusMsgs)
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
        WriteToLog("DBFuncEventHandler.DBSetQueryParams Error: " & errMsg & ", caller: " & callID, EventLogEntryType.Warning)

        statusCont.statusMsg = errMsg
        ' need to mark calc container here as excel won't return to main event proc in case of error
        ' calc container is then removed in calling function
        allCalcContainers(callID).errOccured = True
        allCalcContainers(callID).callsheet.Range(allCalcContainers(callID).caller.Address).Dirty
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
        Dim callID As String
        Dim caller As Range = ToRange(XlCall.Excel(XlCall.xlfCaller))
        Dim EnvPrefix As String = ""
        resolveConnstring(ConnString, EnvPrefix)

2:      If dontCalcWhileClearing Then
3:          DBListFetch = EnvPrefix & ", dontCalcWhileClearing = True !"
            Exit Function
        End If
4:      If TypeName(targetRange) <> "ExcelReference" Then
5:          DBListFetch = EnvPrefix & ", Invalid targetRange or range name doesn't exist!"
            Exit Function
        End If
        On Error GoTo DBListFetch_Error
        ' calcContainers are identified by wbname + Sheetname + function caller cell Address
6:      callID = "[" & caller.Parent.Parent.name & "]" & caller.Parent.name & "!" & caller.Address
        LogInfo("entering DBListFetch:" & callID)
12:     If TypeName(formulaRange) <> "ExcelMissing" And TypeName(formulaRange) <> "ExcelReference" Then
13:         DBListFetch = EnvPrefix & ", Invalid FormulaRange or range name doesn't exist!"
            Exit Function
        End If

        ' get target range name ...
14:     Dim functionArgs = functionSplit(caller.Formula, ",", """", "DBListFetch", "(", ")")
15:     Dim targetRangeName As String : targetRangeName = functionArgs(2)
        ' check if fetched argument targetRangeName is really a name or just a plain range address
16:     If Not existsNameInWb(targetRangeName, caller.Parent.Parent) And Not existsNameInSheet(targetRangeName, caller.Parent) Then targetRangeName = String.Empty

        ' get formula range name ...
        Dim formulaRangeName As String
17:     If UBound(functionArgs) > 2 Then
18:         formulaRangeName = functionArgs(3)
19:         If Not existsNameInWb(formulaRangeName, caller.Parent.Parent) And Not existsNameInSheet(formulaRangeName, caller.Parent) Then formulaRangeName = String.Empty
        Else
            formulaRangeName = String.Empty
        End If

        ' check query, also converts query to string (if it is a range)
20:     DBListFetch = checkParams(Query)
        ' error message is returned from checkParams, if OK then returns nothing
21:     If DBListFetch.Length > 0 Then
22:         DBListFetch = EnvPrefix & ", " & DBListFetch
            Exit Function
        End If

        ' second call (we're being set to dirty in calc event handler)
23:     If existsCalcCont(callID) Then
24:         If allCalcContainers(callID).errOccured Then
                ' commented this to prevent endless loops !!
                'allCalcContainers.Remove callID
                ' special case for invocations from function wizard
25:         ElseIf Not allCalcContainers(callID).working Then
26:             allCalcContainers.Remove(callID)
27:             makeCalcMsgContainer(callID, CStr(Query), caller, Nothing, ToRange(targetRange), CStr(ConnString), ToRange(formulaRange), extendDataArea, HeaderInfo, AutoFit, autoformat, ShowRowNums, String.Empty, String.Empty, String.Empty, String.Empty, targetRangeName, formulaRangeName, False)
            End If
        Else
            ' reset status messages when starting new query...
28:         If existsStatusCont(callID) Then allStatusContainers(callID).statusMsg = String.Empty
            ' add transportation info for event proc
29:         makeCalcMsgContainer(callID, CStr(Query), caller, Nothing, ToRange(targetRange), CStr(ConnString), ToRange(formulaRange), extendDataArea, HeaderInfo, AutoFit, autoformat, ShowRowNums, String.Empty, String.Empty, String.Empty, String.Empty, targetRangeName, formulaRangeName, False)
        End If
30:     If existsStatusCont(callID) Then
31:         DBListFetch = EnvPrefix & ", " & allStatusContainers(callID).statusMsg
        Else
32:         DBListFetch = EnvPrefix & ", no recalculation done for unchanged query..."
        End If
        LogInfo("leaving DBListFetch:" & callID)
        hostApp.EnableEvents = True
        Exit Function

DBListFetch_Error:
        Dim ErrDesc As String = Err.Description
        WriteToLog("Error (" & ErrDesc & ") in Functions.DBListFetch, callID : " & callID & ", in " & Erl(), EventLogEntryType.Warning)
        DBListFetch = EnvPrefix & ", Error (" & ErrDesc & ") in DBListFetch, callID : " & callID & ", in " & Erl()
        hostApp.EnableEvents = True
    End Function

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
        Dim callID As String
        Dim HeaderInfo As Boolean
        Dim caller As Range = ToRange(XlCall.Excel(XlCall.xlfCaller))
        Dim EnvPrefix As String = ""
        resolveConnstring(ConnString, EnvPrefix)

        On Error GoTo DBRowFetch_Error
        ' calcContainers are identified by wbname + Sheetname + function caller cell Address
        callID = "[" & caller.Parent.Parent.name & "]" & caller.Parent.name & "!" & caller.Address
        DBRowFetch = checkParams(Query)
        If DBRowFetch.Length > 0 Then
            DBRowFetch = EnvPrefix & ", " & DBRowFetch
            Exit Function
        End If

        ' add transportation info for event proc
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
        If existsCalcCont(callID) Then
            If allCalcContainers(callID).errOccured Then
                ' commented this to prevent endless loops !!
                'allCalcContainers.Remove callID
                ' special case for intermediate invocation in function wizard
            ElseIf Not allCalcContainers(callID).working Then
                allCalcContainers.Remove(callID)
                makeCalcMsgContainer(callID, CStr(Query), caller, tempArray, Nothing, CStr(ConnString), Nothing, 0, HeaderInfo, False, False, False, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, False)
            End If
        Else
            ' reset status messages when starting new query...
            If existsStatusCont(callID) Then allStatusContainers(callID).statusMsg = String.Empty
            ' add transportation info for event proc
            makeCalcMsgContainer(callID, CStr(Query), caller, tempArray, Nothing, CStr(ConnString), Nothing, 0, HeaderInfo, False, False, False, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, False)
        End If
        If existsStatusCont(callID) Then DBRowFetch = EnvPrefix & ", " & allStatusContainers(callID).statusMsg
        hostApp.EnableEvents = True

        Exit Function
DBRowFetch_Error:
        Dim ErrDesc As String = Err.Description
        WriteToLog("Error (" & ErrDesc & ") in DBRowFetch, callID : " & callID & ", in " & Erl(), EventLogEntryType.Warning)
        DBRowFetch = EnvPrefix & ", Error (" & ErrDesc & ") in Functions.DBRowFetch, callID : " & callID
        hostApp.EnableEvents = True
    End Function

    Public Function DBAddinEnvironment() As String
        hostApp.Volatile
        DBAddinEnvironment = fetchSetting("ConfigName", String.Empty)
        If hostApp.Calculation = XlCalculation.xlCalculationManual Then DBAddinEnvironment = "calc Mode is manual, please press F9 to get current DBAddin environment !"
    End Function

    Public Function DBAddinServerSetting() As String
        Dim keywordstart As Integer
        Dim theConnString As String

        hostApp.Volatile
        On Error Resume Next
        theConnString = fetchSetting("ConstConnString", String.Empty)
        keywordstart = InStr(1, theConnString, "Server=") + Len("Server=")
        DBAddinServerSetting = Mid$(theConnString, keywordstart, InStr(keywordstart, theConnString, ";") - keywordstart)
        If hostApp.Calculation = XlCalculation.xlCalculationManual Then DBAddinServerSetting = "calc Mode is manual, please press F9 to get current DBAddin server setting !"
        If Err.Number <> 0 Then DBAddinServerSetting = "Error happened: " & Err.Description
    End Function

    ''' <summary>checks query and calculation mode if OK for both DBListFetch and DBRowFetch function</summary>
    ''' <param name="Query"></param>
    ''' <returns>Error String (empty if OK)</returns>
    Private Function checkParams(ByRef Query) As String
        checkParams = String.Empty
        If hostApp.Calculation = XlCalculation.xlCalculationManual Then
            checkParams = "calc Mode is manual, please press F9 to trigger data fetching !"
        Else
            If TypeName(Query) = "ExcelEmpty" Then
                checkParams = "empty query provided !"
            ElseIf Left(TypeName(Query), 10) = "ExcelError" Then
                If Query = ExcelError.ExcelErrorValue Then
                    checkParams = "query contains: #Val! (in case query is an argument of a DBfunction, check if it's > 255 chars)"
                Else
                    checkParams = "query contains: #" + Replace(Query.ToString(), "ExcelError", "") + "!"
                End If
            ElseIf TypeName(Query) = "Object(,)" Then
                ' if query is reference then get the query string out of it..
                Dim myCell
                Dim retval As String = String.Empty
                For Each myCell In Query
                    If TypeName(myCell) = "ExcelEmpty" Then
                        'do nothing here
                    ElseIf Left(TypeName(myCell), 10) = "ExcelError" Then
                        If myCell = ExcelError.ExcelErrorValue Then
                            checkParams = "query contains: #Val! (in case query is an argument of a DBfunction, check if it's > 255 chars)"
                        Else
                            checkParams = "query contains: #" + Replace(myCell.ToString(), "ExcelError", "") + "!"
                        End If
                    ElseIf IsNumeric(myCell) Then
                        retval &= Convert.ToString(myCell, System.Globalization.CultureInfo.InvariantCulture) & " "
                    Else
                        retval &= myCell & " "
                    End If
                    Query = retval
                Next
                If retval.Length = 0 Then checkParams = "empty query provided !"
            ElseIf TypeName(Query) = "String" Then
                If Query.Length = 0 Then checkParams = "empty query provided !"
            Else
                checkParams = "query parameter invalid (not a range and not a string) !"
            End If
        End If
    End Function

    ''' <summary>build/renew transport containers for functions</summary>
    ''' <param name="callID">the key for the calc msg container</param>
    ''' <param name="Query"></param>
    ''' <param name="appCaller"></param>
    ''' <param name="targetArray"></param>
    ''' <param name="targetRange"></param>
    ''' <param name="ConnString"></param>
    ''' <param name="formulaRange"></param>
    ''' <param name="extendArea"></param>
    ''' <param name="HeaderInfo"></param>
    ''' <param name="AutoFit"></param>
    ''' <param name="autoformat"></param>
    ''' <param name="ShowRowNumbers"></param>
    ''' <param name="colSep"></param>
    ''' <param name="rowSep"></param>
    ''' <param name="lastColSep"></param>
    ''' <param name="lastRowSep"></param>
    ''' <param name="targetRangeName"></param>
    ''' <param name="formulaRangeName"></param>
    ''' <param name="InterleaveHeader"></param>
    ''' <remarks>
    ''' for all other parameters, <see cref="ContainerCalcMsgs"/>
    ''' </remarks>
    Private Sub makeCalcMsgContainer(ByRef callID As String, ByRef Query As String, appCaller As Object, targetArray As Object, ByRef targetRange As Range, ByRef ConnString As String, ByRef formulaRange As Object, ByRef extendArea As Integer, ByRef HeaderInfo As Boolean, ByRef AutoFit As Boolean, ByRef autoformat As Boolean, ByRef ShowRowNumbers As Boolean, ByRef colSep As String, ByRef rowSep As String, ByRef lastColSep As String, ByRef lastRowSep As String, ByRef targetRangeName As String, ByRef formulaRangeName As String, ByRef InterleaveHeader As Boolean)
        Dim myCalcCont As ContainerCalcMsgs

        On Error GoTo makeCalcMsgContainer_Error
        ' setup event processing class and container carrying function information...
        If targetFilterCont Is Nothing Then targetFilterCont = New Collection
        If theDBFuncEventHandler Is Nothing Then theDBFuncEventHandler = New DBFuncEventHandler
        If allCalcContainers Is Nothing Then allCalcContainers = New Collection
        ' add components to calc container
        myCalcCont = New ContainerCalcMsgs
        myCalcCont.errOccured = False
        myCalcCont.Query = Query
        myCalcCont.caller = appCaller           'Application.caller
        myCalcCont.callsheet = appCaller.Parent  'Application.caller.Parent
        myCalcCont.targetArray = targetArray
        myCalcCont.targetRange = targetRange
        If ConnString.Length > 0 Then
            myCalcCont.ConnString = ConnString
        Else
            myCalcCont.ConnString = ConstConnString
        End If
        myCalcCont.formulaRange = formulaRange
        myCalcCont.extendArea = extendArea
        myCalcCont.HeaderInfo = HeaderInfo
        myCalcCont.AutoFit = AutoFit
        myCalcCont.autoformat = autoformat
        myCalcCont.ShowRowNumbers = ShowRowNumbers
        myCalcCont.colSep = colSep
        myCalcCont.rowSep = rowSep
        myCalcCont.lastColSep = lastColSep
        myCalcCont.lastRowSep = lastRowSep
        myCalcCont.targetRangeName = targetRangeName
        myCalcCont.formulaRangeName = formulaRangeName
        myCalcCont.InterleaveHeader = InterleaveHeader
        myCalcCont.callID = callID
        myCalcCont.working = False
        'add to global collection of all calc containers
        allCalcContainers.Add(myCalcCont, callID)

        Exit Sub
makeCalcMsgContainer_Error:
        If Err.Number <> 457 Then
            WriteToLog("Error (" & Err.Description & ") in Functions.makeCalcMsgContainer, callID: " & callID & ", in " & Erl(), EventLogEntryType.Warning)
        End If
    End Sub

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
            If ConnString = "" Then
                EnvPrefix = "Env:" & fetchSetting("ConfigName", String.Empty)
            Else
                EnvPrefix = "ConnString set"
            End If
        End If
    End Sub

    ''' <summary>check whether a calcContainer exists in allCalcContainers or not</summary>
    ''' <param name="theName">name of calcContainer</param>
    ''' <returns>true if it exists</returns>
    Private Function existsCalcCont(ByVal theName As String) As Boolean
        Dim dummy As String

        On Error GoTo err1
        existsCalcCont = True
        dummy = allCalcContainers(theName).Query
        Exit Function
err1:
        Err.Clear()
        existsCalcCont = False
    End Function

    ''' <summary>check whether a statusMsgContainer exists in allStatusContainers or not</summary>
    ''' <param name="theName">name of statusMsgContainer</param>
    ''' <returns>true if it exists</returns>
    Private Function existsStatusCont(ByVal theName As String) As Boolean
        Dim dummy As String

        On Error GoTo err1
        existsStatusCont = True
        dummy = allStatusContainers(theName).statusMsg
        Exit Function
err1:
        Err.Clear()
        existsStatusCont = False
    End Function

    ''' <summary>checks whether theName exists as a name in Workbook theWb</summary>
    ''' <param name="theName"></param>
    ''' <param name="theWb"></param>
    ''' <returns>true if it exists</returns>
    Private Function existsNameInWb(ByRef theName As String, theWb As Workbook) As Boolean
        Dim dummy As Name

        On Error GoTo err1
        existsNameInWb = True
        dummy = theWb.Names(theName)
        Exit Function
err1:
        Err.Clear()
        existsNameInWb = False
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

    ''' <summary>checks whether theName exists as a name in Worksheet theWs</summary>
    ''' <param name="theName"></param>
    ''' <param name="theWs"></param>
    ''' <returns>true if it exists</returns>
    Private Function existsNameInSheet(ByRef theName As String, theWs As Worksheet) As Boolean
        Dim dummy As Name

        On Error GoTo err1
        existsNameInSheet = True
        dummy = theWs.Names(theName)
        Exit Function
err1:
        Err.Clear()
        existsNameInSheet = False
    End Function

End Module
