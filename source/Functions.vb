Imports Microsoft.Office.Interop.Excel
Imports ExcelDna.Integration
Imports ExcelDna.Integration.XlCall

''' <summary>Contains the public callable DB functions and some helper functions</summary>
Public Module Functions

    ''' <summary>Create database compliant date, time or datetime string from excel datetype value</summary>
    ''' <param name="datVal">date/time/datetime</param>
    ''' <param name="formatting">see remarks</param>
    ''' <returns>the DB compliant formatted date/time/datetime</returns>
    ''' <remarks>
    ''' formatting = 0: A simple datestring (format 'YYYYMMDD'), datetime values are converted to 'YYYYMMDD HH:MM:SS' and time values are converted to 'HH:MM:SS'.
    ''' formatting = 1: An ANSI compliant Date string (format date 'YYYY-MM-DD'), datetime values are converted to timestamp 'YYYY-MM-DD HH:MM:SS' and time values are converted to time time 'HH:MM:SS'.
    ''' formatting = 2: An ODBC compliant Date string (format {d 'YYYY-MM-DD'}), datetime values are converted to {ts 'YYYY-MM-DD HH:MM:SS'} and time values are converted to {t 'HH:MM:SS'}.
    ''' formatting = 99 (default value): take the formatting option from setting DefaultDBDateFormatting (0 if not given)
    ''' </remarks>
    <ExcelFunction(Description:="Create database compliant date, time or datetime string from excel datetype value")>
    Public Function DBDate(<ExcelArgument(Description:="date/time/datetime")> ByVal datVal As Date,
                           <ExcelArgument(Description:="formatting option")> Optional formatting As Integer = 99) As String
        On Error GoTo DBDate_Error
        If formatting = 99 Then formatting = DefaultDBDateFormatting
        If Int(datVal.ToOADate()) = datVal.ToOADate() Then
            If formatting = 0 Then
                DBDate = "'" & Format$(datVal, "yyyyMMdd") & "'"
            ElseIf formatting = 1 Then
                DBDate = "date '" & Format$(datVal, "yyyy-MM-dd") & "'"
            ElseIf formatting = 2 Then
                DBDate = "{d '" & Format$(datVal, "yyyy-MM-dd") & "'}"
            End If
        ElseIf CInt(datVal.ToOADate()) > 1 Then
            If formatting = 0 Then
                DBDate = "'" & Format$(datVal, "yyyyMMdd hh:mm:ss") & "'"
            ElseIf formatting = 1 Then
                DBDate = "timestamp '" & Format$(datVal, "yyyy-MM-dd hh:mm:ss") & "'"
            ElseIf formatting = 2 Then
                DBDate = "{ts '" & Format$(datVal, "yyyy-MM-dd hh:mm:ss") & "'}"
            End If
        Else
            If formatting = 0 Then
                DBDate = "'" & Format$(datVal, "hh:mm:ss") & "'"
            ElseIf formatting = 1 Then
                DBDate = "time '" & Format$(datVal, "hh:mm:ss") & "'"
            ElseIf formatting = 2 Then
                DBDate = "{t '" & Format$(datVal, "hh:mm:ss") & "'}"
            End If
        End If
        Exit Function

DBDate_Error:
        Dim ErrDesc As String = Err.Description
        LogToEventViewer("Error (" & ErrDesc & ") in Functions.DBDate, in " & Erl(), EventLogEntryType.Error)
        DBDate = "Error (" & ErrDesc & ") in function DBDate"
    End Function

    ''' <summary>Create a database compliant string from cell values, potentially concatenating with other parts for easy inclusion of wildcards (%,_)</summary>
    ''' <param name="StringPart">array of strings/wildcards or ranges containing strings/wildcards</param>
    ''' <returns>database compliant string</returns>
    <ExcelFunction(Description:="Create a database compliant string from cell values, potentially concatenating with other parts for easy inclusion of wildcards (%,_)")>
    Public Function DBString(<ExcelArgument(Description:="array of strings/wildcards or ranges containing strings/wildcards")> ParamArray StringPart() As Object) As String
        Dim myRange
        Dim myCell
        Dim retval As String = vbNullString

        On Error GoTo DBString_Error
        For Each myRange In StringPart
            If TypeName(myRange) = "Object(,)" Then
                For Each myCell In myRange
                    If TypeName(myCell) = "ExcelEmpty" Then
                        ' do nothing here
                    Else
                        retval = retval & myCell.ToString()
                    End If
                Next
            ElseIf TypeName(myRange) = "ExcelEmpty" Then
                ' do nothing here
            Else
                retval = retval & myRange.ToString()
            End If
        Next
        DBString = "'" & retval & "'"
        Exit Function

DBString_Error:
        Dim ErrDesc As String = Err.Description
        LogToEventViewer("Error (" & ErrDesc & ") in Functions.DBString, in " & Erl(), EventLogEntryType.Error)
        DBString = "Error (" & ErrDesc & ") in DBString"
    End Function

    ''' <summary>Create an in clause from cell values, strings are created with quotation marks,
    '''             dates are created with DBDate (see below)</summary>
    ''' <param name="inPart">array of values or ranges containing values</param>
    ''' <returns>database compliant in-clause string</returns>
    Public Function DBinClause(ParamArray inPart() As Object) As String
        Dim myRange
        Dim Cell As Range = Nothing
        Dim inlist As String = vbNullString

        On Error GoTo DBinClause_Error
        For Each myRange In inPart
            If TypeName(myRange) = "Range" Then
                For Each Cell In myRange
                    If Cell Is ExcelEmpty.Value Then
                    ElseIf IsNumeric(Cell) Then
                        inlist = inlist & "," & Cell.Value
                    ElseIf IsDate(Cell.Value2) Then
                        inlist = inlist & "," & DBDate(Cell.Value2)
                    Else
                        inlist = inlist & ",'" & Cell.Value & "'"
                    End If
                Next
            Else
                If myRange Is ExcelEmpty.Value Then
                ElseIf IsNumeric(myRange) Then
                    inlist = inlist & "," & myRange
                ElseIf IsDate(myRange) Then
                    inlist = inlist & "," & DBDate(myRange)
                Else
                    inlist = inlist & ",'" & myRange & "'"
                End If
            End If
        Next
        DBinClause = "in (" & Mid$(inlist, 2) & ")"
        Exit Function
DBinClause_Error:
        If VBDEBUG Then Debug.Print("Error (" & Err.Description & ") in DBinClause") : Stop : Resume
        LogToEventViewer("Error (" & Err.Description & ") in Functions.DBinClause, in " & Erl(), EventLogEntryType.Error)
        DBinClause = "Error (" & Err.Description & ") in DBinClause"
    End Function

    ''' <summary>private function that actually concatenates values contained in Object array myRange together (either using .text or .value for cells in myrange)</summary>
    ''' <param name="asText">should cell values be taken as displayed (.text attribute) or their value (.value attribute)</param>
    ''' <param name="myRange">Object array, whose values should be concatenated</param>
    ''' <returns>concatenated String</returns>
    Private Function DoConcatCells(asText As Boolean, myRange As Object) As String
        Dim Cell As Range
        Dim retval As String = vbNullString

        On Error GoTo DoConcatCells_Error
        If TypeName(myRange) = "Range" Then
            For Each Cell In myRange
                If Cell.ToString().Length > 0 Then
                    retval = retval & CStr(IIf(asText, Cell.Text, Cell.Value))
                End If
            Next
            ' this happens when functions are called in matrix context
        ElseIf TypeName(myRange) = "Variant()" Then
            Dim cellValue
            For Each cellValue In myRange
                If TypeName(cellValue) = "Boolean" Then cellValue = IIf(cellValue, cellValue, vbNullString)
                If cellValue.ToString().Length > 0 Then
                    retval = retval & CStr(cellValue)
                End If
            Next
        Else
            retval = retval & CStr(myRange)
        End If
        DoConcatCells = retval
        Exit Function

DoConcatCells_Error:
        If VBDEBUG Then Debug.Print("Error (" & Err.Description & ") in DoConcatCells ") : Stop : Resume
        LogToEventViewer("Error (" & Err.Description & ") in Functions.DoConcatCells, in " & Erl(), EventLogEntryType.Error)
        DoConcatCells = "Error (" & Err.Description & ") !"
    End Function

    ''' <summary>concatenates values contained in thetarget together (using .value attribute for cells)</summary>
    ''' <param name="thetarget">all cells/values which should be concatenated</param>
    ''' <returns>concatenated String</returns>
    Public Function concatCells(ParamArray thetarget() As Object) As String
        Dim myRange
        concatCells = vbNullString
        For Each myRange In thetarget
            concatCells = concatCells & DoConcatCells(False, myRange)
        Next
    End Function

    ''' <summary>concatenates values contained in thetarget together (using .text for cells)</summary>
    ''' <param name="thetarget">all cells/values which should be concatenated</param>
    ''' <returns>concatenated String</returns>
    Public Function concatCellsText(ParamArray thetarget() As Object) As String
        Dim myRange
        concatCellsText = vbNullString
        For Each myRange In thetarget
            concatCellsText = concatCellsText & DoConcatCells(True, myRange)
        Next
    End Function

    ''' <summary>private function that actually concatenates values contained in Object array myRange together (either using .text or .value for cells in myrange) using a separator</summary>
    ''' <param name="asText">should cell values be taken as displayed (.text attribute) or their value (.value attribute)</param>
    ''' <param name="separator">the separator-string that is filled between values</param>
    ''' <param name="myRange">Object array, whose values should be concatenated</param>
    ''' <returns>concatenated String</returns>
    Private Function DoConcatCellsSep(asText As Boolean, separator As String, myRange As Object) As String
        Dim Cell As Range
        Dim retval As String = vbNullString
        Dim cellValueStr As String

        On Error GoTo DoConcatCellsSep_Error
        If TypeName(myRange) = "Range" Then
            For Each Cell In myRange
                If Not ((TypeName(Cell) = "Boolean" And Cell.Value2 = False) Or Cell.ToString().Length = 0) Then
                    On Error Resume Next
                    cellValueStr = CStr(IIf(asText, Cell.Text, Cell.Value))
                    If Err.Number <> 0 Then cellValueStr = CStr(Cell.Value)
                    On Error GoTo DoConcatCellsSep_Error
                    retval = retval & cellValueStr & separator
                End If
            Next
        ElseIf TypeName(myRange) = "Variant()" Then
            Dim cellValue
            For Each cellValue In myRange
                If TypeName(cellValue) = "Boolean" Then cellValue = IIf(cellValue, cellValue, vbNullString)
                If cellValue.ToString().Length > 0 Then
                    retval = retval & CStr(cellValue) & separator
                End If
            Next
        Else
            retval = retval & CStr(myRange) & separator
        End If
        DoConcatCellsSep = retval
        Exit Function

DoConcatCellsSep_Error:
        If VBDEBUG Then Debug.Print("Error (" & Err.Description & ") in DoConcatCellsSep ") : Stop : Resume
        LogToEventViewer("Error (" & Err.Description & ") in Functions.DoConcatCellsSep, in " & Erl(), EventLogEntryType.Error)
        DoConcatCellsSep = "Error (" & Err.Description & ") !"
    End Function

    ''' <summary>concatenates values contained in thetarget (using .value for cells) using a separator</summary>
    ''' <param name="separator">the separator</param>
    ''' <param name="thetarget">all cells/values which should be concatenated</param>
    ''' <returns>concatenated String</returns>
    Public Function concatCellsSep(separator As String, ParamArray thetarget() As Object) As String
        Dim myRange
        concatCellsSep = vbNullString
        For Each myRange In thetarget
            concatCellsSep = concatCellsSep & DoConcatCellsSep(False, separator, myRange)
        Next
        concatCellsSep = getRidOfLastSep(separator, concatCellsSep)
    End Function

    ''' <summary>concatenates values contained in thetarget using a separator using cells text property instead of the value (displayed)</summary>
    ''' <param name="separator">the separator</param>
    ''' <param name="thetarget">all cells/values which should be concatenated</param>
    ''' <returns>concatenated String</returns>
    Public Function concatCellsSepText(separator As String, ParamArray thetarget() As Object) As String
        Dim myRange
        concatCellsSepText = vbNullString
        For Each myRange In thetarget
            concatCellsSepText = concatCellsSepText & DoConcatCellsSep(True, separator, myRange)
        Next
        concatCellsSepText = getRidOfLastSep(separator, concatCellsSepText)
    End Function

    ''' <summary>gets rid of separator at end of totalString</summary>
    ''' <param name="separator">the separator</param>
    ''' <param name="totalString">the string to be modified</param>
    ''' <returns>modified String</returns>
    Private Function getRidOfLastSep(separator As String, totalString As String) As String
        If Len(totalString) > Len(separator) Then
            getRidOfLastSep = Left$(totalString, Len(totalString) - Len(separator))
        Else
            getRidOfLastSep = vbNullString
        End If
    End Function

    ''' <summary>chains values contained in thetarget together with commas, mainly used for creating select header</summary>
    ''' <param name="thetarget">range where values should be chained</param>
    ''' <returns>chained String</returns>
    Public Function chainCells(ParamArray thetarget() As Object) As String
        Dim myRange
        Dim Cell As Range
        Dim retval As String = vbNullString

        On Error GoTo chainCells_Error
        For Each myRange In thetarget
            If TypeName(myRange) = "Range" Then
                For Each Cell In myRange
                    If Cell.ToString().Length > 0 Then
                        retval = retval & Cell.ToString() & ","
                    End If
                Next
            Else
                retval = retval & CStr(myRange) & ","
            End If
        Next
        If Len(retval) > 1 Then
            chainCells = Left$(retval, Len(retval) - 1)
        Else
            chainCells = vbNullString
        End If
        Exit Function
chainCells_Error:
        If VBDEBUG Then Debug.Print("Error (" & Err.Description & ") in chainCells ") : Stop : Resume
        'LogToEventViewer("Error (" & Err.Description & ") in Functions.chainCells, in " & Erl(), EventLogEntryType.Error, 1)
        chainCells = "Error (" & Err.Description & ") in chainCells "
    End Function

    ''' <summary>creates a Listbox or Dropdown filled with data defined by query</summary>
    ''' <param name="Query">query for getting data</param>
    ''' <param name="ConnString">connection string defining DB, user, etc...</param>
    ''' <param name="controlType">type of control to be inserted (0 = Listbox, 1 = Dropdown)</param>
    ''' <param name="HeaderInfo">should header label be included in list</param>
    ''' <param name="autoArrange"></param>
    ''' <param name="ControlName"></param>
    ''' <param name="dataTargetRange">Range (String) to put the selection into (default = left cell from function address)</param>
    ''' <param name="controlLocation"></param>
    ''' <param name="subscribeTo">Range (String), where control/header should be placed (default = function address)</param>
    ''' <returns>Status Message</returns>
    Public Function DBMakeControl(Query As Object, ConnString As Object, Optional controlType As Integer = 0, Optional HeaderInfo As Boolean = False, Optional autoArrange As Boolean = False, Optional ControlName As String = vbNullString, Optional dataTargetRange As String = vbNullString, Optional controlLocation As String = vbNullString, Optional subscribeTo As String = vbNullString) As String
        Dim callID As String
        Dim setEnv As String

        On Error GoTo DBMakeControl_Error
        setEnv = fetchSetting("ConfigName", vbNullString)
        ' calcContainers are identified by wbname + Sheetname + function caller cell Address
        callID = "[" & theHostApp.caller.Parent.Parent.name & "]" & theHostApp.caller.Parent.name & "!" & theHostApp.caller.Address
        If TypeName(ConnString) = "Error" Then ConnString = vbNullString
        If TypeName(ConnString) = "Range" Then ConnString = ConnString.Value
        ' in case of number as connection string, take the stored Connection string .. 1 usually prod, 2 .. usually test, 3.. development)
        If TypeName(ConnString) = "Integer" Or TypeName(ConnString) = "Long" Or TypeName(ConnString) = "Double" Or TypeName(ConnString) = "Short" Then
            setEnv = fetchSetting("ConfigName" & ConnString, vbNullString)
            ConnString = fetchSetting("ConstConnString" & ConnString, vbNullString)
        End If

        DBMakeControl = checkParams(Query)
        If DBMakeControl.Length > 0 Then
            DBMakeControl = "Env:" & setEnv & ", " & DBMakeControl
            Exit Function
        End If
        If existsCalcCont(callID) Then
            If allCalcContainers(callID).errOccured Then
#If DEBUGME = 1 Then
              LogToEventViewer "DBControlQuery returned Error, removing container with callID = " & callID, LogInf, 0
#End If
                ' commented this to prevent endless loops !!
                'allCalcContainers.Remove callID
                ' special case for intermediate invocation in function wizard
            ElseIf Not allCalcContainers(callID).working Then
                allCalcContainers.Remove(callID)
                makeCalcMsgContainer(callID, CStr(Query), theHostApp.caller, Nothing, Nothing, CStr(ConnString), Nothing, 0, HeaderInfo, autoArrange, False, False, controlType, dataTargetRange, controlLocation, ControlName, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, False)
            End If
        Else
            ' reset status messages when starting new query...
            If existsStatusCont(callID) Then allStatusContainers(callID).statusMsg = vbNullString
            ' add transportation info for event proc
            makeCalcMsgContainer(callID, CStr(Query), theHostApp.caller, Nothing, Nothing, CStr(ConnString), Nothing, 0, HeaderInfo, autoArrange, False, False, controlType, dataTargetRange, controlLocation, ControlName, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, False)
        End If
        If existsStatusCont(callID) Then DBMakeControl = "Env:" & setEnv & ", " & allStatusContainers(callID).statusMsg
        theHostApp.EnableEvents = True
        Exit Function

DBMakeControl_Error:
        If VBDEBUG Then Debug.Print("Error (" & Err.Description & ") in DBMakeControl, callID : " & callID) : Stop : Resume
        LogToEventViewer("Error (" & Err.Description & ") in Functions.DBMakeControl, callID : " & callID & ", in " & Erl(), EventLogEntryType.Error)
        DBMakeControl = "Env:" & setEnv & ", Error (" & Err.Description & ") in DBMakeControl, callID : " & callID
        theHostApp.EnableEvents = True
    End Function

    ''' <summary>Stores a query into an Object defined in targetRange (an embedded MS Query/Listobject, Pivot table, etc.)</summary>
    ''' <param name="Query">query for getting data</param>
    ''' <param name="ConnString">connection string defining DB, user, etc...</param>
    ''' <param name="targetRange">Range with Object beneath to put the Query/ConnString into</param>
    ''' <returns>Status Message</returns>
    Public Function DBSetQuery(Query As Object, ConnString As Object, targetRange As Range) As String
        Dim callID As String
        Dim setEnv As String

        On Error GoTo DBSetQuery_Error
        setEnv = fetchSetting("ConfigName", vbNullString)
        ' calcContainers are identified by wbname + Sheetname + function caller cell Address
        'If Not IsObject(theHostApp.caller) Then Exit Function
        callID = "[" & theHostApp.caller.Parent.Parent.name & "]" & theHostApp.caller.Parent.name & "!" & theHostApp.caller.Address
        If TypeName(ConnString) = "Error" Then ConnString = vbNullString
        If TypeName(ConnString) = "Range" Then ConnString = ConnString.Value
        ' in case of number as connection string, take the stored Connection string .. 1 usually prod, 2 .. usually test, 3.. development)
        If TypeName(ConnString) = "Integer" Or TypeName(ConnString) = "Long" Or TypeName(ConnString) = "Double" Or TypeName(ConnString) = "Short" Then
            setEnv = fetchSetting("ConfigName" & ConnString, vbNullString)
            ConnString = fetchSetting("ConstConnString" & ConnString, vbNullString)
        End If
        ' check query, also converts query to string (if it is a range)
        DBSetQuery = checkParams(Query)
        ' error message is returned from checkParams, if OK then returns nothing
        If DBSetQuery.Length > 0 Then
            DBSetQuery = "Env:" & setEnv & ", checkParams error: " & DBSetQuery
            Exit Function
        End If

        ' second call (we're being set to dirty in calc event handler)
        If existsCalcCont(callID) Then
            If allCalcContainers(callID).errOccured Then
#If DEBUGME = 1 Then
              LogToEventViewer "Env:" & setEnv & ", DBSetQuery returned Error, NOT removing container with callID = " & callID, LogInf, 0
#End If
                ' commented this to prevent endless loops !!
                'allCalcContainers.Remove callID
                ' special case for invocations from function wizard
            ElseIf Not allCalcContainers(callID).working Then
                allCalcContainers.Remove(callID)
                makeCalcMsgContainer(callID, CStr(Query), theHostApp.caller, Nothing, targetRange, CStr(ConnString), Nothing, 0, False, False, False, False, 0, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, False)
            End If
        Else
            ' reset status messages when starting new query...
            If existsStatusCont(callID) Then allStatusContainers(callID).statusMsg = vbNullString
            ' add transportation info for event proc
            makeCalcMsgContainer(callID, CStr(Query), theHostApp.caller, Nothing, targetRange, CStr(ConnString), Nothing, 0, False, False, False, False, 0, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, False)
        End If
        If existsStatusCont(callID) Then
            DBSetQuery = "Env:" & setEnv & ", statusMsg: " & allStatusContainers(callID).statusMsg
        Else
            DBSetQuery = "Env:" & setEnv & ", no recalculation done for unchanged query..."
        End If
        theHostApp.EnableEvents = True
        Exit Function

DBSetQuery_Error:
        LogToEventViewer("Error (" & Err.Description & ") in Functions.DBSetQuery, callID : " & callID & ", in " & Erl(), EventLogEntryType.Error)
        DBSetQuery = "Env:" & setEnv & ", Error (" & Err.Description & ") in DBSetQuery, callID : " & callID
        theHostApp.EnableEvents = True
    End Function

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
    ''' <param name="subscribeTo">not yet implemented</param>
    ''' <returns>Status Message, data values are returned outside of function cell (@see DBFuncEventHandler)</returns>
    Public Function DBListFetch(Query As Object, ConnString As Object, Optional targetRange As Range = Nothing, Optional formulaRange As Object = Nothing, Optional extendDataArea As Integer = 0, Optional HeaderInfo As Boolean = False, Optional AutoFit As Boolean = False, Optional autoformat As Boolean = False, Optional ShowRowNums As Boolean = False, Optional subscribeTo As String = vbNullString) As String
        Dim callID As String
        Dim setEnv As String

        setEnv = fetchSetting("ConfigName", vbNullString)
        If dontCalcWhileClearing Then
            DBListFetch = "Env:" & setEnv & ", dontCalcWhileClearing = True !"
            Exit Function
        End If
        On Error GoTo DBListFetch_Error
        ' calcContainers are identified by wbname + Sheetname + function caller cell Address
        'If Not IsObject(theHostApp.caller) Then Exit Function
        callID = "[" & theHostApp.caller.Parent.Parent.name & "]" & theHostApp.caller.Parent.name & "!" & theHostApp.caller.Address
        If TypeName(ConnString) = "Error" Then ConnString = vbNullString
        If TypeName(ConnString) = "Range" Then ConnString = ConnString.Value
        ' in case of number as connection string, take the stored Connection string .. 1 usually prod, 2 .. usually test, 3.. development)
        If TypeName(ConnString) = "Integer" Or TypeName(ConnString) = "Long" Or TypeName(ConnString) = "Double" Or TypeName(ConnString) = "Short" Then
            setEnv = fetchSetting("ConfigName" & ConnString, vbNullString)
            ConnString = fetchSetting("ConstConnString" & ConnString, vbNullString)
        End If
        If TypeName(targetRange) <> "Range" Then
            DBListFetch = "Env:" & setEnv & ", Target range not given, is no valid range or range name doesn't exist!"
            Exit Function
        End If
        ' can't check for nothing with error handler enabled
        On Error Resume Next
        If Not formulaRange Is Nothing Then
            If TypeName(formulaRange) <> "Range" Then
                DBListFetch = "Env:" & setEnv & ", Invalid FormulaRange or range name doesn't exist!"
                Exit Function
            End If
        End If
        On Error GoTo DBListFetch_Error

        ' get target range name ...
        Dim functionArgs : functionArgs = functionSplit(theHostApp.caller.Formula, ",", """", "DBListFetch", "(", ")")
        Dim targetRangeName As String : targetRangeName = functionArgs(2)
        ' check if fetched argument targetRangeName is really a name or just a plain range address
        If Not existsNameInWb(targetRangeName, theHostApp.caller.Parent.Parent) And Not existsNameInSheet(targetRangeName, theHostApp.caller.Parent) Then targetRangeName = vbNullString

        ' get formula range name ...
        Dim formulaRangeName As String
        If UBound(functionArgs) > 2 Then
            formulaRangeName = functionArgs(3)
            If Not existsNameInWb(formulaRangeName, theHostApp.caller.Parent.Parent) And Not existsNameInSheet(formulaRangeName, theHostApp.caller.Parent) Then formulaRangeName = vbNullString
        Else
            formulaRangeName = vbNullString
        End If

        ' check query, also converts query to string (if it is a range)
        DBListFetch = checkParams(Query)
        ' error message is returned from checkParams, if OK then returns nothing
        If DBListFetch.Length > 0 Then
            DBListFetch = "Env:" & setEnv & ", " & DBListFetch
            Exit Function
        End If

        ' second call (we're being set to dirty in calc event handler)
        If existsCalcCont(callID) Then
            If allCalcContainers(callID).errOccured Then
#If DEBUGME = 1 Then
              LogToEventViewer "Env:" & setEnv & ", DBListQuery returned Error, removing container with callID = " & callID, LogInf, 0
#End If
                ' commented this to prevent endless loops !!
                'allCalcContainers.Remove callID
                ' special case for invocations from function wizard
            ElseIf Not allCalcContainers(callID).working Then
                allCalcContainers.Remove(callID)
                makeCalcMsgContainer(callID, CStr(Query), theHostApp.caller, Nothing, targetRange, CStr(ConnString), formulaRange, extendDataArea, HeaderInfo, AutoFit, autoformat, ShowRowNums, 0, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, targetRangeName, formulaRangeName, False)
            End If
        Else
            ' reset status messages when starting new query...
            If existsStatusCont(callID) Then allStatusContainers(callID).statusMsg = vbNullString
            ' add transportation info for event proc
            makeCalcMsgContainer(callID, CStr(Query), theHostApp.caller, Nothing, targetRange, CStr(ConnString), formulaRange, extendDataArea, HeaderInfo, AutoFit, autoformat, ShowRowNums, 0, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, targetRangeName, formulaRangeName, False)
        End If
        If existsStatusCont(callID) Then
            DBListFetch = "Env:" & setEnv & ", " & allStatusContainers(callID).statusMsg
        Else
            DBListFetch = "Env:" & setEnv & ", no recalculation done for unchanged query..."
        End If
        theHostApp.EnableEvents = True
        Exit Function

DBListFetch_Error:
        If VBDEBUG Then Debug.Print("Error (" & Err.Description & ") in DBListFetch, callID : " & callID) : Stop : Resume
        LogToEventViewer("Error (" & Err.Description & ") in Functions.DBListFetch, callID : " & callID & ", in " & Erl(), EventLogEntryType.Error)
        DBListFetch = "Env:" & setEnv & ", Error (" & Err.Description & ") in DBListFetch, callID : " & callID
        theHostApp.EnableEvents = True
    End Function

    ''' <summary>Fetches a row (single record) queried (defined in query) from DB (defined in ConnString) into targetArray</summary>
    ''' <param name="Query">query for getting data</param>
    ''' <param name="ConnString">connection string defining DB, user, etc...</param>
    ''' <param name="targetArray">Range to put the data into</param>
    ''' <returns>Status Message, data values are returned outside of function cell (@see DBFuncEventHandler)</returns>
    Public Function DBRowFetch(Query As Object, ConnString As Object, ParamArray targetArray() As Object) As String
        Dim tempArray() As Range = Nothing ' final target array that is passed to makeCalcMsgContainer (after removing header flag)
        Dim callID As String
        Dim HeaderInfo As Boolean
        Dim subscribeTo As String
        Dim setEnv As String

        On Error GoTo DBRowFetch_Error
        setEnv = fetchSetting("ConfigName", vbNullString)
        ' calcContainers are identified by wbname + Sheetname + function caller cell Address
        callID = "[" & theHostApp.caller.Parent.Parent.name & "]" & theHostApp.caller.Parent.name & "!" & theHostApp.caller.Address
        DBRowFetch = checkParams(Query)
        If DBRowFetch.Length > 0 Then
            DBRowFetch = "Env:" & setEnv & ", " & DBRowFetch
            Exit Function
        End If
        ' add transportation info for event proc
        If TypeName(ConnString) = "Error" Then ConnString = vbNullString
        If TypeName(ConnString) = "Range" Then ConnString = ConnString.Value
        ' in case of number as connection string, take the stored Connection string .. 1 usually prod, 2 .. usually test, 3.. development)
        If TypeName(ConnString) = "Integer" Or TypeName(ConnString) = "Long" Or TypeName(ConnString) = "Double" Or TypeName(ConnString) = "Short" Then
            setEnv = fetchSetting("ConfigName" & ConnString, vbNullString)
            ConnString = fetchSetting("ConstConnString" & ConnString, vbNullString)
        End If

        Dim i As Long
        If TypeName(targetArray(0)) = "Boolean" Then
            HeaderInfo = targetArray(0)
            If TypeName(targetArray(1)) = "String" Then
                subscribeTo = targetArray(1)
                For i = 2 To UBound(targetArray)
                    ReDim Preserve tempArray(i - 2)
                    tempArray(i - 2) = targetArray(i)
                Next
            Else
                For i = 1 To UBound(targetArray)
                    ReDim Preserve tempArray(i - 1)
                    tempArray(i - 1) = targetArray(i)
                Next
            End If
        ElseIf TypeName(targetArray(0)) = "String" Then
            subscribeTo = targetArray(0)
            If TypeName(targetArray(1)) = "Boolean" Then
                HeaderInfo = targetArray(1)
                For i = 2 To UBound(targetArray)
                    ReDim Preserve tempArray(i - 2)
                    tempArray(i - 2) = targetArray(i)
                Next
            Else
                For i = 1 To UBound(targetArray)
                    ReDim Preserve tempArray(i - 1)
                    tempArray(i - 1) = targetArray(i)
                Next
            End If
        ElseIf TypeName(targetArray(0)) = "Error" Then
            DBRowFetch = "Env:" & setEnv & ", Error: First argument empty or error !"
            Exit Function
        Else
            For i = 0 To UBound(targetArray)
                ReDim Preserve tempArray(i)
                tempArray(i) = targetArray(i)
            Next
        End If
        If existsCalcCont(callID) Then
            If allCalcContainers(callID).errOccured Then
#If DEBUGME = 1 Then
              LogToEventViewer "DBRowQuery returned Error, removing container with callID = " & callID, LogInf, 0
#End If
                ' commented this to prevent endless loops !!
                'allCalcContainers.Remove callID
                ' special case for intermediate invocation in function wizard
            ElseIf Not allCalcContainers(callID).working Then
                allCalcContainers.Remove(callID)
                makeCalcMsgContainer(callID, CStr(Query), theHostApp.caller, tempArray, Nothing, CStr(ConnString), Nothing, 0, HeaderInfo, False, False, False, 0, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, False)
            End If
        Else
            ' reset status messages when starting new query...
            If existsStatusCont(callID) Then allStatusContainers(callID).statusMsg = vbNullString
            ' add transportation info for event proc
            makeCalcMsgContainer(callID, CStr(Query), theHostApp.caller, tempArray, Nothing, CStr(ConnString), Nothing, 0, HeaderInfo, False, False, False, 0, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, False)
        End If
        If existsStatusCont(callID) Then DBRowFetch = "Env:" & setEnv & ", " & allStatusContainers(callID).statusMsg
        theHostApp.EnableEvents = True

        Exit Function
DBRowFetch_Error:
        If VBDEBUG Then Debug.Print("Error (" & Err.Description & ") in DBRowFetch, callID : " & callID) : Stop : Resume
        LogToEventViewer("Error (" & Err.Description & ") in DBRowFetch, callID : " & callID & ", in " & Erl(), EventLogEntryType.Error)
        DBRowFetch = "Env:" & setEnv & ", Error (" & Err.Description & ") in Functions.DBRowFetch, callID : " & callID
        theHostApp.EnableEvents = True
    End Function

    ''' <summary>Fetches results of a query (defined in query) from DB (optionally defined in ConnString)
    '''             into current cell. Columns are separated by colSeparator, rows by rowSeparator</summary>
    ''' <param name="Query">query for getting data</param>
    ''' <param name="ConnString">connection string defining DB, user, etc...</param>
    ''' <param name="HeaderInfo">additionally fill headings of query</param>
    ''' <param name="colSep">usual column separator</param>
    ''' <param name="rowSep">usual row separator</param>
    ''' <param name="lastColSep">if given, last column is separated from others with that one</param>
    ''' <param name="lastRowSep">if given, last row is separated from others with that one</param>
    ''' <param name="InterleaveHeader"></param>
    ''' <param name="subscribeTo"></param>
    ''' <returns>query result</returns>
    ''' <remarks>layout of value is as follows (carriage returns are here just for clarity of display, actually this is rowSep, resp. lastRowSep):
    ''' header1 (colSep) header2 (colSep) header3 (colSep)... (colSep/lastColSep) headerN (rowSep)
    ''' value11 (colSep) value12 (colSep) value13 (colSep)... (colSep/lastColSep) value1N (rowSep)
    ''' ...
    ''' value(M-1)1 (colSep) value(M-1)2 (colSep) value(M-1)3 (colSep)... (colSep/lastColSep) value(M-1)N (rowSep/lastRowSep)
    ''' valueM1 (colSep) valueM2 (colSep) valueM3 (colSep)... (colSep/lastColSep) valueMN
    '''</remarks>
    Public Function DBCellFetch(Query As Object, Optional ConnString As Object = vbNullString, Optional HeaderInfo As Boolean = False, Optional colSep As String = ",", Optional rowSep As String = vbLf, Optional lastColSep As String = vbNullString, Optional lastRowSep As String = vbNullString, Optional InterleaveHeader As Boolean = False, Optional subscribeTo As String = vbNullString) As String
        Dim callID As String
        Dim setEnv As String

        On Error GoTo DBCellFetch_Error
        setEnv = fetchSetting("ConfigName", vbNullString)
        ' calcContainers are identified by wbname + Sheetname + function caller cell Address
        callID = "[" & theHostApp.caller.Parent.Parent.name & "]" & theHostApp.caller.Parent.name & "!" & theHostApp.caller.Address
        ' check for errors in query parameter and convert query to string
        DBCellFetch = checkParams(Query)
        If DBCellFetch.Length > 0 Then
            DBCellFetch = "Env:" & setEnv & ", " & DBCellFetch
            Exit Function
        End If
        ' add transportation info for event proc
        If TypeName(ConnString) = "Error" Then ConnString = vbNullString
        If TypeName(ConnString) = "Range" Then ConnString = ConnString.Value
        ' in case of number as connection string, take the stored Connection string .. 1 usually prod, 2 .. usually test, 3.. development)
        If TypeName(ConnString) = "Integer" Or TypeName(ConnString) = "Long" Or TypeName(ConnString) = "Double" Or TypeName(ConnString) = "Short" Then
            setEnv = fetchSetting("ConfigName" & ConnString, vbNullString)
            ConnString = fetchSetting("ConstConnString" & ConnString, vbNullString)
        End If
        If existsCalcCont(callID) Then
            If Not allCalcContainers(callID).working Then
                allCalcContainers.Remove(callID)
                makeCalcMsgContainer(callID, CStr(Query), theHostApp.caller, Nothing, Nothing, CStr(ConnString), Nothing, 0, HeaderInfo, False, False, False, 0, vbNullString, vbNullString, vbNullString, colSep, rowSep, lastColSep, lastRowSep, vbNullString, vbNullString, InterleaveHeader)
            End If
        Else
            ' reset status messages when starting new query...
            If existsStatusCont(callID) Then allStatusContainers(callID).statusMsg = vbNullString
            ' add transportation info for event proc
            makeCalcMsgContainer(callID, CStr(Query), theHostApp.caller, Nothing, Nothing, CStr(ConnString), Nothing, 0, HeaderInfo, False, False, False, 0, vbNullString, vbNullString, vbNullString, colSep, rowSep, lastColSep, lastRowSep, vbNullString, vbNullString, InterleaveHeader)
        End If
        If existsStatusCont(callID) Then DBCellFetch = "Env:" & setEnv & ", " & allStatusContainers(callID).statusMsg
        theHostApp.EnableEvents = True
        Exit Function

DBCellFetch_Error:
        If VBDEBUG Then Debug.Print("Error (" & Err.Description & ") in DBCellFetch") : Stop : Resume
        LogToEventViewer("Error (" & Err.Description & ") in Functions.DBCellFetch, in " & Erl(), EventLogEntryType.Error)
        DBCellFetch = "Env:" & setEnv & ", Error (" & Err.Description & ") in DBCellFetch"
        theHostApp.EnableEvents = True
    End Function

    Public Function DBAddinEnvironment() As String
        theHostApp.Volatile
        DBAddinEnvironment = fetchSetting("ConfigName", vbNullString)
        If theHostApp.Calculation = XlCalculation.xlCalculationManual Then DBAddinEnvironment = "calc Mode is manual, please press F9 to get current DBAddin environment !"
    End Function

    Public Function DBAddinServerSetting() As String
        Dim keywordstart As Integer
        Dim theConnString As String

        theHostApp.Volatile
        On Error Resume Next
        theConnString = fetchSetting("ConstConnString", vbNullString)
        keywordstart = InStr(1, theConnString, "Server=") + Len("Server=")
        DBAddinServerSetting = Mid$(theConnString, keywordstart, InStr(keywordstart, theConnString, ";") - keywordstart)
        If theHostApp.Calculation = XlCalculation.xlCalculationManual Then DBAddinServerSetting = "calc Mode is manual, please press F9 to get current DBAddin server setting !"
        If Err.Number <> 0 Then DBAddinServerSetting = "Error happened: " & Err.Description
    End Function

    ''' <summary>checks query and calculation mode if OK for both DBListFetch and DBRowFetch function</summary>
    ''' <param name="Query"></param>
    ''' <returns>Error String (empty if OK)</returns>
    Private Function checkParams(ByRef Query As Object) As String
        Dim errval As String, AddInfo As String = vbNullString

        checkParams = vbNullString
        If theHostApp.Calculation = XlCalculation.xlCalculationManual Then
            checkParams = "calc Mode is manual, please press F9 to trigger data fetching !"
        Else
            If IsError(Query) Then
                Select Case Query
                    Case ExcelError.ExcelErrorDiv0 : errval = "#DIV/0!"
                    Case ExcelError.ExcelErrorNA : errval = "#N/A"
                    Case ExcelError.ExcelErrorName : errval = "#NAME?"
                    Case ExcelError.ExcelErrorNull : errval = "#NULL!"
                    Case ExcelError.ExcelErrorNum : errval = "#NUM!"
                    Case ExcelError.ExcelErrorRef : errval = "#REF!"
                    Case ExcelError.ExcelErrorValue : errval = "#VALUE!" : AddInfo = "(in case query is inside DBfunc, check if it's > 255 chars)"
                    Case Else : errval = "This should never happen!!"
                End Select
                checkParams = "query contains: '" & errval & "' " & AddInfo
            ElseIf TypeName(Query) = "Range" Then
                ' if query is range then get the query string out of it..
                Query = concatCellsSep(vbLf, Query)
                If TypeName(Query) <> "String" Then checkParams = "query parameter invalid (not a string) !"
                If Query.ToString.Length = 0 Then checkParams = "empty query provided !"
            ElseIf TypeName(Query) = "String" Then
                If Query.ToString.Length = 0 Then checkParams = "empty query provided !"
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
    ''' <param name="controlType"></param>
    ''' <param name="dataTargetRange"></param>
    ''' <param name="controlLocation"></param>
    ''' <param name="ControlName"></param>
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
    Private Sub makeCalcMsgContainer(ByRef callID As String, ByRef Query As String, appCaller As Object, targetArray As Object, ByRef targetRange As Range, ByRef ConnString As String, ByRef formulaRange As Object, ByRef extendArea As Integer, ByRef HeaderInfo As Boolean, ByRef AutoFit As Boolean, ByRef autoformat As Boolean, ByRef ShowRowNumbers As Boolean, ByRef controlType As Integer, ByRef dataTargetRange As String, ByRef controlLocation As String, ByRef ControlName As String, ByRef colSep As String, ByRef rowSep As String, ByRef lastColSep As String, ByRef lastRowSep As String, ByRef targetRangeName As String, ByRef formulaRangeName As String, ByRef InterleaveHeader As Boolean)
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
        myCalcCont.controlType = controlType
        myCalcCont.dataTargetRange = dataTargetRange
        myCalcCont.controlLocation = controlLocation
        myCalcCont.ControlName = ControlName
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
#If DEBUGME = 1 Then
      LogToEventViewer "Leaving makeCalcMsgContainer, caller: " & callID, LogInf, 0
#End If

        Exit Sub
makeCalcMsgContainer_Error:
        If Err.Number <> 457 Then
            If VBDEBUG Then Debug.Print("Error (" & Err.Description & ") in makeCalcMsgContainer, callID: " & callID) : Stop : Resume
            LogToEventViewer("Error (" & Err.Description & ") in Functions.makeCalcMsgContainer, callID: " & callID & ", in " & Erl(), EventLogEntryType.Error)
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

    ''' <summary> maintenance procedure to purge names used for dbfunctions from workbook</summary>
    Public Sub purge()
        On Error GoTo err1
        Dim DBname As Name
        For Each DBname In theHostApp.ActiveWorkbook.Names
            If DBname.Name Like "*ExterneDaten*" Then
                DBname.Delete()
            ElseIf DBname.Name Like "DBListArea*" Then
                DBname.Delete()
            ElseIf DBname.Name Like "DBFtarget*" Then
                DBname.Delete()
            ElseIf DBname.Name Like "DBFsource*" Then
                DBname.Delete()
            ElseIf InStr(1, DBname.RefersTo, "#REF!") > 0 Then
                DBname.Delete()
            End If
        Next
        Exit Sub
err1:
        LogError("purge error: " & Err.Description & ", in " & Erl(), , , 1)
    End Sub
End Module
