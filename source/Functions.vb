Imports ADODB
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports System.Collections.Generic
Imports System.Linq


''' <summary>Provides a data structure for transporting information back from the calculation action procedure to the calling function</summary>
Public Class ContainedStatusMsg
    ''' <summary>any status msg used for displaying in the result of function</summary>
    Public statusMsg As String
End Class

''' <summary>Contains the public callable DB functions and helper functions</summary>
Public Module Functions
    ' Global objects/variables for DBFuncs
    ''' <summary>global collection of information transport containers between function and calc event procedure</summary>
    Public StatusCollection As Dictionary(Of String, ContainedStatusMsg)
    ''' <summary>connection object: always use the same, if possible (same ConnectionString)</summary>
    Public conn As ADODB.Connection
    ''' <summary>connection string can be changed for calls with different connection strings</summary>
    Public CurrConnString As String
    ''' <summary>query cache for avoiding unnecessary recalculations/data retrievals by volatile inputs to DB Functions (now(), etc.)</summary>
    Public queryCache As Dictionary(Of String, String)
    ''' <summary>prevent multiple connection retries for each function in case of error</summary>
    Public dontTryConnection As Boolean
    ''' <summary>avoid entering dblistfetch/dbrowfetch functions during clearing of listfetch areas (before saving)</summary>
    Public dontCalcWhileClearing As Boolean

    ''' <summary>Create database compliant date, time or datetime string from excel datetype value</summary>
    ''' <param name="DatePart">date/time/datetime single parameter or range reference</param>
    ''' <param name="formatting">formatting instruction for Date format, see remarks</param>
    ''' <returns>the DB compliant formatted date/time/datetime</returns>
    ''' <remarks>
    ''' formatting = 0: A simple datestring (format 'YYYYMMDD'), datetime values are converted to 'YYYYMMDD HH:MM:SS' and time values are converted to 'HH:MM:SS'.
    ''' formatting = 1: An ANSI compliant Date string (format date 'YYYY-MM-DD'), datetime values are converted to timestamp 'YYYY-MM-DD HH:MM:SS' and time values are converted to time time 'HH:MM:SS'.
    ''' formatting = 2: An ODBC compliant Date string (format {d 'YYYY-MM-DD'}), datetime values are converted to {ts 'YYYY-MM-DD HH:MM:SS'} and time values are converted to {t 'HH:MM:SS'}.
    ''' formatting = 3: An Access/JetDB compliant Date string (format #YYYY-MM-DD#), datetime values are converted to #YYYY-MM-DD HH:MM:SS# and time values are converted to #HH:MM:SS#.
    ''' add 10 to formatting to include fractions of a second (1000) 
    ''' formatting >13 or empty (99=default value): take the formatting option from setting DefaultDBDateFormatting (0 if not given)
    ''' </remarks>
    <ExcelFunction(Description:="Create database compliant date, time or datetime string from excel datetype value")>
    Public Function DBDate(<ExcelArgument(Description:="date/time/datetime")> ByVal DatePart As Object,
                           <ExcelArgument(Description:="formatting option, 0:'YYYYMMDD', 1:'YYYY-MM-DD'), 2:{d 'YYYY-MM-DD'},3:Access/JetDB #DD/MM/YYYY#, add 10 to formatting to include fractions of a second (1000)")> Optional formatting As Integer = 99) As String
        DBDate = ""
        Try
            If formatting > 3 Then formatting = DefaultDBDateFormatting
            If TypeName(DatePart) = "Object(,)" Then
                For Each myCell In DatePart
                    If TypeName(myCell) = "ExcelEmpty" Then
                        ' do nothing here
                    Else
                        DBDate += formatDBDate(myCell, formatting) + ","
                    End If
                Next
                ' cut last comma
                If DBDate.Length > 0 Then DBDate = Left(DBDate, Len(DBDate) - 1)
            Else
                ' direct value in DBDate..
                If TypeName(DatePart) = "ExcelEmpty" Or TypeName(DatePart) = "ExcelMissing" Then
                    ' do nothing here
                Else
                    DBDate = formatDBDate(DatePart, formatting)
                End If
            End If
        Catch ex As Exception
            LogWarn(ex.Message)
            DBDate = "Error (" + ex.Message + ") in function DBDate"
        End Try
    End Function

    ''' <summary>takes an OADate and formats it as a DB Compliant string, using formatting as formatting instruction</summary>
    ''' <param name="datVal">OADate (double) date parameter</param>
    ''' <param name="formatting">formatting flag (see DBDate for details)</param>
    ''' <returns>formatted Date string</returns>
    Private Function formatDBDate(datVal As Double, formatting As Integer) As String
        formatDBDate = ""
        If Int(datVal) = datVal Then
            If formatting = 0 Then
                formatDBDate = "'" + Format(Date.FromOADate(datVal), "yyyyMMdd") + "'"
            ElseIf formatting = 1 Then
                formatDBDate = "DATE '" + Format(Date.FromOADate(datVal), "yyyy-MM-dd") + "'"
            ElseIf formatting = 2 Then
                formatDBDate = "{d '" + Format(Date.FromOADate(datVal), "yyyy-MM-dd") + "'}"
            ElseIf formatting = 3 Then
                formatDBDate = "#" + Format(Date.FromOADate(datVal), "yyyy-MM-dd") + "#"
            End If
        ElseIf CInt(datVal) > 1 Then
            If formatting = 0 Then
                formatDBDate = "'" + Format(Date.FromOADate(datVal), "yyyyMMdd HH:mm:ss") + "'"
            ElseIf formatting = 1 Then
                formatDBDate = "timestamp '" + Format(Date.FromOADate(datVal), "yyyy-MM-dd HH:mm:ss") + "'"
            ElseIf formatting = 2 Then
                formatDBDate = "{ts '" + Format(Date.FromOADate(datVal), "yyyy-MM-dd HH:mm:ss") + "'}"
            ElseIf formatting = 3 Then
                formatDBDate = "#" + Format(Date.FromOADate(datVal), "yyyy-MM-dd HH:mm:ss") + "#"
            ElseIf formatting = 10 Then
                formatDBDate = "'" + Format(Date.FromOADate(datVal), "yyyyMMdd HH:mm:ss.fff") + "'"
            ElseIf formatting = 11 Then
                formatDBDate = "timestamp '" + Format(Date.FromOADate(datVal), "yyyy-MM-dd HH:mm:ss.fff") + "'"
            ElseIf formatting = 12 Then
                formatDBDate = "{ts '" + Format(Date.FromOADate(datVal), "yyyy-MM-dd HH:mm:ss.fff") + "'}"
            ElseIf formatting = 13 Then
                formatDBDate = "#" + Format(Date.FromOADate(datVal), "yyyy-MM-dd HH:mm:ss.fff") + "#"
            End If
        Else
            If formatting = 0 Then
                formatDBDate = "'" + Format(Date.FromOADate(datVal), "HH:mm:ss") + "'"
            ElseIf formatting = 1 Then
                formatDBDate = "time '" + Format(Date.FromOADate(datVal), "HH:mm:ss") + "'"
            ElseIf formatting = 2 Then
                formatDBDate = "{t '" + Format(Date.FromOADate(datVal), "HH:mm:ss") + "'}"
            ElseIf formatting = 3 Then
                formatDBDate = "#" + Format(Date.FromOADate(datVal), "HH:mm:ss") + "#"
            ElseIf formatting = 10 Then
                formatDBDate = "'" + Format(Date.FromOADate(datVal), "HH:mm:ss.fff") + "'"
            ElseIf formatting = 11 Then
                formatDBDate = "time '" + Format(Date.FromOADate(datVal), "HH:mm:ss.fff") + "'"
            ElseIf formatting = 12 Then
                formatDBDate = "{t '" + Format(Date.FromOADate(datVal), "HH:mm:ss.fff") + "'}"
            ElseIf formatting = 13 Then
                formatDBDate = "#" + Format(Date.FromOADate(datVal), "HH:mm:ss.fff") + "#"
            End If
        End If
    End Function

    ''' <summary>Create a database compliant string from cell values, potentially concatenating with other parts for easy inclusion of wildcards (%,_)</summary>
    ''' <param name="StringPart">array of strings/wildcards or ranges containing strings/wildcards</param>
    ''' <returns>database compliant string</returns>
    <ExcelFunction(Description:="Create a database compliant string from cell values, potentially concatenating with other parts for easy inclusion of wildcards (%,_)")>
    Public Function DBString(<ExcelArgument(Description:="array of strings/wildcards or ranges containing strings/wildcards")> ParamArray StringPart() As Object) As String
        Dim myRef, myCell
        Try
            Dim retval As String = ""
            For Each myRef In StringPart
                If TypeName(myRef) = "Object(,)" Then
                    For Each myCell In myRef
                        If TypeName(myCell) = "ExcelEmpty" Then
                            ' do nothing here
                        Else
                            retval += myCell.ToString()
                        End If
                    Next
                ElseIf TypeName(myRef) = "ExcelEmpty" Or TypeName(myRef) = "ExcelMissing" Then
                    ' do nothing here
                Else
                    retval += myRef.ToString()
                End If
            Next
            DBString = "'" + retval + "'"
        Catch ex As Exception
            LogWarn(ex.Message)
            DBString = "Error (" + ex.Message + ") in DBString"
        End Try
    End Function

    ''' <summary>Create an in clause from cell values, strings are created with quotation marks,
    '''             dates are created with DBDate</summary>
    ''' <param name="inClausePart">array of values or ranges containing values</param>
    ''' <returns>database compliant in-clause string</returns>
    <ExcelFunction(Description:="Create an in clause from cell values, strings are created with quotation marks")>
    Public Function DBinClause(<ExcelArgument(AllowReference:=True, Description:="array of values or ranges containing values")> ParamArray inClausePart As Object()) As String
        Dim concatResult As String = DoConcatCellsSep(",", True, False, inClausePart)
        DBinClause = If(Left(concatResult, 5) = "Error", concatResult, "in (" + concatResult + ")")
    End Function

    ''' <summary>Create an in clause from cell values, strings are created with quotation marks,
    '''             dates are created with DBDate</summary>
    ''' <param name="inClausePart">array of values or ranges containing values</param>
    ''' <returns>database compliant in-clause string</returns>
    <ExcelFunction(Description:="Create an in clause from cell values, all arguments are treated as strings (and will be created with quotation marks)")>
    Public Function DBinClauseStr(<ExcelArgument(AllowReference:=True, Description:="array of values or ranges containing values")> ParamArray inClausePart As Object()) As String
        Dim concatResult As String = DoConcatCellsSep(",", True, True, inClausePart)
        DBinClauseStr = If(Left(concatResult, 5) = "Error", concatResult, "in (" + concatResult + ")")
    End Function

    ''' <summary>concatenates values contained in thetarget together (using .value attribute for cells)</summary>
    ''' <param name="concatPart">all cells/values which should be concatenated</param>
    ''' <returns>concatenated String</returns>
    <ExcelFunction(Description:="concatenates values contained in thetarget together (using .value attribute for cells)")>
    Public Function concatCells(<ExcelArgument(AllowReference:=True, Description:="all cells/values which should be concatenated")> ParamArray concatPart As Object()) As String
        concatCells = DoConcatCellsSep("", False, False, concatPart)
    End Function

    ''' <summary>concatenates values contained in thetarget (using .value for cells) using a separator</summary>
    ''' <param name="separator">the separator</param>
    ''' <param name="concatPart">all cells/values which should be concatenated</param>
    ''' <returns>concatenated String</returns>
    <ExcelFunction(Description:="concatenates values contained in thetarget (using .value for cells) using a separator")>
    Public Function concatCellsSep(<ExcelArgument(AllowReference:=True, Description:="the separator")> separator As String,
                                   <ExcelArgument(AllowReference:=True, Description:="all cells/values which should be concatenated")> ParamArray concatPart As Object()) As String
        concatCellsSep = DoConcatCellsSep(separator, False, False, concatPart)
    End Function

    ''' <summary>chains values contained in thetarget together with commas, mainly used for creating select header</summary>
    ''' <param name="chainPart">range where values should be chained</param>
    ''' <returns>chained String</returns>
    <ExcelFunction(Description:="chains values contained in thetarget together with commas, mainly used for creating select header")>
    Public Function chainCells(<ExcelArgument(AllowReference:=True, Description:="range where values should be chained")> ParamArray chainPart As Object()) As String
        chainCells = DoConcatCellsSep(",", False, False, chainPart)
    End Function

    ''' <summary>private function that actually concatenates values contained in Object array myRange together (either using .text or .value for cells in myrange) using a separator</summary>
    ''' <param name="separator">the separator-string that is filled between values</param>
    ''' <param name="DBcompliant">should a potential string or date part be formatted database compliant (surrounded by quotes)?</param>
    ''' <param name="concatParts">Object array, whose values should be concatenated</param>
    ''' <returns>concatenated String</returns>
    Private Function DoConcatCellsSep(separator As String, DBcompliant As Boolean, OnlyString As Boolean, ParamArray concatParts As Object()) As String
        Dim myRef, myCell

        Try
            Dim retval As String = ""
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
                            If OnlyString Then
                                retval = retval + separator + IIf(DBcompliant, "'", "") + Convert.ToString(myCell, System.Globalization.CultureInfo.InvariantCulture) + IIf(DBcompliant, "'", "")
                            Else
                                retval = retval + separator + Convert.ToString(myCell, System.Globalization.CultureInfo.InvariantCulture)
                            End If
                        Else
                            ' avoid double quoting if passed string is already quoted (by using DBDate or DBString as input to this) and DBcompliant quoting is requested
                            retval = retval + separator + IIf(DBcompliant And Left(myCell, 1) <> "'", "'", "") + myCell + IIf(DBcompliant And Right(myCell, 1) <> "'", "'", "")
                        End If
                    Next
                Else
                    ' and other direct values in formulas..
                    If TypeName(myRef) = "ExcelEmpty" Or TypeName(myRef) = "ExcelMissing" Then
                        ' do nothing here
                    ElseIf IsNumeric(myRef) And Not OnlyString Then ' no separate Date type for direct formula values
                        If OnlyString Then
                            retval = retval + separator + IIf(DBcompliant, "'", "") + Convert.ToString(myRef, System.Globalization.CultureInfo.InvariantCulture) + IIf(DBcompliant, "'", "")
                        Else
                            retval = retval + separator + Convert.ToString(myRef, System.Globalization.CultureInfo.InvariantCulture)
                        End If
                    Else
                        ' avoid double quoting if passed string is already quoted (by using DBDate or DBString as input to this) and DBcompliant quoting is requested
                        retval = retval + separator + IIf(DBcompliant And Left(myRef, 1) <> "'", "'", "") + myRef + IIf(DBcompliant And Right(myRef, 1) <> "'", "'", "")
                    End If
                End If
            Next
            DoConcatCellsSep = Mid$(retval, Len(separator) + 1) ' skip first separator
        Catch ex As Exception
            LogWarn(ex.Message)
            DoConcatCellsSep = "Error (" + ex.Message + ") in DoConcatCellsSep"
        End Try
    End Function

    ''' <summary>Stores a query into an Object defined in targetRange (an embedded MS Query/Listobject, Pivot table, etc.)</summary>
    ''' <param name="Query">query for getting data</param>
    ''' <param name="ConnString">connection string defining DB, user, etc...</param>
    ''' <param name="targetRange">Range with Object beneath to put the Query/ConnString into</param>
    ''' <returns>Status Message</returns>
    <ExcelFunction(Description:="Stores a query into an Object (embedded Listobject or Pivot table) defined in targetRange")>
    Public Function DBSetQuery(<ExcelArgument(Description:="query for getting data")> Query As Object,
                               <ExcelArgument(Description:="connection string defining DB, user, etc...")> ConnString As Object,
                               <ExcelArgument(Description:="Range with embedded Object to put the Query/ConnString into", AllowReference:=True)> targetRange As Object) As String
        Dim callID As String = ""
        Dim caller As Excel.Range
        Dim EnvPrefix As String = ""
        If ExcelDnaUtil.IsInFunctionWizard() Then Return "invoked from function wizard..."
        Try
            DBSetQuery = checkQueryAndTarget(Query, targetRange)
            If DBSetQuery.Length > 0 Then Exit Function
            caller = ToRange(XlCall.Excel(XlCall.xlfCaller))
            resolveConnstring(ConnString, EnvPrefix, True)
            ' calcContainers are identified by wbname + Sheetname + function caller cell Address
            callID = "[" + caller.Parent.Parent.Name + "]" + caller.Parent.Name + "!" + caller.Address
            LogInfo("entering function, callID: " + callID)
            ' check query, also converts query to string (if it is a range)
            ' error message or cached status message is returned from checkParamsAndCache, if query OK and result was not already calculated (cached) then empty string
            DBSetQuery = checkParamsAndCache(Query, callID, ConnString)
            If DBSetQuery.Length > 0 Then
                DBSetQuery = EnvPrefix + ", " + DBSetQuery
                Exit Function
            End If

            ' needed for check whether target range is actually a table Listobject reference
            Dim functionArgs = functionSplit(caller.Formula, ",", """", "DBSetQuery", "(", ")")
            Dim targetRangeName As String = functionArgs(2)
            If UBound(functionArgs) = 3 Then targetRangeName += "," + functionArgs(3)

            ' first call: actually perform query
            If Not StatusCollection.ContainsKey(callID) Then
                Dim statusCont As ContainedStatusMsg = New ContainedStatusMsg
                StatusCollection.Add(callID, statusCont)
                StatusCollection(callID).statusMsg = "" ' need this to prevent object not set errors in checkCache
                ExcelAsyncUtil.QueueAsMacro(Sub()
                                                DBSetQueryAction(callID, Query, targetRange, ConnString, caller, targetRangeName)
                                            End Sub)
            End If

        Catch ex As Exception
            LogWarn(ex.Message + ", callID: " + callID)
            DBSetQuery = EnvPrefix + ", Error (" + ex.Message + ") in DBSetQuery, callID: " + callID
        End Try
        LogInfo("leaving function, callID: " + callID)
    End Function

    ''' <summary>set Query parameters (query text and connection string) of Query List or pivot table (incl. chart)</summary>
    ''' <param name="callID">the key for the statusMsg container</param>
    ''' <param name="Query"></param>
    ''' <param name="targetRange"></param>
    ''' <param name="ConnString"></param>
    ''' <param name="caller"></param>
    Sub DBSetQueryAction(callID As String, Query As String, targetRange As Object, ConnString As String, caller As Excel.Range, targetRangeName As String)
        Dim TargetCell As Excel.Range
        Dim targetSH As Excel.Worksheet
        Dim targetWB As Excel.Workbook
        Dim thePivotTable As Excel.PivotTable = Nothing
        Dim theListObject As Excel.ListObject = Nothing

        Dim calcMode = ExcelDnaUtil.Application.Calculation
        Try
            ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual
        Catch ex As Exception : End Try
        ' this works around the data validation input bug and being called when COM Model is not ready
        ' when selecting a value from a list of a validated field or being invoked from a hyperlink (e.g. word), excel won't react to
        ' Application.Calculation changes, so just leave here...
        If ExcelDnaUtil.Application.Calculation <> Excel.XlCalculation.xlCalculationManual Then
            LogWarn("Error in setting Application.Calculation to Manual in query: " + Query + ", caller: " + callID)
            StatusCollection(callID).statusMsg = "Error in setting Application.Calculation to Manual in query: " + Query
            caller.Formula += " " ' trigger recalculation to return error message to calling function
            Exit Sub
        End If
        ' when being called from DBSequence.doDBModif, targetRange is an Excel.Range, otherwise it's a reference
        If TypeName(targetRange) = "ExcelReference" Then
            TargetCell = ToRange(targetRange)
        Else
            TargetCell = targetRange
        End If

        targetSH = TargetCell.Parent
        targetWB = TargetCell.Parent.Parent
        Dim callerFormula As String = caller.Formula.ToString
        Dim srcExtent As String = ""
        Dim errHappened As Boolean = False
        Try
            srcExtent = caller.Name.Name
        Catch ex As Exception
            errHappened = True
        End Try
        If errHappened Or InStr(1, srcExtent, "DBFsource") = 0 Then
            srcExtent = "DBFsource" + Replace(Guid.NewGuid().ToString, "-", "")
            Try
                caller.Name = srcExtent
                caller.Parent.Parent.Names(srcExtent).Visible = False
            Catch ex As Exception
                Throw New Exception("Error in setting srcExtent name: " + ex.Message + " in query: " + Query)
            End Try
        End If

        ' try to get either a pivot table object or a list object from the target cell. What we have, is checked later...
        Try : thePivotTable = TargetCell.PivotTable : Catch ex As Exception : End Try
        Try : theListObject = TargetCell.ListObject : Catch ex As Exception : End Try

        Dim connType As String = ""
        Dim bgQuery As Boolean
        DBModifs.preventChangeWhileFetching = True
        Dim targetExtent As String = Replace(srcExtent, "DBFsource", "DBFtarget")
        StatusCollection(callID).statusMsg = ""
        Try
            ' first, get the connection type from the underlying PivotCache or QueryTable (OLEDB or ODBC)
            If Not thePivotTable Is Nothing Then
                Try
                    connType = Left$(thePivotTable.PivotCache.Connection.ToString, InStr(1, thePivotTable.PivotCache.Connection.ToString, ";"))
                Catch ex As Exception
                    Throw New Exception("couldn't get connection from Pivot Table, please create Pivot Table with external data source !")
                End Try
            End If
            If Not theListObject Is Nothing Then
                Try
                    connType = Left$(theListObject.QueryTable.Connection.ToString, InStr(1, theListObject.QueryTable.Connection.ToString, ";"))
                Catch ex As Exception
                    Throw New Exception("couldn't get connection from ListObject, please create ListObject with external data source !")
                End Try
            End If

            ' if we haven't already set the connection type in the alternative connection string then set it now..
            If Left(ConnString, 6) <> "OLEDB;" And Left(ConnString, 5) <> "ODBC;" Then ConnString = connType + ConnString

            ' now set the connection string and the query and refresh it.
            If Not thePivotTable Is Nothing Then
                bgQuery = thePivotTable.PivotCache.BackgroundQuery
                thePivotTable.PivotCache.Connection = ConnString
                thePivotTable.PivotCache.CommandType = Excel.XlCmdType.xlCmdSql
                thePivotTable.PivotCache.CommandText = Query
                thePivotTable.PivotCache.BackgroundQuery = False
                thePivotTable.PivotCache.Refresh()
                StatusCollection(callID).statusMsg = "Set " + connType + " PivotTable to (bgQuery= " + bgQuery.ToString + "): " + Query
                thePivotTable.PivotCache.BackgroundQuery = bgQuery
                ' give hidden name to target range of pivot query (jump function)
                thePivotTable.TableRange1.Name = targetExtent
                thePivotTable.TableRange1.Parent.Parent.Names(targetExtent).Visible = False
            End If
            If Not theListObject Is Nothing Then
                bgQuery = theListObject.QueryTable.BackgroundQuery
                ' check whether target range is actually a table Listobject reference, if so, replace with simple address as this doesn't produce a #REF! error on QueryTable.Refresh
                ' this simple address is below being set to caller.Formula
                If InStr(targetRangeName, theListObject.Name) > 0 Then callerFormula = Replace(callerFormula, targetRangeName, Replace(TargetCell.Cells(1, 1).Address, "$", ""))
                ' in case list object is sorted externally, give a warning (otherwise this leads to confusion when trying to order in the query)...
                If theListObject.Sort.SortFields.Count > 0 Then MsgBox("List Object " + theListObject.Name + " set by DBSetQuery in " + callID + " is already sorted by Excel, ordering statements in the query don't have any effect !", MsgBoxStyle.Exclamation)
                ' in case of CUDFlags, reset them now (before resizing)...
                Dim dbMapperRangeName As String = getDBModifNameFromRange(TargetCell)
                If Left(dbMapperRangeName, 8) = "DBMapper" Then
                    Dim dbMapper As DBMapper = Globals.DBModifDefColl("DBMapper").Item(dbMapperRangeName)
                    dbMapper.resetCUDFlags()
                End If
                theListObject.QueryTable.Connection = ConnString
                theListObject.QueryTable.CommandType = Excel.XlCmdType.xlCmdSql
                theListObject.QueryTable.CommandText = Query
                theListObject.QueryTable.BackgroundQuery = False
                Try
                    theListObject.QueryTable.Refresh()
                Catch ex As Exception
                    Throw New Exception("Error in query table refresh: " + ex.Message)
                End Try
                StatusCollection(callID).statusMsg = "Set " + connType + " ListObject to (bgQuery= " + bgQuery.ToString + ", " + If(theListObject.QueryTable.FetchedRowOverflow, "Too many rows fetched to display !", "") + "): " + Query
                theListObject.QueryTable.BackgroundQuery = bgQuery
                Try
                    Dim testTarget = TargetCell.Address
                Catch ex As Exception
                    caller.Formula = callerFormula ' restore formula as excel deletes target range when changing query fundamentally
                End Try
                ' give hidden name to target range of listobject (jump function)
                Dim oldRange As Excel.Range = Nothing
                ' first invocation of DBSetQuery will have no defined targetExtent Range name, so this will fail:
                Try : oldRange = theListObject.Range.Parent.Parent.Names(targetExtent).RefersToRange : Catch ex As Exception : End Try
                If oldRange Is Nothing Then oldRange = theListObject.Range
                theListObject.Range.Name = targetExtent
                theListObject.Range.Parent.Parent.Names(targetExtent).Visible = False
                ' if refreshed range is a DBMapper and it is in the current workbook, resize it, but ONLY if it the DBMapper is the same area as the old range
                DBModifs.resizeDBMapperRange(theListObject.Range, oldRange)
            End If
        Catch ex As Exception
            LogWarn(ex.Message + " in query: " + Query + ", caller: " + callID)
            StatusCollection(callID).statusMsg = ex.Message + " in query: " + Query
        End Try

        ' neither PivotTable or ListObject could be found in TargetCell
        If StatusCollection(callID).statusMsg = "" Then
            StatusCollection(callID).statusMsg = "No PivotTable or ListObject with external data connection could be found in TargetRange " + TargetCell.Address
        End If
        DBModifs.preventChangeWhileFetching = False
        ExcelDnaUtil.Application.Calculation = calcMode
        caller.Formula += " " ' trigger recalculation to return below error message to calling function
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
                                <ExcelArgument(Description:="should headers be included in list")> Optional HeaderInfo As Object = Nothing,
                                <ExcelArgument(Description:="should columns be autofitted ?")> Optional AutoFit As Object = Nothing,
                                <ExcelArgument(Description:="should 1st row formats be autofilled down?")> Optional autoformat As Object = Nothing,
                                <ExcelArgument(Description:="should row numbers be displayed in 1st column?")> Optional ShowRowNums As Object = Nothing) As String
        Dim callID As String = ""
        Dim EnvPrefix As String = ""
        If ExcelDnaUtil.IsInFunctionWizard() Then Return "invoked from function wizard..."
        Try
            DBListFetch = checkQueryAndTarget(Query, targetRange)
            If DBListFetch.Length > 0 Then Exit Function
            Dim caller As Excel.Range = ToRange(XlCall.Excel(XlCall.xlfCaller))
            resolveConnstring(ConnString, EnvPrefix, False)
            ' calcContainers are identified by wbname + Sheetname + function caller cell Address
            callID = "[" + caller.Parent.Parent.Name + "]" + caller.Parent.Name + "!" + caller.Address
            LogInfo("entering function, callID: " + callID)
            ' prepare information for action proc
            If dontCalcWhileClearing Then
                DBListFetch = EnvPrefix + ", dontCalcWhileClearing = True !"
                Exit Function
            End If
            If TypeName(targetRange) <> "ExcelReference" Then
                DBListFetch = EnvPrefix + ", Invalid targetRange or range name doesn't exist!"
                Exit Function
            End If
            If TypeName(formulaRange) <> "ExcelMissing" And TypeName(formulaRange) <> "ExcelReference" Then
                DBListFetch = EnvPrefix + ", Invalid FormulaRange or range name doesn't exist!"
                Exit Function
            End If

            ' check query, also converts query to string (if it is a range)
            ' error message or cached status message is returned from checkParamsAndCache, if query OK and result was not already calculated (cached) then empty string
            DBListFetch = checkParamsAndCache(Query, callID, ConnString)
            If DBListFetch.Length > 0 Then
                DBListFetch = EnvPrefix + ", " + DBListFetch
                Exit Function
            End If

            ' get target range name ...
            Dim functionArgs = functionSplit(caller.Formula, ",", """", "DBListFetch", "(", ")")
            Dim targetRangeName As String : targetRangeName = functionArgs(2)
            ' check if fetched argument targetRangeName is really a name or just a plain range address
            If Not existsNameInWb(targetRangeName, caller.Parent.Parent) And Not existsNameInSheet(targetRangeName, caller.Parent) Then targetRangeName = ""
            ' get formula range name ...
            Dim formulaRangeName As String
            If UBound(functionArgs) > 2 Then
                formulaRangeName = functionArgs(3)
                If Not existsNameInWb(formulaRangeName, caller.Parent.Parent) And Not existsNameInSheet(formulaRangeName, caller.Parent) Then formulaRangeName = ""
            Else
                formulaRangeName = ""
            End If
            HeaderInfo = convertToBool(HeaderInfo) : AutoFit = convertToBool(AutoFit) : autoformat = convertToBool(autoformat) : ShowRowNums = convertToBool(ShowRowNums)

            ' first call: Status Container not set, actually perform query
            If Not StatusCollection.ContainsKey(callID) Then
                Dim statusCont As ContainedStatusMsg = New ContainedStatusMsg
                StatusCollection.Add(callID, statusCont)
                ExcelAsyncUtil.QueueAsMacro(Sub()
                                                DBListFetchAction(callID, CStr(Query), caller, ToRange(targetRange), CStr(ConnString), ToRange(formulaRange), extendDataArea, HeaderInfo, AutoFit, autoformat, ShowRowNums, targetRangeName, formulaRangeName)
                                            End Sub)
            End If
        Catch ex As Exception
            LogWarn(ex.Message + ", callID : " + callID)
            DBListFetch = EnvPrefix + ", Error (" + ex.Message + ") in DBListFetch, callID : " + callID
        End Try
        LogInfo("leaving function, callID: " + callID)
    End Function

    Function convertToBool(value As Object) As Boolean
        Dim tempBool As Boolean
        If TypeName(value) = "String" Then
            Dim success As Boolean = Boolean.TryParse(value, tempBool)
            If Not success Then tempBool = False
        ElseIf TypeName(value) = "Boolean" Then
            tempBool = value
        Else
            tempBool = False
        End If
        Return tempBool
    End Function

    ''' <summary>Actually do the work for DBListFetch: Query list of data delimited by maxRows and maxCols, write it into targetCells
    '''             additionally copy formulas contained in formulaRange and extend list depending on extendArea</summary>
    ''' <param name="callID"></param>
    ''' <param name="Query"></param>
    ''' <param name="caller"></param>
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
    Public Sub DBListFetchAction(callID As String, Query As String, caller As Excel.Range, targetRange As Excel.Range, ConnString As String, formulaRange As Object, extendArea As Integer, HeaderInfo As Boolean, AutoFit As Boolean, autoformat As Boolean, ShowRowNumbers As Boolean, targetRangeName As String, formulaRangeName As String)
        Dim tableRst As ADODB.Recordset
        Dim formulaFilledRange As Excel.Range = Nothing
        Dim targetSH As Excel.Worksheet, formulaSH As Excel.Worksheet = Nothing
        Dim NumFormat() As String = Nothing, NumFormatF() As String = Nothing
        Dim headingOffset, rowDataStart, startRow, startCol, arrayCols, arrayRows, copyDown As Integer
        Dim oldRows, oldCols, oldFRows, oldFCols, retrievedRows, targetColumns, formulaStart As Integer
        Dim warning As String, errMsg As String, tmpname As String

        LogInfo("Entering DBListFetchAction: callID " + callID)

        On Error Resume Next
        Dim calcMode = ExcelDnaUtil.Application.Calculation
        ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual
        ' this works around the data validation input bug and being called when COM Model is not ready
        ' when selecting a value from a list of a validated field or being invoked from a hyperlink (e.g. word), excel won't react to
        ' Application.Calculation changes, so just leave here...
        If ExcelDnaUtil.Application.Calculation <> Excel.XlCalculation.xlCalculationManual Then
            errMsg = "Error in setting Application.Calculation to Manual: " + Err.Description + " in query: " + Query
            GoTo err_0
        End If
        ExcelDnaUtil.Application.Cursor = Excel.XlMousePointer.xlWait  ' show the hourglass
        targetSH = targetRange.Parent
        formulaRange = formulaRange
        warning = ""

        Dim srcExtent As String, targetExtent As String, targetExtentF As String
        srcExtent = caller.Name.Name
        If Err.Number <> 0 Or InStr(1, srcExtent, "DBFsource") = 0 Then
            Err.Clear()
            srcExtent = "DBFsource" + Replace(Guid.NewGuid().ToString, "-", "")
            caller.Name = srcExtent
            caller.Parent.Parent.Names(srcExtent).Visible = False
            If Err.Number <> 0 Then
                errMsg = "Error in setting srcExtent name: " + Err.Description + " in query: " + Query
                GoTo err_0
            End If
        End If
        targetExtent = Replace(srcExtent, "DBFsource", "DBFtarget")
        targetExtentF = Replace(srcExtent, "DBFsource", "DBFtargetF")

        If Not formulaRange Is Nothing Then
            formulaSH = formulaRange.Parent
            ' only first row of formulaRange is important, rest will be autofilled down (actually this is needed to make the autoformat work)
            formulaRange = formulaRange.Rows(1)
        End If
        Err.Clear()

        startRow = targetRange.Cells(1, 1).Row : startCol = targetRange.Cells(1, 1).Column
        If Err.Number <> 0 Then
            errMsg = "Error in setting startRow/startCol: " + Err.Description + " in query: " + Query
            GoTo err_0
        End If

        DBModifs.preventChangeWhileFetching = True
        ' to prevent flickering...
        ExcelDnaUtil.Application.ScreenUpdating = False
        oldRows = targetSH.Parent.Names(targetExtent).RefersToRange.Rows.Count
        oldCols = targetSH.Parent.Names(targetExtent).RefersToRange.Columns.Count
        If Err.Number = 0 Then
            ' clear old data area
            targetSH.Parent.Names(targetExtent).RefersToRange.ClearContents
            If Err.Number <> 0 Then
                errMsg = "Error in clearing old data for targetExtent: (" + Err.Description + ") in query: " + Query
                GoTo err_0
            End If
        End If
        Err.Clear()
        ' if formulas are adjacent to data extend total range to formula range ! total range is used to extend DBMappers defined under the DB Function target...
        Dim additionalFormulaColumns As Integer = 0
        If Not formulaRange Is Nothing Then
            If targetSH Is formulaSH And formulaRange.Column = startCol + oldCols Then additionalFormulaColumns = formulaRange.Columns.Count
        End If
        Dim oldTotalTargetRange As Excel.Range = targetRange
        If additionalFormulaColumns > 0 Then
            oldTotalTargetRange = targetSH.Range(targetSH.Cells(startRow, startCol), targetSH.Cells(startRow + oldRows - 1, startCol + oldCols + additionalFormulaColumns))
        End If

        On Error Resume Next
        oldFRows = formulaSH.Parent.Names(targetExtentF).RefersToRange.Rows.Count
        oldFCols = formulaSH.Parent.Names(targetExtentF).RefersToRange.Columns.Count
        If Err.Number = 0 And oldFRows > 2 Then
            Err.Clear()
            ' clear old formulas
            formulaSH.Range(formulaSH.Cells(formulaRange.Row + 1, formulaRange.Column), formulaSH.Cells(formulaRange.Row + oldFRows - 1, formulaRange.Column + oldFCols - 1)).ClearContents()

            If Err.Number <> 0 Then
                errMsg = "Error in clearing old data for formulaSH: (" + Err.Description + ") in query: " + Query
                GoTo err_0
            End If
        End If
        Err.Clear()

        Dim ODBCconnString As String = ""
        If InStr(1, UCase$(ConnString), ";ODBC;") > 0 Then
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
            ExcelDnaUtil.Application.StatusBar = "Trying " + CnnTimeout.ToString + " sec. with connstring: " + ConnString
            Err.Clear()
            conn.Open(ConnString)

            If Err.Number <> 0 Then
                LogWarn("Connection Error: " + Err.Description)
                ' prevent multiple reconnecting if connection errors present...
                dontTryConnection = True
                StatusCollection(callID).statusMsg = "Connection Error: " + Err.Description
            End If
            CurrConnString = ConnString
        End If

        ExcelDnaUtil.Application.StatusBar = "Retrieving data for DBList: " + If(targetRangeName.Length > 0, targetRangeName, targetSH.Name + "!" + targetRange.Address)
        tableRst = New ADODB.Recordset
        tableRst.Open(Query, conn, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly, CommandTypeEnum.adCmdText)
        Dim dberr As String = ""
        If conn.Errors.Count > 0 Then
            Dim errcount As Integer
            For errcount = 0 To conn.Errors.Count - 1
                If conn.Errors.Item(errcount).Description <> Err.Description Then dberr = dberr + ";" + conn.Errors.Item(errcount).Description
            Next
            If dberr.Length > 0 Then dberr = " (" + dberr + ")"
        End If
        If Err.Number <> 0 Then
            errMsg = "Error in retrieving data: " + Err.Description + dberr + " in query: " + Query
            GoTo err_1
        End If
        ' this fails in case of known issue with OLEDB driver...
        retrievedRows = tableRst.RecordCount
        If Err.Number <> 0 Then
            errMsg = "Error in opening recordset: " + Err.Description + dberr + " in query: " + Query
            GoTo err_1
        End If
        Dim aborted As Boolean = XlCall.Excel(XlCall.xlAbort) ' for long running actions, allow interruption
        If aborted Then
            errMsg = "data fetching interrupted by user !"
            GoTo err_1
        End If

        ' from now on we don't propagate any errors as data is modified in sheet....
        ExcelDnaUtil.Application.StatusBar = "Displaying data for DBList: " + If(targetRangeName.Length > 0, targetRangeName, targetSH.Name + "!" + targetRange.Address)
        If tableRst.EOF Then warning = "Warning: No Data returned in query: " + Query
        ' set size for named range (size: arrayRows, arrayCols) used for resizing the data area (old extent)
        arrayCols = tableRst.Fields.Count
        arrayRows = retrievedRows
        ' need to shift down 1 row if headings are present
        arrayRows += If(HeaderInfo, 1, 0)
        rowDataStart = 1 + If(HeaderInfo, 1, 0)

        ' check whether retrieved data exceeds excel's limits and limit output (arrayRows/arrayCols) in case ...
        ' check rows
        If targetRange.Row + arrayRows > (targetRange.EntireColumn.Rows.Count + 1) Then
            warning = "row count of returned data exceeds max row of excel: start row:" + targetRange.Row.ToString + " + row count:" + arrayRows.ToString + " > max row+1:" + (targetRange.EntireColumn.Rows.Count + 1).ToString
            arrayRows = targetRange.EntireColumn.Rows.Count - startRow + 1
        End If
        ' check columns
        If targetRange.Column + arrayCols > (targetRange.EntireRow.Columns.Count + 1) Then
            warning = warning + ", column count of returned data exceed max column of excel: start column:" + targetRange.Column.ToString + " + column count:" + arrayCols.ToString + " > max column+1:" + (targetRange.EntireRow.Columns.Count + 1).ToString
            arrayCols = targetRange.EntireRow.Columns.Count - startCol + 1
        End If

        ' autoformat: copy 1st rows formats range to apply them afterwards to whole column(s)
        targetColumns = arrayCols - If(ShowRowNumbers, 0, 1)
        If autoformat Then
            arrayRows += If(HeaderInfo And arrayRows = 1, 1, 0)  ' need special case for autoformat
            For i As Integer = 0 To targetColumns
                ReDim Preserve NumFormat(i)
                NumFormat(i) = targetSH.Cells(startRow + rowDataStart - 1, startCol + i).NumberFormat
            Next
            ' now for the calculated data area
            If Not formulaRange Is Nothing Then
                For i = 0 To formulaRange.Columns.Count - 1
                    ReDim Preserve NumFormatF(i)
                    NumFormatF(i) = formulaSH.Cells(startRow + rowDataStart - 1, formulaRange.Column + i).NumberFormat
                Next
            End If
        End If
        If arrayRows = 0 Then arrayRows = 1  ' sane behavior of named range in case no data retrieved...

        ' check if formulaRange and targetRange overlap !
        Dim possibleIntersection As Excel.Range = ExcelDnaUtil.Application.Intersect(formulaRange, targetSH.Range(targetRange.Cells(1, 1), targetRange.Cells(1, 1).Offset(arrayRows - 1, arrayCols - 1)))
        Err.Clear()
        If Not possibleIntersection Is Nothing Then
            warning += ", formulaRange and targetRange intersect (" + targetSH.Name + "!" + possibleIntersection.Address + "), formula copying disabled !!"
            formulaRange = Nothing
        End If

        ' data list and formula range extension (ignored in first call after creation -> no defined name is set -> oldRows=0)...
        'TODO: check problems with shifting formula ranges in case of shifting mode 1 (cells)
        headingOffset = IIf(HeaderInfo, 1, 0)  ' use that for generally regarding headings !!
        If oldRows > 0 Then
            ' either cells/rows are shifted down (old data area was smaller than current) ...
            If oldRows < arrayRows Then
                'prevent insertion from heading row if headings are present (to not get the header formats..)
                Dim headingFirstRowPrevent As Integer = IIf(HeaderInfo And oldRows = 1 And arrayRows > 2, 1, 0)
                '1: add cells (not whole rows)
                If extendArea = 1 Then
                    targetSH.Range(targetSH.Cells(startRow + oldRows + headingFirstRowPrevent, startCol), targetSH.Cells(startRow + arrayRows - 1, startCol + oldCols - 1)).Insert(Shift:=Excel.XlDirection.xlDown)
                    If Not formulaRange Is Nothing Then
                        formulaSH.Range(formulaSH.Cells(startRow + oldFRows + headingOffset, formulaRange.Column), formulaSH.Cells(startRow + arrayRows - 1 - headingFirstRowPrevent, formulaRange.Column + oldFCols - 1)).Insert(Shift:=Excel.XlDirection.xlDown)
                    End If
                    '2: add whole rows
                ElseIf extendArea = 2 Then
                    targetSH.Rows((startRow + oldRows + headingFirstRowPrevent).ToString + ":" + (startRow + arrayRows - 1).ToString).Insert(Shift:=Excel.XlDirection.xlDown)
                    If Not formulaRange Is Nothing Then
                        ' take care not to insert twice (if we're having formulas in the same sheet)
                        If Not targetSH Is formulaSH Then formulaSH.Rows((startRow + oldFRows + headingOffset).ToString + ":" + (startRow + arrayRows - 1 - headingFirstRowPrevent).ToString).Insert(Shift:=Excel.XlDirection.xlDown)
                    End If
                End If
                'else 0: just overwrite -> no special action

                ' .... or cells/rows are shifted up (old data area was larger than current)
            ElseIf oldRows > arrayRows Then
                'prevent deletion of last row if headings are present (to not get the header formats, lose formulas, etc..)
                Dim headingLastRowPrevent As Integer = IIf(HeaderInfo And arrayRows = 1 And oldRows > 2, 1, 0)
                '1: add cells (not whole rows)
                If extendArea = 1 Then
                    targetSH.Range(targetSH.Cells(startRow + arrayRows + headingLastRowPrevent, startCol), targetSH.Cells(startRow + oldRows - 1, startCol + oldCols - 1)).Delete(Shift:=Excel.XlDirection.xlUp)
                    If Not formulaRange Is Nothing Then formulaSH.Range(formulaSH.Cells(startRow + arrayRows + headingLastRowPrevent, formulaRange.Column), formulaSH.Cells(startRow + oldFRows - 1 + headingOffset, formulaRange.Column + oldFCols - 1)).Delete(Shift:=Excel.XlDirection.xlUp)
                    '2: add whole rows
                ElseIf extendArea = 2 Then
                    targetSH.Rows((startRow + arrayRows + headingLastRowPrevent).ToString + ":" + (startRow + oldRows - 1).ToString).Delete(Shift:=Excel.XlDirection.xlUp)
                    If Not formulaRange Is Nothing Then
                        ' take care not to delete twice (if we're having formulas in the same sheet)
                        If Not targetSH Is formulaSH Then formulaSH.Rows((startRow + arrayRows + headingLastRowPrevent).ToString + ":" + (startRow + oldFRows - 1 + headingOffset).ToString).Delete(Shift:=Excel.XlDirection.xlUp)
                    End If
                End If
                '0: just overwrite -> no special action
            End If
            If Err.Number <> 0 Then
                errMsg = "Error in resizing area: " + Err.Description + " in query: " + Query
                GoTo err_1
            End If
        End If
        Dim curSheet As Excel.Worksheet = ExcelDnaUtil.Application.ActiveSheet
        targetSH.Activate()
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
                .RefreshStyle = Excel.XlCellInsertionMode.xlOverwriteCells
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
                .RefreshStyle = Excel.XlCellInsertionMode.xlOverwriteCells   ' this is required to prevent "right" shifting of cells at the beginning !
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .Refresh()
                tmpname = .Name
                .Delete()
            End With
        End If
        If Err.Number <> 0 Then
            errMsg = "Error in adding QueryTable: " + Err.Description + " in query: " + Query
            GoTo err_2
        End If
        tableRst.Close()
        ' excel doesn't delete the querytable name if it is not on the active sheet, 
        ' so Switch to the querytables sheet and back again:
        curSheet.Activate()

        '''' formulas recreation (removal and autofill new ones)
        If Not formulaRange Is Nothing Then
            formulaSH = formulaRange.Parent
            With formulaRange
                If .Row < startRow + rowDataStart - 1 Then
                    warning += "Error: formulaRange start above data-area, no formulas filled down !"
                Else
                    ' retrieve bottom of formula range
                    ' check for excels boundaries !!
                    If .Cells.Row + arrayRows > .EntireColumn.Rows.Count + 1 Then
                        warning += ", formulas would exceed max row of excel: start row:" + formulaStart.ToString + " + row count:" + arrayRows.ToString + " > max row+1:" + (.EntireColumn.Rows.Count + 1).ToString
                        copyDown = .EntireColumn.Rows.Count
                    Else
                        'the normal end of our autofilled rows = formula start + list size,
                        'reduced by offset of formula start and startRow if formulas start below data area top
                        copyDown = .Cells.Row + arrayRows - 1 - IIf(.Cells.Row > startRow, .Cells.Row - startRow, 0)
                    End If
                    ' sanity check not to fill upwards !
                    If copyDown > .Cells.Row Then .Cells.AutoFill(Destination:=formulaSH.Range(.Cells, formulaSH.Cells(copyDown, .Column + .Columns.Count - 1)))
                    formulaFilledRange = formulaSH.Range(formulaSH.Cells(.Row, .Column), formulaSH.Cells(copyDown, .Column + .Columns.Count - 1))

                    ' reassign internal name to changed formula area
                    ' delete the name to have a "clean" name area (otherwise visible = false setting wont work for dataTargetRange)
                    formulaFilledRange.Parent.Parent.Names(targetExtentF).Delete
                    Err.Clear() ' might not exist, so ignore errors here...
                    formulaFilledRange.Name = targetExtentF
                    formulaFilledRange.Name.Visible = False
                    ' reassign visible defined name to changed formula area only if defined...
                    If formulaRangeName.Length > 0 Then
                        formulaFilledRange.Name = formulaRangeName    ' NOT USING formulaFilledRange.Name.Visible = True, or hidden range will also be visible...
                    End If
                    If Err.Number <> 0 Then
                        errMsg = "Error in (re)assigning formula range name: " + Err.Description + " in query: " + Query
                        GoTo err_0
                    End If
                End If
            End With
        End If
        If Err.Number <> 0 Then
            errMsg = "Error in filling formulas: " + Err.Description + " in query: " + Query
            GoTo err_0
        End If

        ' reassign name to changed data area
        ' set the new hidden targetExtent name...
        Dim newTargetRange As Excel.Range = targetSH.Range(targetSH.Cells(startRow, startCol), targetSH.Cells(startRow + arrayRows - 1, startCol + targetColumns))
        Err.Clear() ' might not exist, so ignore errors here...

        ' delete the name to have a "clean" name area (otherwise visible = false setting wont work for dataTargetRange)
        newTargetRange.Name = targetExtent
        newTargetRange.Parent.Parent.Names(targetExtent).Visible = False
        Dim totalTargetRange As Excel.Range = newTargetRange
        If additionalFormulaColumns > 0 Then
            totalTargetRange = targetSH.Range(targetSH.Cells(startRow, startCol), targetSH.Cells(startRow + arrayRows - 1, startCol + targetColumns + additionalFormulaColumns))
        End If
        ' (re)assign visible name for the total area, if given
        If targetRangeName.Length > 0 Then
            totalTargetRange.Name = targetRangeName
        End If

        If Err.Number <> 0 Then
            errMsg = "Error in (re)assigning data target name: " + Err.Description + " (maybe known issue with 'cell like' sheetnames, e.g. 'C701 country' ?) in query: " + Query
            GoTo err_0
        End If

        ' if refreshed range is a DBMapper and it is in the current workbook, resize it
        DBModifs.resizeDBMapperRange(totalTargetRange, oldTotalTargetRange)

        '''' any warnings, errors ?
        If warning.Length > 0 Then
            If InStr(1, warning, "Error:") = 0 And InStr(1, warning, "No Data") = 0 Then
                If Left$(warning, 1) = "," Then
                    warning = Right$(warning, Len(warning) - 2)
                End If
                StatusCollection(callID).statusMsg = "Retrieved " + retrievedRows.ToString + " record" + If(retrievedRows > 1, "s", "") + ", Warning: " + warning
            Else
                StatusCollection(callID).statusMsg = warning
            End If
        Else
            StatusCollection(callID).statusMsg = "Retrieved " + retrievedRows.ToString + " record" + If(retrievedRows > 1, "s", "") + " from: " + Query
        End If

        ' autoformat: restore formats
        If autoformat Then
            For i = 0 To UBound(NumFormat)
                newTargetRange.Columns(i + 1).NumberFormat = NumFormat(i)
            Next
            ' now for the calculated cells...
            If Not formulaRange Is Nothing Then
                For i = 0 To UBound(NumFormatF)
                    formulaFilledRange.Columns(i + 1).NumberFormat = NumFormatF(i)
                Next
            End If
        End If

        If Err.Number <> 0 Then
            errMsg = "Error in restoring formats: " + Err.Description + " in query: " + Query
            LogWarn(errMsg + ", caller: " + callID)
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
            errMsg = "Error in autofitting: " + Err.Description + " in query: " + Query
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
        'targetRange.Cells(1, 1).Value = IIf(targetRange.Cells(1, 1).Value = "", " ", "")
        If errMsg.Length = 0 Then errMsg = Err.Description + " in query: " + Query
        LogWarn(errMsg + ", caller: " + callID)
        StatusCollection(callID).statusMsg = errMsg
        finishAction(calcMode, callID, "Error")
        caller.Formula += " " ' recalculate to trigger return of error messages to calling function
    End Sub

    ''' <summary>common sub to finish the action procedures, resetting anything (Cursor, calcmode, statusbar, screenupdating) that was set otherwise...</summary>
    ''' <param name="calcMode">reset calcmode to this</param>
    ''' <param name="callID">for logging purpose</param>
    ''' <param name="additionalLogInfo">for logging purpose</param>
    Private Sub finishAction(calcMode As Excel.XlCalculation, callID As String, Optional additionalLogInfo As String = "")
        DBModifs.preventChangeWhileFetching = False
        ExcelDnaUtil.Application.Cursor = Excel.XlMousePointer.xlDefault  ' To return cursor to normal
        ExcelDnaUtil.Application.StatusBar = False
        LogInfo("callID: " + callID + If(additionalLogInfo <> "", ", additionalInfo: " + additionalLogInfo, ""))
        ExcelDnaUtil.Application.ScreenUpdating = True ' coming from refresh, this might be off for dirtying "foreign" (being on a different sheet than the calling function) data targets 
        ExcelDnaUtil.Application.Calculation = calcMode
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
        Dim tempArray() As Excel.Range = Nothing ' final target array that is passed to makeCalcMsgContainer (after removing header flag)
        Dim callID As String = ""
        Dim HeaderInfo As Boolean
        Dim EnvPrefix As String = ""
        If ExcelDnaUtil.IsInFunctionWizard() Then Return "invoked from function wizard..."
        Try
            DBRowFetch = checkQueryAndTarget(Query, targetArray)
            If DBRowFetch.Length > 0 Then Exit Function
            Dim caller As Excel.Range = ToRange(XlCall.Excel(XlCall.xlfCaller))
            resolveConnstring(ConnString, EnvPrefix, False)
            ' calcContainers are identified by wbname + sheetname + function caller cell Address
            callID = "[" + caller.Parent.Parent.Name + "]" + caller.Parent.Name + "!" + caller.Address
            LogInfo("entering function, callID: " + callID)
            If dontCalcWhileClearing Then
                DBRowFetch = EnvPrefix + ", dontCalcWhileClearing = True !"
                Exit Function
            End If

            ' prepare information for action proc
            If TypeName(targetArray(0)) = "Boolean" Or TypeName(targetArray(0)) = "String" Then
                HeaderInfo = convertToBool(targetArray(0))
                For i As Integer = 1 To UBound(targetArray)
                    ReDim Preserve tempArray(i - 1)
                    If IsNothing(ToRange(targetArray(i))) Then
                        DBRowFetch = EnvPrefix + ", Part " + i.ToString() + " of Target is not a valid Range !"
                        Exit Function
                    End If
                    tempArray(i - 1) = ToRange(targetArray(i))
                Next
            ElseIf TypeName(targetArray(0)) = "ExcelEmpty" Or TypeName(targetArray(0)) = "ExcelError" Or TypeName(targetArray(0)) = "ExcelMissing" Then
                ' return appropriate error message...
                DBRowFetch = EnvPrefix + ", First argument (header) " + Replace(TypeName(targetArray(0)), "Excel", "") + " !"
                Exit Function
            Else
                For i = 0 To UBound(targetArray)
                    ReDim Preserve tempArray(i)
                    If IsNothing(ToRange(targetArray(i))) Then
                        DBRowFetch = EnvPrefix + ", Part " + (i + 1).ToString() + " of Target is not a valid Range !"
                        Exit Function
                    End If
                    tempArray(i) = ToRange(targetArray(i))
                Next
            End If
            ' check query, also converts query to string (if it is a range)
            ' error message or cached status message is returned from checkParamsAndCache, if query OK and result was not already calculated (cached) then empty string
            DBRowFetch = checkParamsAndCache(Query, callID, ConnString)
            If DBRowFetch.Length > 0 Then
                DBRowFetch = EnvPrefix + ", " + DBRowFetch
                Exit Function
            End If

            ' first call: actually perform query
            If Not StatusCollection.ContainsKey(callID) Then
                Dim statusCont As ContainedStatusMsg = New ContainedStatusMsg
                StatusCollection.Add(callID, statusCont)
                StatusCollection(callID).statusMsg = "" ' need this to prevent object not set errors in checkCache
                ExcelAsyncUtil.QueueAsMacro(Sub()
                                                DBRowFetchAction(callID, CStr(Query), caller, tempArray, CStr(ConnString), HeaderInfo)
                                            End Sub)
            End If
        Catch ex As Exception
            LogWarn(ex.Message + ", callID: " + callID)
            DBRowFetch = EnvPrefix + ", Error (" + ex.Message + ") in DBRowFetch, callID: " + callID
        End Try
        LogInfo("leaving function, callID: " + callID)
    End Function

    ''' <summary>Actually do the work for DBRowFetch: Query (assumed) one row of data, write it into targetCells</summary>
    ''' <param name="callID"></param>
    ''' <param name="Query"></param>
    ''' <param name="caller"></param>
    ''' <param name="targetArray"></param>
    ''' <param name="ConnString"></param>
    ''' <param name="HeaderInfo"></param>
    Public Sub DBRowFetchAction(callID As String, Query As String, caller As Excel.Range, targetArray As Object, ConnString As String, HeaderInfo As Boolean)
        Dim tableRst As ADODB.Recordset = Nothing
        Dim targetCells As Object
        Dim errMsg As String = "", refCollector As Excel.Range
        Dim headerFilled As Boolean, DeleteExistingContent As Boolean, fillByRows As Boolean
        Dim returnedRows As Long, fieldIter As Integer, rangeIter As Integer
        Dim theCell As Excel.Range, targetSlice As Excel.Range, targetSlices As Excel.Range
        Dim targetSH As Excel.Worksheet

        On Error Resume Next
        targetCells = targetArray
        targetSH = targetCells(0).Parent
        StatusCollection(callID).statusMsg = ""
        Dim calcMode = ExcelDnaUtil.Application.Calculation
        ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual
        ' this works around the data validation input bug and being called when COM Model is not ready
        ' when selecting a value from a list of a validated field or being invoked from a hyperlink (e.g. word), excel won't react to
        ' Application.Calculation changes, so just leave here...
        If ExcelDnaUtil.Application.Calculation <> Excel.XlCalculation.xlCalculationManual Then
            errMsg = "Error in setting Application.Calculation to Manual: " + Err.Description + " in query: " + Query
            GoTo err_1
        End If
        ExcelDnaUtil.Application.Cursor = Excel.XlMousePointer.xlWait  ' show the hourglass
        On Error GoTo err_1
        ExcelDnaUtil.Application.StatusBar = "Retrieving data for DBRows: " + targetSH.Name + "!" + targetCells(0).Address

        Dim srcExtent As String, targetExtent As String
        On Error Resume Next
        srcExtent = caller.Name.Name
        If Err.Number <> 0 Or InStr(1, srcExtent, "DBFsource") = 0 Then
            Err.Clear()
            srcExtent = "DBFsource" + Replace(Guid.NewGuid().ToString, "-", "")
            caller.Name = srcExtent
            ' dbfsource is a workbook name
            caller.Parent.Parent.Names(srcExtent).Visible = False
            If Err.Number <> 0 Then
                errMsg = "Error in setting srcExtent name: " + Err.Description + " in query: " + Query
                GoTo err_1
            End If
        End If
        targetExtent = Replace(srcExtent, "DBFsource", "DBFtarget")
        DBModifs.preventChangeWhileFetching = True
        ' to prevent flickering...
        ExcelDnaUtil.Application.ScreenUpdating = False
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
            ExcelDnaUtil.Application.StatusBar = "Trying " + CnnTimeout.ToString + " sec. with connstring: " + ConnString
            Err.Clear()
            conn.Open(ConnString)

            If Err.Number <> 0 Then
                LogWarn("Connection Error: " + Err.Description)
                ' prevent multiple reconnecting if connection errors present...
                dontTryConnection = True
                StatusCollection(callID).statusMsg = "Connection Error: " + Err.Description
            End If
            CurrConnString = ConnString
        End If

        tableRst = New ADODB.Recordset
        tableRst.Open(Query, conn, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly, CommandTypeEnum.adCmdText)
        On Error Resume Next
        Dim dberr As String = ""
        If conn.Errors.Count > 0 Then
            Dim errcount As Integer
            For errcount = 0 To conn.Errors.Count - 1
                If conn.Errors.Item(errcount).Description <> Err.Description Then _
                   dberr = dberr + ";" + conn.Errors.Item(errcount).Description
            Next
            errMsg = "Error in retrieving data: " + dberr + " in query: " + Query
            GoTo err_1
        End If

        ' this fails in case of known issue with OLEDB driver...
        returnedRows = tableRst.RecordCount
        If Err.Number <> 0 Then
            errMsg = "Error in opening recordset: " + Err.Description + " in query: " + Query
            GoTo err_1
        End If

        On Error GoTo err_1
        ' check whether anything retrieved? if not, delete possible existing content...
        DeleteExistingContent = tableRst.EOF
        If DeleteExistingContent Then StatusCollection(callID).statusMsg = "Warning: No Data returned in query: " + Query

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
                        theCell.Value = ""
                    Else
                        If Not headerFilled Then
                            theCell.Value = tableRst.Fields(fieldIter).Name
                        ElseIf DeleteExistingContent Then
                            theCell.Value = ""
                        Else
                            On Error Resume Next
                            theCell.Value = tableRst.Fields(fieldIter).Value
                            If Err.Number <> 0 Then errMsg += "Field '" + tableRst.Fields(fieldIter).Name + "' caused following error: '" + Err.Description + "'"
                            On Error GoTo err_1
                        End If
                        If fieldIter = tableRst.Fields.Count - 1 Then
                            If headerFilled Then
                                ExcelDnaUtil.Application.StatusBar = "Displaying data for DBRows: " + targetSH.Name + "!" + targetCells(0).Address.ToString + ", record " + tableRst.AbsolutePosition.ToString + "/" + returnedRows.ToString
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
            If Not rangeIter > UBound(targetCells) Then refCollector = ExcelDnaUtil.Application.Union(refCollector, targetCells(rangeIter))
        Loop Until rangeIter > UBound(targetCells)

        ' delete the name to have a "clean" name area (otherwise visible = false setting wont work for dataTargetRange)
        refCollector.Name = targetExtent
        refCollector.Parent.Parent.Names(targetExtent).Visible = False

        tableRst.Close()
        If StatusCollection(callID).statusMsg.Length = 0 Then StatusCollection(callID).statusMsg = "Retrieved " + returnedRows.ToString + " record" + If(returnedRows > 1, "s", "") + " from: " + Query
        finishAction(calcMode, callID)
        Exit Sub

err_1:
        If errMsg.Length = 0 Then errMsg = Err.Description + " in query: " + Query
        If tableRst.State <> 0 Then tableRst.Close()
        LogWarn(errMsg + ", caller: " + callID)
        StatusCollection(callID).statusMsg = errMsg
        finishAction(calcMode, callID, "Error")
        caller.Formula += " " ' recalculate to trigger return of error messages to calling function
    End Sub


    ''' <summary>remove all names from Range Target except the passed name (theName) and store them into list storedNames</summary>
    ''' <param name="Target"></param>
    ''' <param name="theName"></param>
    ''' <returns>the removed names as a string list for restoring them later (see restoreRangeNames)</returns>
    Private Function removeRangeNames(Target As Excel.Range, theName As String) As String()
        Dim storedNames() As String = {}
        Dim nextName As String

        Dim i As Integer = 0
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
        removeRangeNames = storedNames
    End Function

    ''' <summary>restore the passed storedNames into Range Target</summary>
    ''' <param name="Target"></param>
    ''' <param name="storedNames"></param>
    Private Sub restoreRangeNames(Target As Excel.Range, storedNames() As String)
        If UBound(storedNames) > 0 Then
            For Each theName As String In storedNames
                If theName.Length > 0 Then Target.Name = theName
            Next
        End If
    End Sub

    ''' <summary>Get the current selected Environment for DB Functions</summary>
    ''' <returns>ConfigName of environment</returns>
    <ExcelFunction(Description:="Get the current selected Environment for DB Functions")>
    Public Function DBAddinEnvironment() As String
        ExcelDnaUtil.Application.Volatile()
        Try
            DBAddinEnvironment = fetchSetting("ConfigName" + Globals.env(), "")
            If ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual Then DBAddinEnvironment = "calc Mode is manual, please press F9 to get current DBAddin environment !"
        Catch ex As Exception
            DBAddinEnvironment = "Error happened: " + ex.Message
            LogWarn(ex.Message)
        End Try
    End Function

    ''' <summary>Get the server settings for the currently selected Environment for DB Functions</summary>
    ''' <returns>Server part from connection string of environment</returns>
    <ExcelFunction(Description:="Get the server settings for the currently selected Environment for DB Functions")>
    Public Function DBAddinServerSetting() As String
        ExcelDnaUtil.Application.Volatile()
        Try
            Dim theConnString As String = fetchSetting("ConstConnString" + Globals.env(), "")
            Dim keywordstart As Integer = InStr(1, UCase(theConnString), "SERVER=") + Len("SERVER=")
            DBAddinServerSetting = Mid$(theConnString, keywordstart, InStr(keywordstart, theConnString, ";") - keywordstart)
            If ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual Then DBAddinServerSetting = "calc Mode is manual, please press F9 to get current DBAddin server setting !"
        Catch ex As Exception
            DBAddinServerSetting = "Error happened: " + ex.Message
            LogWarn(ex.Message)
        End Try
    End Function

    ''' <summary>checks Query and targetRange parameters for existence and return error message.</summary>
    ''' <param name="Query"></param>
    ''' <param name="targetRange"></param>
    ''' <returns>Error String or cached status message (empty if OK)</returns>
    Private Function checkQueryAndTarget(Query As Object, targetRange As Object) As String
        If TypeName(Query) = "ExcelMissing" Then
            checkQueryAndTarget = "missing Query parameter !"
        ElseIf TypeName(targetRange) = "ExcelMissing" Then
            checkQueryAndTarget = "missing target range parameter !"
        ElseIf TypeName(targetRange) = "Object()" AndAlso targetRange.Length = 0 Then
            checkQueryAndTarget = "missing target parameter array !"
        Else
            checkQueryAndTarget = ""
        End If
    End Function

    ''' <summary>checks calculation mode, query and cached status message.</summary>
    ''' <param name="Query"></param>
    ''' <param name="callID"></param>
    ''' <param name="ConnString"></param>
    ''' <returns>Error String or cached status message (empty if OK)</returns>
    Private Function checkParamsAndCache(ByRef Query, callID, ConnString) As String
        ' can't give types to parameters as Query can be Object(,) and other types, so Object is not possible.
        ' additionally VB.NET forces us to give ALL parameters a type, so no Option Strict here !
        checkParamsAndCache = ""
        If ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual Then
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
                        retval += Convert.ToString(myCell, System.Globalization.CultureInfo.InvariantCulture) + " "
                    Else
                        retval += myCell.ToString + " "
                    End If
                    Query = retval
                Next
                If retval.Length = 0 Then checkParamsAndCache = "empty query provided !"
            ElseIf TypeName(Query) = "String" Then
                If Query.ToString.Length = 0 Then checkParamsAndCache = "empty query provided !"
            Else
                checkParamsAndCache = "query parameter invalid (not a range and not a string) !"
            End If
        End If
        If checkParamsAndCache.Length > 0 Then
            ' refresh the query cache ...
            If queryCache.ContainsKey(callID) Then queryCache.Remove(callID)
            Exit Function
        End If

        ' caching check mechanism to avoid unnecessary recalculations/refetching
        Dim doFetching As Boolean
        If queryCache.ContainsKey(callID) Then
            doFetching = (ConnString + Query.ToString <> queryCache(callID))
            ' refresh the query cache with new query/connstring ...
            queryCache.Remove(callID)
            queryCache.Add(callID, ConnString + Query.ToString)
        Else
            queryCache.Add(callID, ConnString + Query.ToString)
            doFetching = True
        End If
        If doFetching Then
            ' remove Status Container to signal a new calculation request
            If StatusCollection.ContainsKey(callID) Then StatusCollection.Remove(callID)
        Else
            ' return Status Containers Message as last result
            If StatusCollection.ContainsKey(callID) Then
                If Not IsNothing(StatusCollection(callID).statusMsg) Then checkParamsAndCache = "(last result:)" + StatusCollection(callID).statusMsg
            End If
        End If
    End Function


    ''' <summary>create a final connection string from passed String or number (environment), as well as a EnvPrefix for showing the environment (or set ConnString)</summary>
    ''' <param name="ConnString">passed connection string or environment number, resolved (=returned) to actual connection string</param>
    ''' <param name="EnvPrefix">prefix for showing environment (ConnString set if no environment)</param>
    Public Sub resolveConnstring(ByRef ConnString As Object, ByRef EnvPrefix As String, getConnStrForDBSet As Boolean)
        If Left(TypeName(ConnString), 10) = "ExcelError" Then Exit Sub
        If TypeName(ConnString) = "ExcelReference" Then ConnString = ConnString.Value
        If TypeName(ConnString) = "ExcelMissing" Then ConnString = ""
        ' in case ConnString is a number (set environment, retrieve ConnString from Setting ConstConnString<Number>
        If TypeName(ConnString) = "Double" Then
            Dim env As String = ConnString.ToString()
            EnvPrefix = "Env:" + fetchSetting("ConfigName" + env, "")
            ConnString = fetchSetting("ConstConnString" + env, "")
            If getConnStrForDBSet Then
                ' if an alternate connection string is given, use this one...
                Dim altConnString = fetchSetting("AltConnString" + env, "")
                If altConnString <> "" Then
                    ConnString = altConnString
                Else
                    ' To get the connection string work also for SQLOLEDB provider for SQL Server, change to ODBC driver setting (this can be generally used to fix connection string problems with ListObjects)
                    ConnString = Replace(ConnString, fetchSetting("ConnStringSearch" + env, "provider=SQLOLEDB"), fetchSetting("ConnStringReplace" + env, "driver=SQL SERVER"))
                End If
            End If
        ElseIf TypeName(ConnString) = "String" Then
            If ConnString.ToString = "" Then ' no ConnString or environment number set: get connection string of currently selected evironment
                EnvPrefix = "Env:" + fetchSetting("ConfigName" + Globals.env(), "")
                ConnString = fetchSetting("ConstConnString" + Globals.env(), "")
                If getConnStrForDBSet Then
                    ' if an alternate connection string is given, use this one...
                    Dim altConnString = fetchSetting("AltConnString" + Globals.env(), "")
                    If altConnString <> "" Then
                        ConnString = altConnString
                    Else
                        ' To get the connection string work also for SQLOLEDB provider for SQL Server, change to ODBC driver setting (this can be generally used to fix connection string problems with ListObjects)
                        ConnString = Replace(ConnString, fetchSetting("ConnStringSearch" + Globals.env(), "provider=SQLOLEDB"), fetchSetting("ConnStringReplace" + Globals.env(), "driver=SQL SERVER"))
                    End If
                End If
            Else
                EnvPrefix = "ConnString set"
            End If
        End If
    End Sub

    ''' <summary>checks whether theName exists as a name in Workbook theWb</summary>
    ''' <param name="theName"></param>
    ''' <param name="theWb"></param>
    ''' <returns>true if it exists</returns>
    Public Function existsNameInWb(ByRef theName As String, theWb As Excel.Workbook) As Boolean
        existsNameInWb = False
        For Each aName As Excel.Name In theWb.Names()
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
    Public Function existsNameInSheet(ByRef theName As String, theWs As Excel.Worksheet) As Boolean
        existsNameInSheet = False
        For Each aName As Excel.Name In theWs.Names()
            If aName.Name = theWs.Name + "!" + theName Then
                existsNameInSheet = True
                Exit Function
            End If
        Next
    End Function

    ''' <summary>converts ExcelDna (C API) reference to excel (COM Based) Range</summary>
    ''' <param name="reference">reference to be converted</param>
    ''' <returns>range for passed reference</returns>
    Private Function ToRange(reference As Object) As Excel.Range
        If TypeName(reference) <> "ExcelReference" Then Return Nothing

        Dim item As String = XlCall.Excel(XlCall.xlSheetNm, reference)
        Dim index As Integer = item.LastIndexOf("]")
        Dim wbname As String = item.Substring(0, index).Substring(1)
        item = item.Substring(index + 1)
        Dim ws As Excel.Worksheet = ExcelDnaUtil.Application.Workbooks(wbname).Worksheets(item)
        Return ws.Range(ws.Cells(reference.RowFirst + 1, reference.ColumnFirst + 1), ws.Cells(reference.RowLast + 1, reference.ColumnLast + 1))
    End Function
End Module