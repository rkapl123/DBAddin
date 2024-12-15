Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports System.Collections.Generic
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Data.SqlClient

''' <summary>Provides a data structure for transporting information back from the calculation action procedure to the calling function resp. the AfterCalculate event procedure</summary>
Public Class ContainedStatusMsg
    ''' <summary>any status message used for displaying in the result of function</summary>
    Public statusMsg As String
    ''' <summary>formula range passed from dblistfetchAction to overcome the problem of auto-fitting AFTER calculation</summary>
    Public formulaRange As Excel.Range
End Class

''' <summary>Contains the public callable DB functions and helper functions</summary>
Public Module Functions
    ' Global objects/variables for DBFuncs
    ''' <summary>global collection of information transport containers between action function and user-defined function resp. calc event procedure</summary>
    Public StatusCollection As Dictionary(Of String, ContainedStatusMsg)
    ''' <summary>connection object</summary>
    Public conn As System.Data.IDbConnection
    ''' <summary>connection string can be changed for calls with different connection strings</summary>
    Public CurrConnString As String
    ''' <summary>query cache for avoiding unnecessary recalculations/data retrievals by volatile inputs to DB Functions (now(), etc.)</summary>
    Public queryCache As Dictionary(Of String, String)
    ''' <summary>avoid entering dblistfetch/dbrowfetch functions during clearing of list-fetch areas (before saving)</summary>
    Public dontCalcWhileClearing As Boolean
    ''' <summary>globally avoid refreshing of DB Functions</summary>
    Public preventRefreshFlag As Boolean
    ''' <summary>avoid refreshing of DB Functions on Workbook level (has precedence on global preventRefreshFlag)</summary>
    Public preventRefreshFlagColl As New Dictionary(Of String, Boolean)

    ''' <summary>set preventing of refreshing DB Functions</summary>
    ''' <param name="setPreventRefresh">set preventRefreshFlag to this value</param>
    ''' <returns>whether the preventRefresh flag was changed</returns>
    <ExcelFunction(Description:="set preventing of refreshing DB Functions")>
    Public Function preventRefresh(<ExcelArgument(Description:="set to TRUE to prevent refresh, FALSE to enable again")> setPreventRefresh As Boolean, <ExcelArgument(Description:="set to TRUE to prevent/enable only for this Workbook")> Optional onlyForThisWB As Boolean = True) As String
        If onlyForThisWB Then
            If preventRefreshFlagColl.ContainsKey(ExcelDnaUtil.Application.ActiveWorkbook.Name) Then
                preventRefreshFlagColl(ExcelDnaUtil.Application.ActiveWorkbook.Name) = setPreventRefresh
            Else
                preventRefreshFlagColl.Add(ExcelDnaUtil.Application.ActiveWorkbook.Name, setPreventRefresh)
            End If
        Else
            If preventRefreshFlagColl.ContainsKey(ExcelDnaUtil.Application.ActiveWorkbook.Name) Then preventRefreshFlagColl.Remove(ExcelDnaUtil.Application.ActiveWorkbook.Name)
            preventRefreshFlag = setPreventRefresh
        End If
        theRibbon.InvalidateControl("preventRefresh")
        Dim toWhichExtent As String = IIf(onlyForThisWB, "for " + ExcelDnaUtil.Application.ActiveWorkbook.Name, "globally ")
        Return "preventRefresh was set " + toWhichExtent + " to " + setPreventRefresh.ToString() + IIf(setPreventRefresh, ", DB Functions will " + toWhichExtent + " not refresh on any change or refresh context menu", ", DB Functions will refresh again " + toWhichExtent)
    End Function

    ''' <summary>Create database compliant date, time or datetime string from excel date type value</summary>
    ''' <param name="DatePart">date/time/datetime single parameter or range reference</param>
    ''' <param name="formatting">formatting instruction for Date format, see remarks</param>
    ''' <returns>the DB compliant formatted date/time/datetime</returns>
    ''' <remarks>
    ''' formatting = 0: A simple date string (format 'YYYYMMDD'), datetime values are converted to 'YYYYMMDD HH:MM:SS' and time values are converted to 'HH:MM:SS'.
    ''' formatting = 1: An ANSI compliant Date string (format date 'YYYY-MM-DD'), datetime values are converted to 'YYYY-MM-DD HH:MM:SS' and time values are converted to 'HH:MM:SS'.
    ''' formatting = 2: An ODBC compliant Date string (format {d 'YYYY-MM-DD'}), datetime values are converted to {ts 'YYYY-MM-DD HH:MM:SS'} and time values are converted to {t 'HH:MM:SS'}.
    ''' formatting = 3: An Access/JetDB compliant Date string (format #YYYY-MM-DD#), datetime values are converted to #YYYY-MM-DD HH:MM:SS# and time values are converted to #HH:MM:SS#.
    ''' add 10 to formatting to include fractions of a second (1000) 
    ''' formatting >13 or empty (99=default value): take the formatting option from setting DefaultDBDateFormatting (0 if not given)
    ''' </remarks>
    <ExcelFunction(Description:="Create database compliant date, time or datetime string from excel date type value")>
    Public Function DBDate(<ExcelArgument(Description:="date/time/datetime")> ByVal DatePart As Object,
                           <ExcelArgument(Description:="formatting option, 0:'YYYYMMDD', 1: DATE 'YYYY-MM-DD'), 2:{d 'YYYY-MM-DD'},3:Access/JetDB #DD/MM/YYYY#, add 10 to formatting to include fractions of a second (1000)")> Optional formatting As Integer = 99) As String
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

    ''' <summary>Create a powerquery compliant string from cell values, potentially concatenating with other parts for easy inclusion of wildcards (%,_)</summary>
    ''' <param name="StringPart">array of strings/wildcards or ranges containing strings/wildcards</param>
    ''' <returns>powerquery compliant string</returns>
    <ExcelFunction(Description:="Create a powerquery compliant string from cell values, potentially concatenating with other parts for easy inclusion of wildcards (%,_)")>
    Public Function PQString(<ExcelArgument(Description:="array of strings/wildcards or ranges containing strings/wildcards")> ParamArray StringPart() As Object) As String
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
            PQString = """" + retval + """"
        Catch ex As Exception
            LogWarn(ex.Message)
            PQString = "Error (" + ex.Message + ") in PQString"
        End Try
    End Function

    ''' <summary>Creates a powerquery compliant #date function from excel date type value</summary>
    ''' <param name="DatePart">date/time/datetime single parameter or range reference</param>
    ''' <returns>the powerquery #date function</returns>
    <ExcelFunction(Description:="Create powerquery compliant #date, #time or #datetime function from excel date type value")>
    Public Function PQDate(<ExcelArgument(Description:="date/time/datetime")> ByVal DatePart As Object,
                           <ExcelArgument(Description:="enforce datetime for date only values (without fractional part)")> Optional forceDateTime As Boolean = False) As String
        PQDate = ""
        Try
            If TypeName(DatePart) = "Object(,)" Then
                For Each myCell In DatePart
                    If TypeName(myCell) = "ExcelEmpty" Then
                        ' do nothing here
                    Else
                        PQDate = formatPQDate(myCell, forceDateTime) + ","
                    End If
                Next
                ' cut last comma
                If PQDate.Length > 0 Then PQDate = Left(PQDate, Len(PQDate) - 1)
            Else
                ' direct value in DBDate..
                If TypeName(DatePart) = "ExcelEmpty" Or TypeName(DatePart) = "ExcelMissing" Then
                    ' do nothing here
                Else
                    PQDate = formatPQDate(DatePart, forceDateTime)
                End If
            End If
        Catch ex As Exception
            LogWarn(ex.Message)
            PQDate = "Error (" + ex.Message + ") in function PQDate"
        End Try
    End Function

    ''' <summary>takes an OADate and formats it as a powerquery compliant #date, #time or #datetime function</summary>
    ''' <param name="datVal">OADate (double) date parameter</param>
    ''' <returns>powerquery function</returns>
    Private Function formatPQDate(datVal As Double, forceDateTime As Boolean) As String
        formatPQDate = ""
        If Int(datVal) = datVal And Not forceDateTime Then
            formatPQDate += "#date(" + Date.FromOADate(datVal).Year.ToString() + "," + Date.FromOADate(datVal).Month.ToString() + "," + Date.FromOADate(datVal).Day.ToString() + ")"
        ElseIf CInt(datVal) > 1 Or forceDateTime Then
            formatPQDate += "#datetime(" + Date.FromOADate(datVal).Year.ToString() + "," + Date.FromOADate(datVal).Month.ToString() + "," + Date.FromOADate(datVal).Day.ToString() + "," + Date.FromOADate(datVal).Hour.ToString() + "," + Date.FromOADate(datVal).Minute.ToString() + "," + Date.FromOADate(datVal).Second.ToString() + ")"
        Else
            formatPQDate += "#time(" + Date.FromOADate(datVal).Hour.ToString() + "," + Date.FromOADate(datVal).Minute.ToString() + "," + Date.FromOADate(datVal).Second.ToString() + ")"
        End If
    End Function

    ''' <summary>Create an in clause from cell values, strings are created with quotation marks,
    '''             dates are created with DBDate</summary>
    ''' <param name="inClausePart">array of values or ranges containing values</param>
    ''' <returns>database compliant in-clause string</returns>
    <ExcelFunction(Description:="Create an in clause from cell values, strings are created with quotation marks")>
    Public Function DBinClause(<ExcelArgument(AllowReference:=True, Description:="array of values or ranges containing values")> ParamArray inClausePart As Object()) As String
        Dim concatResult As String = DoConcatCellsSep(",", True, False, False, inClausePart)
        ' for empty concatenation results, return "in (NULL)" to get a valid SQL String (required for chained queries!)
        DBinClause = If(Left(concatResult, 5) = "Error", concatResult, If(concatResult = "", "in (NULL)", "in (" + concatResult + ")"))
    End Function

    ''' <summary>Create an in clause from cell values, strings are created with quotation marks,
    '''             dates are created with DBDate</summary>
    ''' <param name="inClausePart">array of values or ranges containing values</param>
    ''' <returns>database compliant in-clause string</returns>
    <ExcelFunction(Description:="Create an in clause from cell values, all arguments are treated as strings (and will be created with quotation marks)")>
    Public Function DBinClauseStr(<ExcelArgument(AllowReference:=True, Description:="array of values or ranges containing values")> ParamArray inClausePart As Object()) As String
        Dim concatResult As String = DoConcatCellsSep(",", True, True, False, inClausePart)
        ' for empty concatenation results, return "in (NULL)" to get a valid SQL String (required for chained queries!)
        DBinClauseStr = If(Left(concatResult, 5) = "Error", concatResult, If(concatResult = "", "in (NULL)", "in (" + concatResult + ")"))
    End Function

    ''' <summary>Create an in clause from cell values, strings are created with quotation marks,
    '''             dates are created with DBDate</summary>
    ''' <param name="inClausePart">array of values or ranges containing values</param>
    ''' <returns>database compliant in-clause string</returns>
    <ExcelFunction(Description:="Create an in clause from cell values, numbers are always treated as dates (formatted using default) and created with quotation marks")>
    Public Function DBinClauseDate(<ExcelArgument(AllowReference:=True, Description:="array of values or ranges containing values")> ParamArray inClausePart As Object()) As String
        Dim concatResult As String = DoConcatCellsSep(",", True, False, True, inClausePart)
        ' for empty concatenation results, return "in (NULL)" to get a valid SQL String (required for chained queries!)
        DBinClauseDate = If(Left(concatResult, 5) = "Error", concatResult, If(concatResult = "", "in (NULL)", "in (" + concatResult + ")"))
    End Function

    ''' <summary>concatenates values contained in concatPart together (using .value attribute for cells)</summary>
    ''' <param name="concatPart">all cells/values which should be concatenated</param>
    ''' <returns>concatenated String</returns>
    <ExcelFunction(Description:="concatenates values contained in concatPart together (using .value attribute for cells)")>
    Public Function concatCells(<ExcelArgument(AllowReference:=True, Description:="all cells/values which should be concatenated")> ParamArray concatPart As Object()) As String
        concatCells = DoConcatCellsSep("", False, False, False, concatPart)
    End Function

    ''' <summary>concatenates values contained in concatPart (using .value for cells) using a separator</summary>
    ''' <param name="separator">the separator</param>
    ''' <param name="concatPart">all cells/values which should be concatenated</param>
    ''' <returns>concatenated String</returns>
    <ExcelFunction(Description:="concatenates values contained in concatPart (using .value for cells) using a separator")>
    Public Function concatCellsSep(<ExcelArgument(AllowReference:=True, Description:="the separator")> separator As String,
                                   <ExcelArgument(AllowReference:=True, Description:="all cells/values which should be concatenated")> ParamArray concatPart As Object()) As String
        concatCellsSep = DoConcatCellsSep(separator, False, False, False, concatPart)
    End Function

    ''' <summary>chains values contained in chainPart together with commas, mainly used for creating select header</summary>
    ''' <param name="chainPart">range where values should be chained</param>
    ''' <returns>chained String</returns>
    <ExcelFunction(Description:="chains values contained in chainPart together with commas, mainly used for creating select header")>
    Public Function chainCells(<ExcelArgument(AllowReference:=True, Description:="range where values should be chained")> ParamArray chainPart As Object()) As String
        chainCells = DoConcatCellsSep(",", False, False, False, chainPart)
    End Function

    ''' <summary>get current Workbook path + filename or Workbook path only, if onlyPath is set</summary>
    ''' <param name="onlyPath">only path of file location?</param>
    ''' <returns>current Workbook path + filename or Workbook path only</returns>
    <ExcelFunction(Description:="get current Workbook path + filename or Workbook path only, if onlyPath is set")>
    Public Function currentWorkbook(<ExcelArgument(Description:="only path of file location?")> Optional onlyPath As Boolean = False) As String
        If onlyPath Then
            currentWorkbook = ExcelDnaUtil.Application.ActiveWorkbook.Path + "\"
        Else
            currentWorkbook = ExcelDnaUtil.Application.ActiveWorkbook.Path + "\" + ExcelDnaUtil.Application.ActiveWorkbook.Name
        End If
    End Function

    ''' <summary>private function that actually concatenates values contained in Object array concatParts together (either using .text or .value for cells in concatParts) using a separator</summary>
    ''' <param name="separator">the separator-string that is filled between values</param>
    ''' <param name="DBcompliant">should a potential string or date part be formatted database compliant (surrounded by quotes)?</param>
    ''' <param name="OnlyString">set when only DB compliant Strings should be produced during concatenation</param>
    ''' <param name="OnlyDate">set when only DB compliant Dates should be produced during concatenation</param>
    ''' <param name="concatParts">Object array, whose values should be concatenated</param>
    ''' <returns>concatenated String</returns>
    Private Function DoConcatCellsSep(separator As String, DBcompliant As Boolean, OnlyString As Boolean, OnlyDate As Boolean, ParamArray concatParts As Object()) As String
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
                            ElseIf OnlyDate Then
                                retval = retval + separator + formatDBDate(myCell, DefaultDBDateFormatting)
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
                    ElseIf IsNumeric(myRef) Then
                        If OnlyString Then
                            retval = retval + separator + IIf(DBcompliant, "'", "") + Convert.ToString(myRef, System.Globalization.CultureInfo.InvariantCulture) + IIf(DBcompliant, "'", "")
                        ElseIf OnlyDate Then
                            retval = retval + separator + formatDBDate(myRef, DefaultDBDateFormatting)
                        Else
                            retval = retval + separator + Convert.ToString(myRef, System.Globalization.CultureInfo.InvariantCulture)
                        End If
                    Else
                        ' avoid double quoting if passed string is already quoted or in date format (by using DBDate or DBString as input to this) and DBcompliant quoting is requested
                        retval = retval + separator + IIf(DBcompliant And Not (Left(myRef, 1) = "'"), "'", "") + myRef + IIf(DBcompliant And Right(myRef, 1) <> "'", "'", "")
                    End If
                End If
            Next
            DoConcatCellsSep = Mid$(retval, Len(separator) + 1) ' skip first separator
        Catch ex As Exception
            LogWarn(ex.Message)
            DoConcatCellsSep = "Error (" + ex.Message + ") in DoConcatCellsSep"
        End Try
    End Function


    ''' <summary>Stores a query into an Object defined in targetRange (an embedded MS Query/List object, Pivot table, etc.)</summary>
    ''' <param name="Query">query for getting data</param>
    ''' <param name="ConnString">connection string defining DB, user, etc...</param>
    ''' <param name="targetRange">Range with Object beneath to put the Query/ConnString into</param>
    ''' <returns>Status Message</returns>
    <ExcelFunction(Description:="Stores a query into an Object (embedded List object or Pivot table) defined in targetRange")>
    Public Function DBSetQuery(<ExcelArgument(Description:="query for getting data")> Query As Object,
                               <ExcelArgument(Description:="connection string defining DB, user, etc...")> ConnString As Object,
                               <ExcelArgument(Description:="Range with embedded Object to put the Query/ConnString into", AllowReference:=True)> targetRange As Object) As String
        Dim callID As String = ""
        Dim caller As Excel.Range = Nothing
        Dim EnvPrefix As String = ""
        If ExcelDnaUtil.IsInFunctionWizard() Then Return "invoked from function wizard..."
        If checkPreventRefreshFlag() Then Return "preventRefresh set, DB Function not refreshing..."
        Try
            DBSetQuery = checkQueryAndTarget(Query, targetRange)
            If DBSetQuery.Length > 0 Then Exit Function
            caller = ToRange(XlCall.Excel(XlCall.xlfCaller))
            resolveConnstring(ConnString, EnvPrefix, True)
            ' calcContainers are identified by workbook name + Sheet name + function caller cell Address
            callID = "[" + caller.Parent.Parent.Name + "]" + caller.Parent.Name + "!" + caller.Address
            ' check query, also converts query to string (if it is a range)
            ' error message or cached status message is returned from checkParamsAndCache, if query OK and result was not already calculated (cached) then empty string
            DBSetQuery = checkParamsAndCache(Query, callID, ConnString)
            If DBSetQuery.Length > 0 Then
                DBSetQuery = EnvPrefix + ", " + DBSetQuery
                Exit Function
            End If

            ' needed for check whether target range is actually a table List object reference
            Dim functionArgs = functionSplit(caller.Formula, ",", """", "DBSetQuery", "(", ")")
            ' targetRangeName can be either a simple range (A1) or a list object reference
            ' this list object reference can either be a complete table (sub)range: Tablename[[#Headerrows]] or Tablename[fieldname] or Tablename[#All]
            ' or a separator split single cell: Tablename[[#Headerrows],[fieldname]]
            ' in the latter case, functionSplit will return 4 arguments and the last needs to be rejoined
            Dim targetRangeName As String = functionArgs(2)
            If UBound(functionArgs) = 3 Then targetRangeName += "," + functionArgs(3)

            ' first call: actually perform query
            If Not StatusCollection.ContainsKey(callID) Then
                Dim statusCont As New ContainedStatusMsg
                StatusCollection.Add(callID, statusCont)
                StatusCollection(callID).statusMsg = "" ' need this to prevent object not set errors in checkCache
                targetRange = ToRange(targetRange)
                ExcelAsyncUtil.QueueAsMacro(Sub()
                                                DBSetQueryAction(callID, Query, targetRange, ConnString, caller, targetRangeName)
                                            End Sub)
            End If
        Catch ex As Exception
            LogWarn(ex.Message + ", callID: " + callID)
            DBSetQuery = EnvPrefix + ", Error (" + ex.Message + ") in DBSetQuery, callID: " + callID
        End Try
    End Function

    ''' <summary>set Query parameters (query text and connection string) of Query List or pivot table (incl. chart)</summary>
    ''' <param name="callID">the key for the statusMsg container</param>
    ''' <param name="Query">query for getting data</param>
    ''' <param name="targetRange">Range with Object beneath to put the Query/ConnString into</param>
    ''' <param name="ConnString">connection string defining DB, user, etc...</param>
    ''' <param name="caller">calling range passed by Action procedure</param>
    ''' <param name="targetRangeName"></param>
    Sub DBSetQueryAction(callID As String, Query As String, targetRange As Excel.Range, ConnString As String, caller As Excel.Range, targetRangeName As String)
        Dim targetSH As Excel.Worksheet
        Dim targetWB As Excel.Workbook
        Dim thePivotTable As Excel.PivotTable = Nothing
        Dim theListObject As Excel.ListObject = Nothing
        Dim errMsg As String

        Dim calcMode = ExcelDnaUtil.Application.Calculation
        Try
            ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual
        Catch ex As Exception : End Try
        ' this works around the data validation input bug and being called when COM Model is not ready
        ' when selecting a value from a list of a validated field or being invoked from a hyperlink (e.g. word), excel won't react to
        ' Application.Calculation changes, so just leave here...
        If ExcelDnaUtil.Application.Calculation <> Excel.XlCalculation.xlCalculationManual Then
            errMsg = "Error in setting Application.Calculation to Manual"
            GoTo err
        End If

        Dim callerFormula As String
        Try
            targetSH = targetRange.Parent
            targetWB = targetRange.Parent.Parent
            callerFormula = caller.Formula.ToString()
        Catch ex As Exception
            errMsg = "Error in getting targetSH/targetWB from TargetCell or getting formula"
            GoTo err
        End Try

        Dim srcExtent As String = "", targetExtent As String = ""
        errMsg = setExtents(caller, srcExtent, targetExtent)
        If errMsg <> "" Then GoTo err

        ' try to get either a pivot table object or a list object from the target cell. What we have is checked later...
        Try : thePivotTable = targetRange.PivotTable : Catch ex As Exception : End Try
        Try : theListObject = targetRange.ListObject : Catch ex As Exception : End Try

        Dim connType As String = ""
        Dim bgQuery As Boolean
        DBModifs.preventChangeWhileFetching = True
        Try
            Dim thePivotCache As Excel.PivotCache = Nothing
            Dim theQueryTable As Excel.QueryTable = Nothing
            ' first, get the connection type from the underlying PivotCache or QueryTable (OLEDB or ODBC)
            If thePivotTable IsNot Nothing Then
                thePivotCache = thePivotTable.PivotCache
                Try
                    connType = Left$(thePivotCache.Connection.ToString(), InStr(1, thePivotCache.Connection.ToString(), ";"))
                Catch ex As Exception
                    errMsg = "couldn't get connection from Pivot Table, please create Pivot Table with external data source, " + ex.Message
                    GoTo err
                End Try
            End If
            If theListObject IsNot Nothing Then
                theQueryTable = theListObject.QueryTable
                Try
                    connType = Left$(theQueryTable.Connection.ToString(), InStr(1, theQueryTable.Connection.ToString(), ";"))
                Catch ex As Exception
                    errMsg = "couldn't get connection from ListObject, please create ListObject with external data source, " + ex.Message
                    GoTo err
                End Try
            End If

            ' if we haven't already set the connection type in the alternative connection string then set it now..
            If Left(ConnString, 6) <> "OLEDB;" And Left(ConnString, 5) <> "ODBC;" Then ConnString = connType + ConnString
            ' now set the connection string and the query and refresh it.
            If thePivotTable IsNot Nothing Then
                bgQuery = thePivotCache.BackgroundQuery
                thePivotCache.Connection = ConnString
                thePivotCache.CommandType = Excel.XlCmdType.xlCmdSql
                thePivotCache.CommandText = Query
                thePivotCache.BackgroundQuery = False
                thePivotCache.Refresh()
                StatusCollection(callID).statusMsg = "Set " + connType + " PivotTable to (bgQuery= " + bgQuery.ToString() + "): " + Query
                thePivotCache.BackgroundQuery = bgQuery
                ' give hidden name to target range of pivot query (jump function)
                Dim theTableRange As Excel.Range = thePivotTable.TableRange1
                theTableRange.Name = targetExtent
                theTableRange.Parent.Parent.Names(targetExtent).Visible = False
            ElseIf theListObject IsNot Nothing Then
                bgQuery = theQueryTable.BackgroundQuery
                ' check whether target range is actually a table List object reference, if so, replace with simple address as this doesn't produce a #REF! error on QueryTable.Refresh
                ' this simple address is below being set to caller.Formula
                If InStr(targetRangeName, theListObject.Name) > 0 Then callerFormula = Replace(callerFormula, targetRangeName, Replace(targetRange.Cells(1, 1).Address, "$", ""))
                ' in case list object is sorted externally, give a warning (otherwise this leads to confusion when trying to order in the query)...
                If theListObject.Sort.SortFields.Count > 0 Then UserMsg("List Object " + theListObject.Name + " set by DBSetQuery in " + callID + " is already sorted by Excel, ordering statements in the query don't have any effect !",, MsgBoxStyle.Exclamation)
                ' in case of CUDFlags, reset them now (before resizing)...
                Dim dbMapperRangeName As String = getDBModifNameFromRange(targetRange)
                If Left(dbMapperRangeName, 8) = "DBMapper" Then
                    Dim dbMapper As DBMapper = DBModifDefColl("DBMapper").Item(dbMapperRangeName)
                    dbMapper.resetCUDFlags()
                End If
                theQueryTable.Connection = ConnString
                theQueryTable.CommandType = Excel.XlCmdType.xlCmdSql
                theQueryTable.CommandText = Query
                theQueryTable.BackgroundQuery = False
                Dim theRefreshStyle As Excel.XlCellInsertionMode = theQueryTable.RefreshStyle
                Dim thePreserveColumnInfo As Boolean = theQueryTable.PreserveColumnInfo
                Dim warning As String = ""
                Try
                    theQueryTable.Refresh()
                Catch ex As Exception
                    warning = "workaround taken for table size changed and side-effects on other data tables: RefreshStyle set to InsertEntireRows and PreserveColumnInfo set to False,"
                    ' this fixes two errors with query tables where the table size was changed: 8000A03EC and out of memory error
                    theQueryTable.RefreshStyle = Excel.XlCellInsertionMode.xlInsertEntireRows
                    theQueryTable.PreserveColumnInfo = False
                    Try
                        theQueryTable.Refresh()
                    Catch ex1 As Exception
                        Throw New Exception("Error in query table refresh: " + ex1.Message + "(after retrying with RefreshStyle=InsertEntireRows and PreserveColumnInfo=False)")
                    Finally
                        theQueryTable.RefreshStyle = theRefreshStyle
                        theQueryTable.PreserveColumnInfo = thePreserveColumnInfo
                    End Try
                End Try
                StatusCollection(callID).statusMsg = "Set " + connType + " ListObject to (bgQuery= " + bgQuery.ToString() + "), " + warning + If(theQueryTable.FetchedRowOverflow, "too many rows fetched to display,", "") + ": " + Query
                theQueryTable.BackgroundQuery = bgQuery
                Try
                    Dim testTarget = targetRange.Address
                Catch ex As Exception
                    caller.Formula = callerFormula ' restore formula as excel deletes target range when changing query fundamentally
                End Try
                ' give hidden name to target range of list object (jump function)
                Dim oldRange As Excel.Range = Nothing
                ' first invocation of DBSetQuery will have no defined targetExtent Range name, so this will fail:
                Try : oldRange = theListObject.Range.Parent.Parent.Names(targetExtent).RefersToRange : Catch ex As Exception : End Try
                If oldRange Is Nothing Then oldRange = theListObject.Range
                theListObject.Range.Name = targetExtent
                theListObject.Range.Parent.Parent.Names(targetExtent).Visible = False
                ' if refreshed range is a DBMapper and it is in the current workbook, resize it, but ONLY if it the DBMapper is the same area as the old range
                resizeDBMapperRange(theListObject.Range, oldRange)
            Else         ' neither PivotTable or ListObject could be found in TargetCell
                errMsg = "No PivotTable or ListObject with external data connection could be found in TargetRange " + targetRange.Address
                GoTo err
            End If
        Catch ex As Exception
            errMsg = ex.Message + " in query: " + Query
            GoTo err
        End Try
        DBModifs.preventChangeWhileFetching = False
        setCalcModeBack(calcMode)
        Exit Sub
err:
        setCalcModeBack(calcMode)
        LogWarn(errMsg + " caller: " + callID)
        If StatusCollection.ContainsKey(callID) Then StatusCollection(callID).statusMsg = errMsg
        DBModifs.preventChangeWhileFetching = False
        ' trigger recalculation to return error message to calling function
        Try : caller.Formula += " " : Catch ex As Exception : End Try
    End Sub

    ''' <summary>Stores a query into an powerquery defined by queryName</summary>
    ''' <param name="Query">(power) query for getting data</param>
    ''' <param name="queryName">powerquery name</param>
    ''' <returns>Status Message</returns>
    <ExcelFunction(Description:="Stores a query into a power query object defined in queryName")>
    Public Function DBSetPowerQuery(<ExcelArgument(Description:="query for getting data")> Query As Object,
                                    <ExcelArgument(Description:="Name of Powerquery where query should be set")> queryName As Object) As String
        Dim callID As String = ""
        Dim caller As Excel.Range
        Dim EnvPrefix As String = ""
        If ExcelDnaUtil.IsInFunctionWizard() Then Return "invoked from function wizard..."
        If checkPreventRefreshFlag() Then Return "preventRefresh set, DB Function not refreshing..."
        Try
            caller = ToRange(XlCall.Excel(XlCall.xlfCaller))
            ' calcContainers are identified by workbook name + Sheet name + function caller cell Address
            callID = "[" + caller.Parent.Parent.Name + "]" + caller.Parent.Name + "!" + caller.Address
            ' check query, also converts query to string (if it is a range)
            ' error message or cached status message is returned from checkParamsAndCache, if query OK and result was not already calculated (cached) then empty string
            DBSetPowerQuery = checkParamsAndCache(Query, callID, "")
            If DBSetPowerQuery.Length > 0 Then Exit Function

            ' first call: actually perform query
            If Not StatusCollection.ContainsKey(callID) Then
                Dim statusCont As New ContainedStatusMsg
                StatusCollection.Add(callID, statusCont)
                StatusCollection(callID).statusMsg = "" ' need this to prevent object not set errors in checkCache
                ExcelAsyncUtil.QueueAsMacro(Sub()
                                                DBSetPowerQueryAction(callID, Query, caller, queryName)
                                            End Sub)
            End If
        Catch ex As Exception
            LogWarn(ex.Message + ", callID: " + callID)
            DBSetPowerQuery = "Error (" + ex.Message + ") in DBSetPowerQuery, callID: " + callID
        End Try
    End Function

    Public avoidRequeryDuringEdit As Boolean = False
    Public queryBackupColl As New Dictionary(Of String, String)

    ''' <summary>set Query parameters (query text and connection string) of Query List or pivot table (incl. chart)</summary>
    ''' <param name="callID">the key for the statusMsg container</param>
    ''' <param name="Query">(power) query for getting data</param>
    ''' <param name="caller">calling range passed by Action procedure</param>
    ''' <param name="queryName">Name of Powerquery where query should be set</param>
    Sub DBSetPowerQueryAction(callID As String, Query As String, caller As Excel.Range, queryName As String)
        Dim targetWB As Excel.Workbook = caller.Parent.Parent
        Dim errMsg As String
        If avoidRequeryDuringEdit Then Exit Sub
        Dim calcMode = ExcelDnaUtil.Application.Calculation
        Try
            ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual
        Catch ex As Exception : End Try

        Try
            StatusCollection(callID).statusMsg = ""
            queryBackupColl(queryName) = targetWB.Queries(queryName).Formula
            ' set the query
            targetWB.Queries(queryName).Formula = Query
            ' refresh all connections where the query is used
            For Each wbConn As Excel.WorkbookConnection In targetWB.Connections
                Dim connectionInfo As String = ""
                Try : connectionInfo = wbConn.OLEDBConnection.Connection : Catch ex As Exception : End Try
                If InStr(LCase(connectionInfo), "location=" + LCase(queryName)) > 0 Then wbConn.Refresh()
            Next
            StatusCollection(callID).statusMsg = "set and refreshed " + queryName
        Catch ex As Exception
            errMsg = ex.Message + " in query: " + Query
            GoTo err
        End Try
        setCalcModeBack(calcMode)
        ' trigger recalculation to return message to calling function
        Try : caller.Formula += " " : Catch ex As Exception : End Try
        Exit Sub
err:
        setCalcModeBack(calcMode)
        LogWarn(errMsg + " caller: " + callID)
        If StatusCollection.ContainsKey(callID) Then StatusCollection(callID).statusMsg = errMsg
        ' trigger recalculation to return error message to calling function
        Try : caller.Formula += " " : Catch ex As Exception : End Try
    End Sub

    ''' <summary>common for DBListFetch, DBRowFetch and DBSetQuery Action procedures, setting the Extent Names at the beginning</summary>
    ''' <param name="caller">calling range passed by Action procedure</param>
    ''' <param name="srcExtent">by ref returned source range extent (place of db function) name</param>
    ''' <param name="targetExtent">by ref returned target range extent (place of results) name</param>
    ''' <param name="targetExtentF">by ref returned target formula range extent (place of formulas automatically extended with data) name</param>
    ''' <returns>error message in case of error, or empty if none</returns>
    Private Function setExtents(caller As Excel.Range, ByRef srcExtent As String, ByRef targetExtent As String, Optional ByRef targetExtentF As String = "") As String
        On Error Resume Next
        srcExtent = caller.Name.Name
        If Err.Number <> 0 Or InStr(1, srcExtent, "DBFsource") = 0 Then
            Err.Clear()
            srcExtent = "DBFsource" + Replace(Guid.NewGuid().ToString(), "-", "")
            caller.Name = srcExtent
            ' all db source and target names are workbook names
            caller.Parent.Parent.Names(srcExtent).Visible = False
            If Err.Number <> 0 Then Return "Error in setting srcExtent name: " + Err.Description
        End If
        caller.Parent.Parent.Names(srcExtent).Visible = False
        targetExtent = Replace(srcExtent, "DBFsource", "DBFtarget")
        targetExtentF = Replace(srcExtent, "DBFsource", "DBFtargetF")
        Return ""
    End Function

    ''' <summary>common for DBListFetch and DBRowFetch Action procedures to finish, resetting anything (Cursor, calc mode, status bar, screen updating) that was set otherwise...</summary>
    ''' <param name="calcMode">reset calc mode to this</param>
    ''' <param name="callID">for logging purpose</param>
    ''' <param name="additionalLogInfo">for logging purpose</param>
    Private Sub finishAction(calcMode As Excel.XlCalculation, callID As String, Optional additionalLogInfo As String = "")
        LogInfo("callID: " + callID + If(additionalLogInfo <> "", ", additionalInfo: " + additionalLogInfo, ""))
        DBModifs.preventChangeWhileFetching = False
        ' To return cursor to normal
        Try : ExcelDnaUtil.Application.Cursor = Excel.XlMousePointer.xlDefault : Catch ex As Exception : End Try
        Try : ExcelDnaUtil.Application.StatusBar = False : Catch ex As Exception : End Try
        ' coming from refresh, this might be off for dirtying "foreign" data targets (as we're on a different sheet than the calling function) 
        setCalcModeBack(calcMode)
    End Sub

    ''' <summary>set calculation mode back to calcMode</summary>
    ''' <param name="calcMode">calculation mode</param>
    Private Sub setCalcModeBack(calcMode As Excel.XlCalculation)
        Try : ExcelDnaUtil.Application.Calculation = calcMode
        Catch ex As Exception
            UserMsg("Error when setting back calculation mode to " + IIf(calcMode = -4135, "manual", IIf(calcMode = -4105, "automatic", IIf(calcMode = 2, "semiautomatic", "unknown: " + calcMode.ToString()))) + " (calculation property errors can occur for workbooks being opened through MS-office hyper-links), it is recommended to refresh the DB function to get valid results!",, MsgBoxStyle.Exclamation)
        End Try
    End Sub

    ''' <summary>check the preventRefreshFlag, either global or workbook level</summary>
    ''' <returns>true if set, workbook has precedence</returns>
    Private Function checkPreventRefreshFlag() As Boolean
        If preventRefreshFlagColl.ContainsKey(ExcelDnaUtil.Application.ActiveWorkbook.Name) Then
            Return preventRefreshFlagColl(ExcelDnaUtil.Application.ActiveWorkbook.Name)
        Else
            Return preventRefreshFlag
        End If
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
    ''' <param name="AutoFit">should columns be auto fitted ?</param>
    ''' <param name="autoformat">should 1st row formats be auto filled down?</param>
    ''' <param name="ShowRowNums">should row numbers be displayed in 1st column?</param>
    ''' <returns>Status Message, data values are returned outside of function cell (@see DBFuncEventHandler)</returns>
    <ExcelFunction(Description:="Fetches a list of data defined by query into TargetRange. Optionally copy formulas contained in FormulaRange, extend list depending on ExtendDataArea (0(default) = overwrite, 1=insert Cells, 2=insert Rows) and add field headers if HeaderInfo = TRUE")>
    Public Function DBListFetch(<ExcelArgument(Description:="query for getting data")> Query As Object,
                                <ExcelArgument(Description:="connection string defining DB, user, etc...")> ConnString As Object,
                                <ExcelArgument(Description:="Range to put the data into", AllowReference:=True)> targetRange As Object,
                                <ExcelArgument(Description:="Range to copy formulas down from", AllowReference:=True)> Optional formulaRange As Object = Nothing,
                                <ExcelArgument(Description:="how to deal with extending List Area")> Optional extendDataArea As Integer = 0,
                                <ExcelArgument(Description:="should headers be included in list")> Optional HeaderInfo As Object = Nothing,
                                <ExcelArgument(Description:="should columns be auto-fitted ?")> Optional AutoFit As Object = Nothing,
                                <ExcelArgument(Description:="should 1st row formats be auto-filled down?")> Optional autoformat As Object = Nothing,
                                <ExcelArgument(Description:="should row numbers be displayed in 1st column?")> Optional ShowRowNums As Object = Nothing) As String
        Dim callID As String = ""
        Dim EnvPrefix As String = ""
        If ExcelDnaUtil.IsInFunctionWizard() Then Return "invoked from function wizard..."
        If checkPreventRefreshFlag() Then Return "preventRefresh set, DB Function not refreshing..."
        Try
            DBListFetch = checkQueryAndTarget(Query, targetRange)
            If DBListFetch.Length > 0 Then Exit Function
            Dim caller As Excel.Range = ToRange(XlCall.Excel(XlCall.xlfCaller))
            resolveConnstring(ConnString, EnvPrefix, False)
            ' calcContainers are identified by workbook name + Sheet name + function caller cell Address
            callID = "[" + caller.Parent.Parent.Name + "]" + caller.Parent.Name + "!" + caller.Address
            ' prepare information for action procedure
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
                Dim statusCont As New ContainedStatusMsg
                StatusCollection.Add(callID, statusCont)
                targetRange = ToRange(targetRange)
                formulaRange = ToRange(formulaRange)
                ExcelAsyncUtil.QueueAsMacro(Sub()
                                                DBListFetchAction(callID, CStr(Query), caller, targetRange, CStr(ConnString), formulaRange, extendDataArea, HeaderInfo, AutoFit, autoformat, ShowRowNums, targetRangeName, formulaRangeName)
                                            End Sub)
            End If
        Catch ex As Exception
            LogWarn(ex.Message + ", callID : " + callID)
            DBListFetch = EnvPrefix + ", Error (" + ex.Message + ") in DBListFetch, callID : " + callID
        End Try
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
        Dim errMsg As String
        Dim calcMode = ExcelDnaUtil.Application.Calculation
        Try
            ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual
        Catch ex As Exception
            ' this works around the data validation input bug and being called when COM Model is not ready
            ' when selecting a value from a list of a validated field or being invoked from a hyperlink (e.g. word), excel won't react to
            ' Application.Calculation changes, so just leave here...
            errMsg = "Error in setting Application.Calculation to Manual: " + ex.Message + " in query: " + Query
            GoTo err
        End Try
        Dim warning As String = ""
        Try
            If ExcelDnaUtil.Application.Calculation <> Excel.XlCalculation.xlCalculationManual Then
                errMsg = "Error in setting Application.Calculation to Manual in query: " + Query
                GoTo err
            End If
            ExcelDnaUtil.Application.Cursor = Excel.XlMousePointer.xlWait  ' show the hourglass
            Dim targetSH As Excel.Worksheet = targetRange.Parent

            Dim srcExtent As String = "", targetExtent As String = "", targetExtentF As String = ""
            errMsg = setExtents(caller, srcExtent, targetExtent, targetExtentF)
            If errMsg <> "" Then
                errMsg += " in query: " + Query
                GoTo err
            End If
            Dim callerFormula As String = caller.Formula.ToString()

            Dim theTargetQueryTable As Object = Nothing
            Try : theTargetQueryTable = targetRange.QueryTable : Catch ex As Exception : End Try
            ' size for existing named range used for resizing the data area (old extent)
            Dim oldCols As Integer = 0, oldRows As Integer = 0
            Try
                oldCols = targetSH.Parent.Names(targetExtent).RefersToRange.Columns.Count
                oldRows = targetSH.Parent.Names(targetExtent).RefersToRange.Rows.Count
                ' legacy dblistfetch (adodb): this is needed to clear old data
                If IsNothing(theTargetQueryTable) OrElse theTargetQueryTable.ResultRange.Address <> targetSH.Parent.Names(targetExtent).RefersToRange.Address Then targetSH.Parent.Names(targetExtent).RefersToRange.ClearContents()
            Catch ex As Exception : End Try

            Dim startRow, startCol As Integer
            Try
                ' don't use targetRange here as it doesn't change during calculations that shift their address. The names do, however.
                startRow = targetSH.Parent.Names(targetExtent).RefersToRange.Row
                startCol = targetSH.Parent.Names(targetExtent).RefersToRange.Column
            Catch ex As Exception
                Try
                    startRow = targetRange.Row
                    startCol = targetRange.Column
                Catch ex2 As Exception
                    errMsg = "Error in getting startRow/startCol of target range: " + ex2.Message + " in query: " + Query
                    GoTo err
                End Try
            End Try

            Dim formulaSH As Excel.Worksheet = Nothing
            Dim formulaStart As Integer
            Dim additionalFormulaColumns As Integer = 0
            If formulaRange IsNot Nothing Then
                formulaSH = formulaRange.Parent
                ' only first row of formulaRange is important, rest will be auto filled down (actually this is needed to make the auto format work)
                formulaRange = formulaRange.Rows(1)
                formulaStart = formulaRange.Row
                ' if formulas are adjacent to data extend total range to formula range ! total range is used to extend DBMappers defined under the DB Function target...
                If targetSH Is formulaSH And formulaRange.Column = startCol + oldCols Then additionalFormulaColumns = formulaRange.Columns.Count
            End If

            DBModifs.preventChangeWhileFetching = True
            ' used for resizing potential DBMapper under DBListfetch TargetRange
            Dim oldTotalTargetRange As Excel.Range = Nothing
            Try : oldTotalTargetRange = targetSH.Range(targetSH.Cells(startRow, startCol), targetSH.Cells(startRow + oldRows - 1, startCol + oldCols + additionalFormulaColumns - 1)) : Catch ex As Exception : End Try

            ' clear old formulas
            Dim oldFRows, oldFCols As Integer
            If formulaSH IsNot Nothing Then
                Try
                    oldFRows = formulaSH.Parent.Names(targetExtentF).RefersToRange.Rows.Count
                    oldFCols = formulaSH.Parent.Names(targetExtentF).RefersToRange.Columns.Count
                Catch ex As Exception : End Try
                If oldFRows > 2 Then
                    Try
                        formulaSH.Range(formulaSH.Cells(formulaStart + 1, formulaRange.Column), formulaSH.Cells(formulaRange.Row + oldFRows - 1, formulaRange.Column + oldFCols - 1)).ClearContents()
                    Catch ex As Exception
                        errMsg = "Error in clearing old data for formula range: (" + ex.Message + ") in query: " + Query
                        GoTo err
                    End Try
                End If
            End If

            ' for oledb drivers add OLEDB; in front, excel ms query needs that!
            If InStr(1, UCase$(ConnString), "OLEDB") > 0 AndAlso Left(UCase$(ConnString), 5) <> "ODBC;" AndAlso Left(UCase$(ConnString), 6) <> "OLEDB;" Then ConnString = "OLEDB;" + ConnString

            ' from now on we don't propagate any errors as data is modified in sheet....
            ExcelDnaUtil.Application.StatusBar = "Displaying data for DBList: " + If(targetRangeName.Length > 0, targetRangeName, targetSH.Name + "!" + targetRange.Address)

            ' auto format: store 1st rows formats range to apply afterwards to whole column(s), this is needed as formats are reset by the query refresh
            Dim NumFormat() As String = Nothing, NumFormatF() As String = Nothing, FormulasF() As String = Nothing
            If autoformat Then
                For i As Integer = 0 To oldCols
                    ReDim Preserve NumFormat(i)
                    NumFormat(i) = targetSH.Cells(startRow + If(HeaderInfo, 1, 0), startCol + i).NumberFormatLocal
                Next
            End If
            ' now for the calculated data area
            If formulaRange IsNot Nothing Then
                For i = 0 To formulaRange.Columns.Count - 1
                    If autoformat Then
                        ReDim Preserve NumFormatF(i)
                        ' why NumberFormatLocal? Because NumberFormat is not working consistently. https://www.add-in-express.com/forum/read.php?FID=5&TID=8333
                        NumFormatF(i) = formulaSH.Cells(formulaStart, formulaRange.Column + i).NumberFormatLocal
                    End If
                    ' also preserve the formulas for potential restoring after refresh (in case of extendArea = 1 (Excel.XlCellInsertionMode.xlInsertDeleteCells) excel shifts down the target range, resulting in formula modification)
                    ReDim Preserve FormulasF(i)
                    FormulasF(i) = formulaSH.Cells(formulaStart, formulaRange.Column + i).Formula
                Next
            End If

            ' check if formulaRange and targetRange overlap !
            Dim possibleIntersection As Excel.Range = Nothing
            Try
                possibleIntersection = ExcelDnaUtil.Application.Intersect(formulaRange, targetSH.Range(targetRange.Cells(1, 1), targetRange.Cells(1, 1).Offset(IIf(oldRows = 0, 1, oldRows) - 1, IIf(oldCols = 0, 1, oldCols) - 1)))
            Catch ex As Exception : End Try
            If possibleIntersection IsNot Nothing Then
                warning += ", formula range and target range intersect (" + targetSH.Name + "!" + possibleIntersection.Address + "), formula copying disabled !!"
                formulaRange = Nothing
            End If

            Dim resultingQueryRange As Excel.Range : Dim targetQueryTableExist As Boolean = False
            If IsNothing(theTargetQueryTable) Then
                ' no underlying query table yet, add one
                Try
                    theTargetQueryTable = targetSH.QueryTables.Add(Connection:=ConnString, Destination:=targetRange)
                Catch ex As Exception
                    errMsg = "Error in adding query table: " + ex.Message + ", query: " + Query
                    GoTo err
                End Try
                extendArea = 0 ' this is required to prevent "right" shifting of cells at the beginning if no QueryTable exists yet!
            Else
                targetQueryTableExist = True
            End If
            With theTargetQueryTable
                ' now fill in the data from the query
                Try
                    ' provide additional encryption information to connection string if required
                    .Connection = ConnString + IIf(InStr(ConnString.ToLower(), "encrypt=yes") > 0, ";Use Encryption for Data=True", "")
                Catch ex As Exception
                    errMsg = IIf(targetQueryTableExist, "Probably the connection was deleted for the query table (you can reset this by removing the query definition of the external data range): ", "Error in setting connection string for QueryTable: ") + ex.Message + " in query: " + Query
                    GoTo err
                End Try
                Try
                    .CommandText = Query
                Catch ex As Exception
                    errMsg = "Error in setting query for query table: " + ex.Message + ", query: " + Query
                    GoTo err
                End Try
                Dim rowNumberHeader As String = "" ' only store row number header if row numbers are already in place (otherwise we would get the false header before the extension)
                If ShowRowNumbers And .RowNumbers Then rowNumberHeader = targetRange.Cells(1, 1).Value
                Try
                    .FieldNames = HeaderInfo
                    .RowNumbers = ShowRowNumbers
                    .AdjustColumnWidth = AutoFit
                    .BackgroundQuery = False
                    .RefreshStyle = IIf(extendArea = 0, Excel.XlCellInsertionMode.xlOverwriteCells, IIf(extendArea = 1, Excel.XlCellInsertionMode.xlInsertDeleteCells, Excel.XlCellInsertionMode.xlInsertEntireRows))
                Catch ex As Exception
                    errMsg = "Error in setting parameters for query table: " + ex.Message + ", query: " + Query
                    GoTo err
                End Try
                Try
                    .Refresh()
                    If .FetchedRowOverflow Then
                        warning += "row count of returned data exceeds max row of excel: start row:" + targetRange.Row.ToString() + " + row count:" + .Recordset.Count.ToString() + " > max row+1:" + (targetRange.EntireColumn.Rows.Count + 1).ToString()
                    End If
                Catch ex As Exception
                    errMsg = "Error in refreshing query table: " + ex.Message + ", query: " + Query
                    GoTo err
                End Try
                Try
                    resultingQueryRange = .ResultRange
                Catch ex As Exception
                    errMsg = "Error in getting resulting range for query table: " + ex.Message + ", query: " + Query
                    GoTo err
                End Try
                ' reset row number header (deleted by refresh)
                If ShowRowNumbers Then targetRange.Cells(1, 1).Value = rowNumberHeader
            End With
            ' in case of extendArea = 1 (Excel.XlCellInsertionMode.xlInsertDeleteCells) excel shifts down the target range, this resets the caller formula and the calculated area back
            If caller.Formula <> callerFormula Then caller.Formula = callerFormula
            If formulaRange IsNot Nothing Then
                For i = 0 To formulaRange.Columns.Count - 1
                    If FormulasF(i) <> formulaSH.Cells(formulaStart, formulaRange.Column + i).Formula Then formulaSH.Cells(formulaStart, formulaRange.Column + i).Formula = FormulasF(i)
                Next
            End If

            ' get new query area dimensions
            Dim qryRows, qryCols As Integer
            Try
                qryRows = resultingQueryRange.Rows.Count
                qryCols = resultingQueryRange.Columns.Count
            Catch ex As Exception
                errMsg = "Error in getting Rows and Columns from query table result range: " + ex.Message + ", query: " + Query
                GoTo err
            End Try

            '''' formulas recreation (removal and auto fill again)
            Dim formulaFilledRange As Excel.Range = Nothing
            Dim copyDownLastRow As Integer
            Dim resultRows As Integer = qryRows - If(HeaderInfo, 1, 0)
            If formulaRange IsNot Nothing Then
                With formulaRange
                    Dim FCols As Integer = .Columns.Count
                    Dim prevRows As Integer = oldRows - If(HeaderInfo, 1, 0)
                    ' only shift formula range if old data existed
                    If oldRows > If(HeaderInfo, 1, 0) Then
                        Try
                            If oldRows < qryRows Then
                                ' either cells/rows are shifted down (old data area was smaller than current) ...
                                If extendArea = 1 Then
                                    formulaSH.Range(formulaSH.Cells(.Row + prevRows, .Column), formulaSH.Cells(.Row + resultRows - 1, .Column + FCols - 1)).Insert(Shift:=Excel.XlDirection.xlDown)
                                ElseIf extendArea = 2 Then
                                    ' take care not to insert twice (if we're having formulas in the same sheet)
                                    If targetSH IsNot formulaSH Then formulaSH.Rows((.Row + prevRows).ToString() + ":" + (.Row + resultRows - 1).ToString()).Insert(Shift:=Excel.XlDirection.xlDown)
                                End If
                                'else extendArea = 0: just overwrite -> no special action
                            ElseIf oldRows > qryRows Then
                                ' .... or cells/rows are shifted up (old data area was larger than current)
                                If extendArea = 1 Then
                                    formulaSH.Range(formulaSH.Cells(.Row + resultRows, .Column), formulaSH.Cells(.Row + prevRows - 1, .Column + FCols - 1)).Delete(Shift:=Excel.XlDirection.xlUp)
                                ElseIf extendArea = 2 Then
                                    ' take care not to delete twice (if we're having formulas in the same sheet)
                                    If targetSH IsNot formulaSH Then formulaSH.Rows((.Row + resultRows).ToString() + ":" + (.Row + prevRows - 1).ToString()).Delete(Shift:=Excel.XlDirection.xlUp)
                                End If
                                'else extendArea = 0: just overwrite -> no special action
                            End If
                        Catch ex As Exception
                            errMsg = "Error in resizing formula range: " + ex.Message + " in query: " + Query
                            GoTo err
                        End Try
                    End If
                    ' determine bottom of formula range
                    ' check for excels boundaries !!
                    If .Row + resultRows > .EntireColumn.Rows.Count Then
                        warning += ", formulas would exceed max row of excel: start row:" + .Row.ToString() + " + row count:" + resultRows.ToString() + " > max row:" + (.EntireColumn.Rows.Count).ToString()
                        copyDownLastRow = .EntireColumn.Rows.Count
                    Else
                        'the normal end of our auto filled rows = formula start + list size
                        copyDownLastRow = .Row + resultRows
                    End If
                    Try
                        ' sanity check to avoid exception (fill same row/upwards)!
                        If copyDownLastRow - 1 > .Row Then formulaSH.Range(formulaSH.Cells(.Row, .Column), formulaSH.Cells(.Row, .Column + FCols - 1)).AutoFill(Destination:=formulaSH.Range(formulaSH.Cells(.Row, .Column), formulaSH.Cells(copyDownLastRow - 1, .Column + FCols - 1)))
                        formulaFilledRange = formulaSH.Range(formulaSH.Cells(.Row, .Column), formulaSH.Cells(copyDownLastRow - 1, .Column + FCols - 1))
                    Catch ex As Exception
                        errMsg = "Error filling down formulas: " + ex.Message + ", query: " + Query
                        GoTo err
                    End Try

                    ' delete the name to have a "clean" name area (otherwise visible = false setting wont work for dataTargetRange)
                    Try : formulaFilledRange.Parent.Parent.Names(targetExtentF).Delete : Catch ex As Exception : End Try ' might not exist, so ignore errors here...
                    Try
                        ' reassign internal name to changed formula area
                        formulaFilledRange.Name = targetExtentF
                        formulaFilledRange.Name.Visible = False
                        ' reassign visible defined name to changed formula area only if defined...
                        If formulaRangeName.Length > 0 Then
                            formulaFilledRange.Name = formulaRangeName    ' NOT USING formulaFilledRange.Name.Visible = True, or hidden range will also be visible...
                        End If
                    Catch ex As Exception
                        errMsg = "Error in (re)assigning formula range name: " + ex.Message + ", query: " + Query
                        GoTo err
                    End Try
                End With
            End If

            ' reassign name to changed data area
            ' set the new hidden targetExtent name...
            Dim newTargetRange, totalTargetRange As Excel.Range
            Try
                newTargetRange = targetSH.Range(targetSH.Cells(startRow, startCol), targetSH.Cells(startRow + qryRows - 1, startCol + qryCols - 1))
                ' delete the name to have a "clean" name area (otherwise visible = false setting wont work for dataTargetRange)
                newTargetRange.Name = targetExtent
                newTargetRange.Parent.Parent.Names(targetExtent).Visible = False
                totalTargetRange = newTargetRange
                If additionalFormulaColumns > 0 Then
                    totalTargetRange = targetSH.Range(targetSH.Cells(startRow, startCol), targetSH.Cells(startRow + qryRows - 1, startCol + qryCols - 1 + additionalFormulaColumns))
                End If
                ' (re)assign visible name for the total area, if given
                If targetRangeName.Length > 0 Then
                    totalTargetRange.Name = targetRangeName
                End If
                ' if refreshed range is a DBMapper and it is in the current workbook, resize it
                resizeDBMapperRange(totalTargetRange, oldTotalTargetRange)
            Catch ex As Exception
                errMsg = "Error in (re)assigning data target name: " + ex.Message + " (maybe known issue with 'cell like' sheet names, e.g. 'C701 country' ?), query: " + Query
                GoTo err
            End Try

            ' get the true returned record count (returned range is always at least one row, if headers are included subtract 1)
            Try : If ExcelDnaUtil.Application.WorksheetFunction.CountA(newTargetRange) = 0 And resultRows = 1 Then resultRows = 0
            Catch ex As Exception : End Try
            '''' any warnings, errors ?
            If warning.Length > 0 Then
                If InStr(1, warning, "Error:") = 0 And InStr(1, warning, "No Data") = 0 Then
                    If Left$(warning, 1) = "," Then
                        warning = Right$(warning, Len(warning) - 2)
                    End If
                    StatusCollection(callID).statusMsg = "Retrieved " + resultRows.ToString() + " record" + If(resultRows > 1 Or resultRows = 0, "s", "") + ", Warnings: " + warning
                Else
                    StatusCollection(callID).statusMsg = warning
                End If
            Else
                StatusCollection(callID).statusMsg = "Retrieved " + resultRows.ToString() + " record" + If(resultRows > 1 Or resultRows = 0, "s", "") + " from: " + Query
            End If
            ' auto format: restore formats
            Try
                If autoformat And NumFormat IsNot Nothing Then
                    For i = 0 To UBound(NumFormat)
                        newTargetRange.Columns(i + 1).NumberFormatLocal = NumFormat(i)
                    Next
                    ' also restore for formula(filled) range
                    If formulaRange IsNot Nothing And NumFormatF IsNot Nothing And formulaFilledRange IsNot Nothing Then
                        For i = 0 To UBound(NumFormatF)
                            formulaFilledRange.Columns(i + 1).NumberFormatLocal = NumFormatF(i)
                        Next
                    End If
                End If
            Catch ex As Exception
                errMsg = "Error in restoring formats: " + ex.Message + ", query: " + Query
                GoTo err
            End Try
            ' set complete header row bold
            Try
                If HeaderInfo Then totalTargetRange.Rows(1).Font.Bold = True
            Catch ex As Exception : End Try
            'auto fit columns AFTER auto format so we don't have problems with applied formats visibility
            Try
                If AutoFit Then
                    If formulaFilledRange IsNot Nothing And formulaFilledRange IsNot ExcelEmpty.Value Then
                        ' auto fit also formula(filled) range, pass it to AfterCalculation Event here
                        StatusCollection(callID).formulaRange = formulaFilledRange
                    End If
                    newTargetRange.Columns.AutoFit()
                    newTargetRange.Rows.AutoFit()
                End If
            Catch ex As Exception
                errMsg = "Error in auto fitting: " + ex.Message + ", query: " + Query
                GoTo err
            End Try
        Catch ex As Exception
            errMsg = "General Error in DBListFetchAction: " + ex.Message + ", query: " + Query
            GoTo err
        End Try
        finishAction(calcMode, callID)
        If warning.Length > 0 Then
            ' recalculate to trigger return of warning messages to calling function
            Try : caller.Formula += " " : Catch ex As Exception : End Try
        End If
        Exit Sub

err:    LogWarn(errMsg + ", caller: " + callID)
        If StatusCollection.ContainsKey(callID) Then StatusCollection(callID).statusMsg = errMsg
        finishAction(calcMode, callID, "Error")
        ' recalculate to trigger return of error messages to calling function
        Try : caller.Formula += " " : Catch ex As Exception : End Try
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
        Dim finalTargetArray() As Excel.Range = Nothing ' final target array that is passed to makeCalcMsgContainer (after removing header flag)
        Dim callID As String = ""
        Dim HeaderInfo As Boolean
        Dim EnvPrefix As String = ""
        If ExcelDnaUtil.IsInFunctionWizard() Then Return "invoked from function wizard..."
        If checkPreventRefreshFlag() Then Return "preventRefresh set, DB Function not refreshing..."
        Try
            DBRowFetch = checkQueryAndTarget(Query, targetArray)
            If DBRowFetch.Length > 0 Then Exit Function
            Dim caller As Excel.Range = ToRange(XlCall.Excel(XlCall.xlfCaller))
            resolveConnstring(ConnString, EnvPrefix, False)
            ' calcContainers are identified by workbook name + Sheet name + function caller cell Address 
            callID = "[" + caller.Parent.Parent.Name + "]" + caller.Parent.Name + "!" + caller.Address
            If dontCalcWhileClearing Then
                DBRowFetch = EnvPrefix + ", dontCalcWhileClearing = True !"
                Exit Function
            End If

            ' prepare information for action proc
            If TypeName(targetArray(0)) = "Boolean" Or TypeName(targetArray(0)) = "String" Then
                HeaderInfo = convertToBool(targetArray(0))
                For i As Integer = 1 To UBound(targetArray)
                    ReDim Preserve finalTargetArray(i - 1)
                    If IsNothing(ToRange(targetArray(i))) Then
                        DBRowFetch = EnvPrefix + ", Part " + i.ToString() + " of Target is not a valid Range !"
                        Exit Function
                    End If
                    finalTargetArray(i - 1) = ToRange(targetArray(i))
                Next
            ElseIf TypeName(targetArray(0)) = "ExcelEmpty" Or TypeName(targetArray(0)) = "ExcelError" Or TypeName(targetArray(0)) = "ExcelMissing" Then
                ' return appropriate error message...
                DBRowFetch = EnvPrefix + ", First argument (header) " + Replace(TypeName(targetArray(0)), "Excel", "") + " !"
                Exit Function
            Else
                For i = 0 To UBound(targetArray)
                    ReDim Preserve finalTargetArray(i)
                    If IsNothing(ToRange(targetArray(i))) Then
                        DBRowFetch = EnvPrefix + ", Part " + (i + 1).ToString() + " of Target is not a valid Range !"
                        Exit Function
                    End If
                    finalTargetArray(i) = ToRange(targetArray(i))
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
                Dim statusCont As New ContainedStatusMsg
                StatusCollection.Add(callID, statusCont)
                StatusCollection(callID).statusMsg = "" ' need this to prevent object not set errors in checkCache
                ExcelAsyncUtil.QueueAsMacro(Sub()
                                                DBRowFetchAction(callID, CStr(Query), caller, finalTargetArray, CStr(ConnString), HeaderInfo)
                                            End Sub)
            End If
        Catch ex As Exception
            LogWarn(ex.Message + ", callID: " + callID)
            DBRowFetch = EnvPrefix + ", Error (" + ex.Message + ") in DBRowFetch, callID: " + callID
        End Try
    End Function

    ''' <summary>Actually do the work for DBRowFetch: Query (assumed) one row of data, write it into targetCells</summary>
    ''' <param name="callID"></param>
    ''' <param name="Query"></param>
    ''' <param name="caller"></param>
    ''' <param name="targetArray"></param>
    ''' <param name="ConnString"></param>
    ''' <param name="HeaderInfo"></param>
    Public Sub DBRowFetchAction(callID As String, Query As String, caller As Excel.Range, targetArray As Object, ConnString As String, HeaderInfo As Boolean)
        Dim errMsg As String
        Dim targetSH As Excel.Worksheet

        StatusCollection(callID).statusMsg = ""
        Dim calcMode = ExcelDnaUtil.Application.Calculation
        Dim targetCells As Object = targetArray
        Try : targetSH = targetCells(0).Parent : Catch ex As Exception
            errMsg = "Error getting parent worksheet from targetCells" + ex.Message
            GoTo err
        End Try
        Try : ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual : Catch ex As Exception : End Try
        ' this works around the data validation input bug and being called when COM Model is not ready
        ' when selecting a value from a list of a validated field or being invoked from a hyperlink (e.g. word), excel won't react to
        ' Application.Calculation changes, so just leave here...
        If ExcelDnaUtil.Application.Calculation <> Excel.XlCalculation.xlCalculationManual Then
            errMsg = "Error in setting Application.Calculation to Manual: " + Err.Description + " in query: " + Query
            GoTo err
        End If
        ExcelDnaUtil.Application.Cursor = Excel.XlMousePointer.xlWait  ' show the hourglass
        ExcelDnaUtil.Application.StatusBar = "Retrieving data for DBRows: " + targetSH.Name + "!" + targetCells(0).Address

        Dim srcExtent As String = "", targetExtent As String = ""
        errMsg = setExtents(caller, srcExtent, targetExtent)
        If errMsg <> "" Then
            errMsg += " in query: " + Query
            GoTo err
        End If
        ' remove old data in case we changed the target range array
        Try : targetSH.Range(targetExtent).ClearContents() : Catch ex As Exception : End Try

        Try
            If Left(ConnString.ToUpper, 5) = "ODBC;" Then
                ' change to ODBC driver setting, if SQLOLEDB
                ConnString = Replace(ConnString, fetchSetting("ConnStringSearch" + env(), "provider=SQLOLEDB"), fetchSetting("ConnStringReplace" + env(), "driver=SQL SERVER"))
                ' remove "ODBC;"
                ConnString = Right(ConnString, ConnString.Length - 5)
                conn = New OdbcConnection(ConnString)
            ElseIf InStr(ConnString.ToLower, "provider=sqloledb") Or InStr(ConnString.ToLower, "driver=sql server") Then
                ' remove provider=SQLOLEDB; (or whatever is in ConnStringSearch<>) for sql server as this is not allowed for ado.net (e.g. from a connection string for MS Query/Office)
                ConnString = Replace(ConnString, fetchSetting("ConnStringSearch" + env(), "provider=SQLOLEDB") + ";", "")
                conn = New SqlConnection(ConnString)
            ElseIf InStr(ConnString.ToLower, "oledb") Then
                conn = New OleDbConnection(ConnString)
            Else
                ' try with odbc
                conn = New OdbcConnection(ConnString)
            End If
        Catch ex As Exception
            errMsg = "Error creating new connection: " + ex.Message + " for connection string: " + ConnString
            GoTo err
        End Try
        Try : conn.Open() : Catch ex As Exception
            errMsg = "Error opening connection: " + ex.Message + " for connection string: " + ConnString
            GoTo err
        End Try
        Dim sqlCommand As System.Data.Common.DbCommand
        Try
            If TypeName(conn) = "SqlConnection" Then
                sqlCommand = New SqlCommand(Query, conn)
            ElseIf TypeName(conn) = "OleDbConnection" Then
                sqlCommand = New OleDbCommand(Query, conn)
            Else
                sqlCommand = New OdbcCommand(Query, conn)
            End If
        Catch ex As Exception
            errMsg = "Error creating new sqlCommand: " + ex.Message + " for Query: " + Query
            GoTo err
        End Try
        Dim recordset As System.Data.Common.DbDataReader : Dim recordsetHasRows As Boolean : Dim returnedRows As Long = 0
        Try
            recordset = sqlCommand.ExecuteReader()
            recordsetHasRows = recordset.Read()
            If recordsetHasRows Then returnedRows = 1
        Catch ex As Exception
            errMsg = "Error executing sqlCommand: " + ex.Message + " for Query: " + Query
            GoTo err
        End Try

        DBModifs.preventChangeWhileFetching = True
        If Not recordsetHasRows Then StatusCollection(callID).statusMsg = "Warning: No Data returned in query: " + Query

        Dim totalFieldsDisplayed As Long = 0 ' needed to calculate displayedRows
        Try
            ' if "heading range" is present then orientation of first range (header) defines layout of data: if "heading range" is column then data is returned column-wise, else row by row.
            ' if there is just one block of data then it is assumed that there are usually more rows than columns and orientation is set by row/column size
            Dim fillByRows As Boolean = IIf(UBound(targetCells) > 0, targetCells(0).Rows.Count < targetCells(0).Columns.Count, targetCells(0).Rows.Count > targetCells(0).Columns.Count)
            ' put values (single record) from Recordset into targetCells
            Dim fieldIter As Integer = 0 ' iterating through recordset fields
            Dim rangeIter As Integer = 0 ' iterating through passed ranges
            Dim headerFilled As Boolean = Not HeaderInfo    ' if we don't need headers the assume they are filled already....
            Dim refCollector As Excel.Range = targetCells(0) ' needed to put together passed ranges to give dbftarget name to them
            Do
                Dim targetSlices As Excel.Range
                If fillByRows Then
                    targetSlices = targetCells(rangeIter).Rows
                Else
                    targetSlices = targetCells(rangeIter).Columns
                End If
                For Each targetSlice As Excel.Range In targetSlices
                    Dim aborted As Boolean = XlCall.Excel(XlCall.xlAbort) ' for long running actions, allow interruption
                    If aborted Then
                        errMsg = "data fetching interrupted by user !"
                        GoTo err
                    End If
                    For Each theCell As Excel.Range In targetSlice.Cells
                        If Not recordsetHasRows Then
                            theCell.Value = ""
                        Else
                            If Not headerFilled Then
                                theCell.Value = recordset.GetName(fieldIter)
                            Else
                                Try : theCell.Value = recordset.GetValue(fieldIter) : Catch ex As Exception
                                    errMsg += "Field '" + recordset.GetName(fieldIter) + "' caused following error: '" + Err.Description + "'" ' don't break operation, just collect message
                                End Try
                                totalFieldsDisplayed += 1
                            End If
                            If fieldIter = recordset.FieldCount - 1 Then
                                ' reached end of fields, get next data row
                                If headerFilled Then
                                    recordsetHasRows = recordset.Read()
                                    If recordsetHasRows Then returnedRows += 1
                                Else
                                    headerFilled = True
                                End If
                                fieldIter = -1 ' reset field iterator
                            End If
                        End If
                        fieldIter += 1
                    Next
                Next
                rangeIter += 1
                If Not rangeIter > UBound(targetCells) Then refCollector = ExcelDnaUtil.Application.Union(refCollector, targetCells(rangeIter))
            Loop Until rangeIter > UBound(targetCells)
            ' get rest of records for returned status message
            While recordset.Read()
                returnedRows += 1
            End While
            ' delete the name to have a "clean" name area (otherwise visible = false setting wont work for dataTargetRange)
            refCollector.Name = targetExtent
            refCollector.Parent.Parent.Names(targetExtent).Visible = False
        Catch ex As Exception
            errMsg = "Error in filling target range: " + ex.Message + ", query: " + Query
            GoTo err
        End Try

        If StatusCollection(callID).statusMsg.Length = 0 Then StatusCollection(callID).statusMsg = "Displayed " + Math.Ceiling(totalFieldsDisplayed / recordset.FieldCount).ToString() + " of " + returnedRows.ToString() + " record" + If(returnedRows > 1 Or returnedRows = 0, "s", "") + " from: " + Query + IIf(errMsg <> "", ";Errors: " + errMsg, "")
        finishAction(calcMode, callID)
        Exit Sub

err:    LogWarn(errMsg + ", caller: " + callID)
        If StatusCollection.ContainsKey(callID) Then StatusCollection(callID).statusMsg = errMsg
        finishAction(calcMode, callID, "Error")
        ' recalculate to trigger return of error messages to calling function
        Try : caller.Formula += " " : Catch ex As Exception : End Try
    End Sub

    ''' <summary>gets the documentation dictionary for the config dropdown (used by ConfigFiles.createConfigTreeMenu()). This is the basis for the documentation provided when clicking the config entries with Ctrl or Shift</summary>
    ''' <param name="ConfigDocQuery">query against current active environment for retrieving the data for the documentation. This query needs to return three fields for each table/view/field object: 1. database of the object (only really needed for tables/views), 2. table/view name (for fields this is the parent table/view) and 3. the documentation for the object. All this has to be ordered by table/view name, with the table/view object being first</param>
    ''' <returns>a dictionary of table/view keys and their associated collected documentation of the table/view and its fields. This is being collected by assuming the ordered way the query returns the data (see above)</returns>
    Public Function getConfigDocCollection(ConfigDocQuery As String) As Dictionary(Of String, String)
        getConfigDocCollection = New Dictionary(Of String, String)

        ' first get the connection as usual
        Dim ConnString As String = fetchSetting("ConstConnString" + env(), "")
        Try
            If Left(ConnString.ToUpper, 5) = "ODBC;" Then
                ' change to ODBC driver setting, if SQLOLEDB
                ConnString = Replace(ConnString, fetchSetting("ConnStringSearch" + env(), "provider=SQLOLEDB"), fetchSetting("ConnStringReplace" + env(), "driver=SQL SERVER"))
                ' remove "ODBC;"
                ConnString = Right(ConnString, ConnString.Length - 5)
                conn = New OdbcConnection(ConnString)
            ElseIf InStr(ConnString.ToLower, "provider=sqloledb") Or InStr(ConnString.ToLower, "driver=sql server") Then
                ' remove provider=SQLOLEDB; (or whatever is in ConnStringSearch<>) for sql server as this is not allowed for ado.net (e.g. from a connection string for MS Query/Office)
                ConnString = Replace(ConnString, fetchSetting("ConnStringSearch" + env(), "provider=SQLOLEDB") + ";", "")
                conn = New SqlConnection(ConnString)
            ElseIf InStr(ConnString.ToLower, "oledb") Then
                conn = New OleDbConnection(ConnString)
            Else
                ' try with odbc
                conn = New OdbcConnection(ConnString)
            End If
        Catch ex As Exception
            UserMsg("Error creating new connection: " + ex.Message + " for connection string: " + ConnString)
            GoTo err
        End Try
        Try : conn.Open() : Catch ex As Exception
            UserMsg("Error opening connection: " + ex.Message + " for connection string: " + ConnString)
            GoTo err
        End Try
        Dim sqlCommand As System.Data.Common.DbCommand
        Try
            If TypeName(conn) = "SqlConnection" Then
                sqlCommand = New SqlCommand(ConfigDocQuery, conn)
            ElseIf TypeName(conn) = "OleDbConnection" Then
                sqlCommand = New OleDbCommand(ConfigDocQuery, conn)
            Else
                sqlCommand = New OdbcCommand(ConfigDocQuery, conn)
            End If
        Catch ex As Exception
            UserMsg("Error creating new sqlCommand: " + ex.Message + " for ConfigDocQuery: " + vbCrLf + ConfigDocQuery)
            GoTo err
        End Try
        Dim recordset As System.Data.Common.DbDataReader
        Try
            recordset = sqlCommand.ExecuteReader()
        Catch ex As Exception
            UserMsg("Error executing sqlCommand: " + ex.Message + " for ConfigDocQuery: " + vbCrLf + ConfigDocQuery)
            GoTo err
        End Try

        ' character before DBname, Path of config files is assumed to be folder\sub-folder\sub-folder-etc\<charBeforeDBnameConfigDoc><DBName>\<Tablename>.xcl
        Dim charBeforeDBnameConfigDoc As String = fetchSetting("charBeforeDBnameConfigDoc", ".")

        ' then collect the documentation from the ConfigDocQuery into getConfigDocCollection
        Try
            Dim DBName As String = ""
            Dim ObjectName As String = ""
            Dim Documentation As String = ""
            Dim currDBName As String = ""
            Dim currObjectName As String = ""
            Dim ConfigDoc As String = ""
            While recordset.Read()
                'DBname1	ObjName1	DescDBname1ObjName1 (no fields)
                'DBname2	ObjName1	DescDBname2ObjName1
                'NULL       ObjName2	DescDBname2ObjName1Field1
                'DBname3	ObjName3	DescObjName3
                'NULL       ObjName3	DescObjName3Field1
                '...
                DBName = If(recordset.IsDBNull(0), "", recordset.GetValue(0))
                Documentation = If(recordset.IsDBNull(2), "", recordset.GetValue(2))
                If recordset.IsDBNull(1) Then
                    UserMsg("Error in filling ConfigDoc: got empty table/view object name, query: " + vbCrLf + ConfigDocQuery)
                    GoTo err
                Else
                    ObjectName = recordset.GetValue(1)
                End If
                ' check for table/view/proc/function object change, also regard DBname change
                If currObjectName <> ObjectName Or (currDBName <> DBName And DBName <> "") Then
                    ' first Object: only fill DBName
                    If currObjectName = "" Then
                        currDBName = DBName
                    Else
                        ' afterwards (object name change)
                        getConfigDocCollection.Add(charBeforeDBnameConfigDoc + currDBName + "\" + currObjectName + ".xcl", ConfigDoc)
                        If DBName = "" Then
                            UserMsg("Error in filling ConfigDoc: got empty database name for new table/view/proc/function object '" + ObjectName + "', query: " + vbCrLf + ConfigDocQuery)
                            GoTo err
                        Else
                            currDBName = DBName
                        End If
                    End If
                    currObjectName = ObjectName
                    ConfigDoc = Documentation
                Else
                    ConfigDoc += Documentation
                End If
            End While
            getConfigDocCollection.Add(charBeforeDBnameConfigDoc + currDBName + "\" + currObjectName + ".xcl", ConfigDoc)
        Catch ex As Exception
            UserMsg("Error in filling ConfigDoc: " + ex.Message + ", query: " + vbCrLf + ConfigDocQuery)
            GoTo err
        End Try
err:
        conn.Close()
    End Function

    ''' <summary>Get the current selected Environment for DB Functions</summary>
    ''' <returns>ConfigName of environment</returns>
    <ExcelFunction(Description:="Get the current selected Environment for DB Functions")>
    Public Function DBAddinEnvironment() As String
        ExcelDnaUtil.Application.Volatile()
        Try
            DBAddinEnvironment = fetchSetting("ConfigName" + env(), "")
            If ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual Then DBAddinEnvironment = "calc Mode is manual, please press F9 to get current DBAddin environment !"
        Catch ex As Exception
            DBAddinEnvironment = "Error happened: " + ex.Message
            LogWarn(ex.Message)
        End Try
    End Function

    ''' <summary>Get the settings as given in keyword (e.g. SERVER=) for the currently selected Environment for DB Functions</summary>
    ''' <returns>Server part from connection string of environment</returns>
    <ExcelFunction(Description:="Get the settings as given in keyword (e.g. SERVER=) for the currently selected Environment for DB Functions")>
    Public Function DBAddinSetting(<ExcelArgument(Description:="keyword for setting to get")> keyword As Object) As String
        ExcelDnaUtil.Application.Volatile()
        Try
            Dim theConnString As String = fetchSetting("ConstConnString" + env(), "")
            If TypeName(keyword) = "ExcelMissing" Or TypeName(keyword) = "ExcelEmpty" Or keyword.ToString() = "" Then
                DBAddinSetting = "No keyword, returning whole connection string of current environment: " + theConnString
            Else
                Dim keywordstart As Integer = InStr(1, UCase(theConnString), UCase(keyword.ToString()))
                If keywordstart > 0 Then
                    keywordstart += Len(keyword.ToString())
                    DBAddinSetting = Mid$(theConnString, keywordstart, IIf(InStr(keywordstart, theConnString, ";") = 0, Len(theConnString) + 1, InStr(keywordstart, theConnString, ";")) - keywordstart)
                Else
                    DBAddinSetting = keyword.ToString() + " was not found in connection string of current environment: " + theConnString
                End If
            End If
            If ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual Then DBAddinSetting = "calc Mode is manual, please press F9 to get current DBAddin server setting !"
        Catch ex As Exception
            DBAddinSetting = "Error happened: " + ex.Message
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
                checkParamsAndCache = "query contains: #" + Replace(Query.ToString(), "ExcelError", "")
                If Query = ExcelError.ExcelErrorValue Then checkParamsAndCache += " (in case query is an argument of a DBfunction, check if it's > 255 chars)"
            ElseIf TypeName(Query) = "Object(,)" Then
                ' if query is reference then get the query string out of it..
                Dim myCell
                Dim retval As String = ""
                For Each myCell In Query
                    If TypeName(myCell) = "ExcelEmpty" Then
                        'do nothing here
                    ElseIf Left(TypeName(myCell), 10) = "ExcelError" Then
                        checkParamsAndCache = "query contains: #" + Replace(myCell.ToString(), "ExcelError", "") + "!"
                        If myCell = ExcelError.ExcelErrorValue Then checkParamsAndCache += " (in case query is an argument of a DBfunction, check if it's > 255 chars)"
                    ElseIf IsNumeric(myCell) Then
                        ' ConnString = "" means a query from DBSetPowerQuery, here preserve the cr-lf !
                        retval += Convert.ToString(myCell, System.Globalization.CultureInfo.InvariantCulture) + IIf(ConnString = "", vbCrLf, " ")
                    Else
                        retval += myCell.ToString() + IIf(ConnString = "", vbCrLf, " ")
                    End If
                    Query = retval
                Next
                If retval.Length = 0 Then checkParamsAndCache = "empty query provided !"
            ElseIf TypeName(Query) = "String" Then
                If Query.ToString().Length = 0 Then checkParamsAndCache = "empty query provided !"
            Else
                checkParamsAndCache = "query parameter invalid (not a range and not a string) !"
            End If
        End If
        If checkParamsAndCache.Length > 0 Then
            ' refresh the query cache ...
            If queryCache.ContainsKey(callID) Then queryCache.Remove(callID)
            Exit Function
        End If

        ' caching check mechanism to avoid unnecessary recalculations/re-fetching
        Dim doFetching As Boolean
        If queryCache.ContainsKey(callID) Then
            doFetching = (ConnString + Query.ToString() <> queryCache(callID))
            ' refresh the query cache with new query/connection string ...
            If doFetching Then
                queryCache.Remove(callID)
                queryCache.Add(callID, ConnString + Query.ToString())
            End If
        Else
            queryCache.Add(callID, ConnString + Query.ToString())
            doFetching = True
        End If
        If doFetching Then
            ' remove Status Container to signal a new calculation request
            If StatusCollection.ContainsKey(callID) Then StatusCollection.Remove(callID)
        Else
            ' return Status Containers Message as last result
            If StatusCollection.ContainsKey(callID) Then
                If Not IsNothing(StatusCollection(callID).statusMsg) Then checkParamsAndCache = If(ConnString = "", "DBSetPowerQuery: ", "(last result:)") + StatusCollection(callID).statusMsg
            End If
        End If
    End Function

    ''' <summary>converts ExcelDna (C API) reference to excel (COM Based) Range</summary>
    ''' <param name="reference">reference to be converted</param>
    ''' <returns>range for passed reference</returns>
    Private Function ToRange(reference As Object) As Excel.Range
        If TypeName(reference) <> "ExcelReference" Then Return Nothing
        Dim item As String
        Try : item = XlCall.Excel(XlCall.xlSheetNm, reference)
        Catch ex As Exception
            Return Nothing
        End Try

        Dim index As Integer
        Dim wbname As String
        Try
            index = item.LastIndexOf("]")
            wbname = item.Substring(0, index).Substring(1)
            item = item.Substring(index + 1)
        Catch ex As Exception
            Return Nothing
        End Try
        Dim returnedRange As Excel.Range = Nothing
        Try
            Dim ws As Excel.Worksheet = ExcelDnaUtil.Application.Workbooks(wbname).Worksheets(item)
            returnedRange = ws.Range(ws.Cells(reference.RowFirst + 1, reference.ColumnFirst + 1), ws.Cells(reference.RowLast + 1, reference.ColumnLast + 1))
        Catch ex As Exception : End Try
        Return returnedRange
    End Function

End Module