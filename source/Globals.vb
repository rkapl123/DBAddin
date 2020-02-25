Imports ExcelDna.Integration
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports System.Collections.Generic
Imports System.Diagnostics


''' <summary>Global variables and functions for DB Addin</summary>
Public Module Globals
    ' general Global objects/variables
    ''' <summary>ribbon menu handler</summary>
    Public theMenuHandler As MenuHandler
    ''' <summary>currently selected environment for DB Functions, zero based (env -1) !!</summary>
    Public selectedEnvironment As Integer
    ''' <summary>reference object for the Addins ribbon</summary>
    Public theRibbon As CustomUI.IRibbonUI
    ''' <summary>environment definitions</summary>
    Public environdefs As String() = {}
    ''' <summary>DBModif definition collections of DBmodif types (key of top level dictionary) with values beinig collections of DBModifierNames (key of contained dictionaries) and DBModifiers (value of contained dictionaries))</summary>
    Public DBModifDefColl As Dictionary(Of String, Dictionary(Of String, DBModif))
    ''' <summary>the selected event level in the About box</summary>
    Public EventLevelSelected As String
    ''' <summary>the log listener</summary>
    Public theLogListener As TraceListener

    ' Global settings
    Public DebugAddin As Boolean
    ''' <summary>Default ConnectionString, if no connection string is given by user....</summary>
    Public ConstConnString As String
    ''' <summary>global connection timeout (can't be set in DB functions)</summary>
    Public CnnTimeout As Integer
    ''' <summary>global command timeout (can't be set in DB functions)</summary>
    Public CmdTimeout As Integer
    ''' <summary>default formatting style used in DBDate</summary>
    Public DefaultDBDateFormatting As Integer
    '''<summary>the folder used to store predefined DB sheet definitions</summary>
    Public DBSheetDefinitionsFolder As String

    ''' <summary>encapsulates setting fetching (currently registry)</summary>
    ''' <param name="Key">registry key to take value from</param>
    ''' <param name="defaultValue">Value that is taken if Key was not found</param>
    ''' <returns>the setting value</returns>
    Public Function fetchSetting(Key As String, defaultValue As String) As String
        fetchSetting = GetSetting("DBAddin", "Settings", Key, defaultValue)
    End Function

    ''' <summary>encapsulates setting storing (currently registry)</summary>
    ''' <param name="Key">registry key to store value to</param>
    ''' <param name="Value">value to store</param>
    Public Sub storeSetting(Key As String, Value As String)
        SaveSetting("DBAddin", "Settings", Key, Value)
    End Sub

    ''' <summary>initializes global configuration variables from registry</summary>
    Public Sub initSettings()
        Try
            DebugAddin = CBool(fetchSetting("DebugAddin", "False"))
            ConstConnString = fetchSetting("ConstConnString", String.Empty)
            CnnTimeout = CInt(fetchSetting("CnnTimeout", "15"))
            CmdTimeout = CInt(fetchSetting("CmdTimeout", "60"))
            ConfigStoreFolder = fetchSetting("ConfigStoreFolder", String.Empty)
            specialConfigStoreFolders = Split(fetchSetting("specialConfigStoreFolders", String.Empty), ":")
            DefaultDBDateFormatting = CInt(fetchSetting("DefaultDBDateFormatting", "0"))
            ' load environments
            Dim i As Integer = 1
            ReDim Preserve environdefs(-1)
            Dim ConfigName As String
            Do
                ConfigName = fetchSetting("ConfigName" + i.ToString(), vbNullString)
                If Len(ConfigName) > 0 Then
                    ReDim Preserve environdefs(environdefs.Length)
                    environdefs(environdefs.Length - 1) = ConfigName + " - " + i.ToString()
                End If
                ' set selectedEnvironment
                If fetchSetting("ConstConnString" + i.ToString(), vbNullString) = ConstConnString Then
                    selectedEnvironment = i - 1
                End If
                i += 1
            Loop Until Len(ConfigName) = 0
        Catch ex As Exception
            ErrorMsg("Error in initialization of Settings (DBAddin.initSettings):" + ex.Message)
        End Try
    End Sub

    ''' <summary>Logs Message of eEventType to System.Diagnostics.Trace</summary>
    ''' <param name="Message">Message to be logged</param>
    ''' <param name="eEventType">event type: info, warning, error</param>
    ''' <param name="caller">reflection based caller information: module.method</param>
    Private Sub WriteToLog(Message As String, eEventType As EventLogEntryType, caller As String)
        Select Case eEventType
            Case EventLogEntryType.Information : Trace.TraceInformation("{0}: {1}", caller, Message)
            Case EventLogEntryType.Warning : Trace.TraceWarning("{0}: {1}", caller, Message)
            Case EventLogEntryType.Error : Trace.TraceError("{0}: {1}", caller, Message)
        End Select
    End Sub

    ''' <summary>Logs error messages</summary>
    ''' <param name="LogMessage">the message to be logged</param>
    Public Sub LogError(LogMessage As String)
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Dim caller As String = theMethod.ReflectedType.FullName & "." & theMethod.Name
        WriteToLog(LogMessage, EventLogEntryType.Error, caller)
    End Sub

    ''' <summary>Logs warning messages</summary>
    ''' <param name="LogMessage">the message to be logged</param>
    Public Sub LogWarn(LogMessage As String)
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Dim caller As String = theMethod.ReflectedType.FullName & "." & theMethod.Name
        WriteToLog(LogMessage, EventLogEntryType.Warning, caller)
    End Sub

    ''' <summary>Logs informational messages</summary>
    ''' <param name="LogMessage">the message to be logged</param>
    Public Sub LogInfo(LogMessage As String)
        If DebugAddin Then
            Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
            Dim caller As String = theMethod.ReflectedType.FullName & "." & theMethod.Name
            WriteToLog(LogMessage, EventLogEntryType.Information, caller)
        End If
    End Sub

    ''' <summary>show Error message to User</summary> 
    ''' <param name="LogMessage">the message to be shown/logged</param>
    ''' <param name="exitMe">can be set to True to let the user avoid repeated error messages, returns true if cancel was clicked</param>
    Public Sub ErrorMsg(LogMessage As String, Optional ByRef exitMe As Boolean = False, Optional exitMsg As String = "(press Cancel to avoid further error messages of this kind)")
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Dim caller As String = theMethod.ReflectedType.FullName & "." & theMethod.Name
        WriteToLog(LogMessage, EventLogEntryType.Warning, caller) ' to avoid popup of trace log
        Dim retval As Integer = MsgBox(LogMessage & IIf(exitMe, vbCrLf & exitMsg, ""), vbExclamation + IIf(exitMe, vbOKCancel, vbOKOnly), "DBAddin Error")
        If retval = vbCancel Then
            exitMe = True
        Else
            exitMe = False
        End If
    End Sub

    ''' <summary>refresh DB Functions (and - if called from outside any db function area - all other external data ranges)</summary>
    <ExcelCommand(Name:="refreshData", ShortCut:="^R")>
    Public Sub refreshData()
        initSettings()

        ' enable events in case there were some problems in procedure with EnableEvents = false
        Try
            ExcelDnaUtil.Application.EnableEvents = True
        Catch ex As Exception
            ErrorMsg("Can't refresh data while lookup dropdown is open !!")
            Exit Sub
        End Try

        ' also reset the database connection in case of errors (might be nothing or not open...)
        Try : conn.Close() : Catch ex As Exception : End Try
        conn = Nothing
        dontTryConnection = False
        Try
            ' reset query cache for current workbook, so we really get new data !
            Dim tempColl1 As Dictionary(Of String, String) = New Dictionary(Of String, String)(queryCache) ' clone dictionary to be able to remove items...
            For Each resetkey As String In tempColl1.Keys
                If InStr(resetkey, "[" & ExcelDnaUtil.Application.ActiveWorkbook.Name & "]") > 0 Then queryCache.Remove(resetkey)
            Next
            Dim tempColl2 As Dictionary(Of String, ContainedStatusMsg) = New Dictionary(Of String, ContainedStatusMsg)(StatusCollection)
            For Each resetkey As String In tempColl2.Keys
                If InStr(resetkey, "[" & ExcelDnaUtil.Application.ActiveWorkbook.Name & "]") > 0 Then StatusCollection.Remove(resetkey)
            Next
            Dim underlyingName As String = getDBunderlyingNameFromRange(ExcelDnaUtil.Application.ActiveCell)

            ' now for DBListfetch/DBRowfetch resetting, first outside of all db function areas...
            If underlyingName = "" Then
                refreshDBFunctions(ExcelDnaUtil.Application.ActiveWorkbook)
                ' general refresh: also refresh all embedded queries and pivot tables..
                Try
                    Dim ws As Excel.Worksheet
                    Dim qrytbl As Excel.QueryTable
                    Dim pivottbl As Excel.PivotTable

                    For Each ws In ExcelDnaUtil.Application.ActiveWorkbook.Worksheets
                        For Each qrytbl In ws.QueryTables
                            qrytbl.Refresh()
                        Next
                        For Each pivottbl In ws.PivotTables
                            pivottbl.PivotCache.Refresh()
                        Next
                    Next
                Catch ex As Exception
                End Try
            Else ' then inside a db function area (target or source = function cell)
                ' we're being called on a target functions area (additionally given in DBListFetch)
                If Left$(underlyingName, 10) = "DBFtargetF" Then
                    underlyingName = Replace(underlyingName, "DBFtargetF", "DBFsource", 1, , vbTextCompare)
                    If Not ExcelDnaUtil.Application.Range(underlyingName).Parent Is ExcelDnaUtil.Application.ActiveSheet Then
                        ExcelDnaUtil.Application.ScreenUpdating = False
                        origWS = ExcelDnaUtil.Application.ActiveSheet
                        Try
                            ExcelDnaUtil.Application.Range(underlyingName).Parent.Parent.Activate
                            ExcelDnaUtil.Application.Range(underlyingName).Parent.Select
                        Catch ex As Exception : End Try
                    End If
                    ExcelDnaUtil.Application.Range(underlyingName).Dirty()
                    ' we're being called on a target area
                ElseIf Left$(underlyingName, 9) = "DBFtarget" Then
                    underlyingName = Replace(underlyingName, "DBFtarget", "DBFsource", 1, , vbTextCompare)
                    ' return to source functions sheet to work around Dirty method problem (cell's sheet needs to be selected for Dirty to work on that cell)
                    If Not ExcelDnaUtil.Application.Range(underlyingName).Parent Is ExcelDnaUtil.Application.ActiveSheet Then
                        ExcelDnaUtil.Application.ScreenUpdating = False
                        origWS = ExcelDnaUtil.Application.ActiveSheet
                        Try
                            ExcelDnaUtil.Application.Range(underlyingName).Parent.Parent.Activate
                            ExcelDnaUtil.Application.Range(underlyingName).Parent.Select
                        Catch ex As Exception : End Try
                    End If
                    ExcelDnaUtil.Application.Range(underlyingName).Dirty()
                    ' we're being called on a source (invoking function) cell
                ElseIf Left$(underlyingName, 9) = "DBFsource" Then
                    Try : ExcelDnaUtil.Application.Range(underlyingName).Dirty() : Catch ex As Exception : End Try
                Else
                    refreshDBFunctions(ExcelDnaUtil.Application.ActiveWorkbook)
                End If
            End If
        Catch ex As Exception
            LogError("Error (" & ex.Message & ") in refreshData !")
        End Try
    End Sub

    ''' <summary>jumps between DB Function and target area</summary>
    <ExcelCommand(Name:="jumpButton", ShortCut:="^J")>
    Public Sub jumpButton()
        If checkMultipleDBRangeNames(ExcelDnaUtil.Application.ActiveCell) Then
            ErrorMsg("Multiple hidden DB Function names in selected cell (making 'jump' ambigous/impossible), please use purge names tool!")
            Exit Sub
        End If

        Dim underlyingName As String = getDBunderlyingNameFromRange(ExcelDnaUtil.Application.ActiveCell)
        If underlyingName = "" Then Exit Sub
        If Left$(underlyingName, 10) = "DBFtargetF" Then
            underlyingName = Replace(underlyingName, "DBFtargetF", "DBFsource", 1, , vbTextCompare)
        ElseIf Left$(underlyingName, 9) = "DBFtarget" Then
            underlyingName = Replace(underlyingName, "DBFtarget", "DBFsource", 1, , vbTextCompare)
        Else
            underlyingName = Replace(underlyingName, "DBFsource", "DBFtarget", 1, , vbTextCompare)
        End If
        Try
            ExcelDnaUtil.Application.Range(underlyingName).Parent.Select()
            ExcelDnaUtil.Application.Range(underlyingName).Select()
        Catch ex As Exception
            ErrorMsg("Can't jump to target/source, corresponding workbook open? " & ex.Message)
        End Try
    End Sub

    ''' <summary>marks a row in a DBMapper for deletion</summary>
    <ExcelCommand(Name:="deleteRow", ShortCut:="^D")>
    Public Sub deleteRow()
        If Globals.DBModifDefColl.ContainsKey("DBMapper") Then
            Dim targetName As String = getDBModifNameFromRange(ExcelDnaUtil.Application.Selection)
            If Left(targetName, 8) = "DBMapper" Then DirectCast(Globals.DBModifDefColl("DBMapper").Item(targetName), DBMapper).doCUDMarks(ExcelDnaUtil.Application.Selection, True)
        End If
    End Sub

    ''' <summary>splits theString into tokens delimited by delimiter, ignoring delimiters inside quotes and brackets</summary>
    ''' <param name="theString">string to be split into tokens</param>
    ''' <param name="delimiter">delimiter that string is to be split by</param>
    ''' <param name="quote">quote character where delimiters should be ignored inside</param>
    ''' <param name="startStr">part of theString where splitting should start after</param>
    ''' <param name="openBracket">opening bracket character</param>
    ''' <param name="closeBracket">closing bracket character</param>
    ''' <returns>the list of tokens</returns>
    ''' <remarks>theString is split starting from startStr up to the first balancing closing Bracket (as defined by openBracket and closeBracket)
    ''' startStr, openBracket and closeBracket are case insensitive for comparing with theString.
    ''' the tokens are not blank trimmed !!</remarks>
    Public Function functionSplit(ByVal theString As String, delimiter As String, quote As String, startStr As String, openBracket As String, closeBracket As String) As Object
        Dim tempString As String
        Dim finalResult

        Try
            ' find startStr
            tempString = Mid$(theString, InStr(1, UCase$(theString), UCase$(startStr)) + Len(startStr))
            ' rip out the balancing string now...
            tempString = balancedString(tempString, openBracket, closeBracket, quote)
            If tempString.Length = 0 Then
                ErrorMsg("couldn't produce balanced string from " & theString)
                functionSplit = Nothing
                Exit Function
            End If
            tempString = replaceDelimsWithSpecialSep(tempString, delimiter, quote, openBracket, closeBracket, vbTab)
            finalResult = Split(tempString, vbTab)
            functionSplit = finalResult
        Catch ex As Exception
            LogError("Error: " & ex.Message)
            functionSplit = Nothing
        End Try
    End Function

    ''' <summary>returns the minimal bracket balancing string contained in theString, opening bracket defined in openBracket, closing bracket defined in closeBracket
    ''' disregarding quoted areas inside optionally given quote charachter/string</summary>
    ''' <param name="theString"></param>
    ''' <param name="openBracket"></param>
    ''' <param name="closeBracket"></param>
    ''' <param name="quote"></param>
    ''' <returns>the balanced string</returns>
    Public Function balancedString(theString As String, openBracket As String, closeBracket As String, Optional quote As String = "") As String
        Dim startBalance As Long, endBalance As Long, i As Long, countOpen As Long, countClose As Long
        balancedString = String.Empty
        Dim quoteMode As Boolean = False
        Try
            startBalance = 0
            For i = 1 To Len(theString)
                If Left$(Mid$(theString, i), Len(quote)) = quote And quote.Length > 0 And Not quoteMode Then
                    quoteMode = True
                Else
                    If Not quoteMode Then
                        If Left$(Mid$(theString, i), Len(openBracket)) = openBracket Then
                            If startBalance = 0 Then startBalance = i
                            countOpen += 1
                        End If
                        If startBalance <> 0 And Left$(Mid$(theString, i), Len(closeBracket)) = closeBracket Then countClose += 1
                    Else
                        If Left$(Mid$(theString, i), Len(quote)) = quote And quote.Length > 0 Then quoteMode = False
                    End If
                End If

                If countOpen = countClose And startBalance <> 0 Then
                    endBalance = i - 1
                    Exit For
                End If
            Next
            If endBalance <> 0 Then
                balancedString = Mid$(theString, startBalance + 1, endBalance - startBalance)
            End If
        Catch ex As Exception
            LogError("Error: " & ex.Message)
        End Try
    End Function

    ''' <summary>replaces the delimiter (delimiter) inside theString with specialSep, regarding both quoted areas inside quote and bracketed areas (inside openBracket/closeBracket)</summary>
    ''' <param name="theString"></param>
    ''' <param name="delimiter"></param>
    ''' <param name="quote"></param>
    ''' <param name="openBracket"></param>
    ''' <param name="closeBracket"></param>
    ''' <param name="specialSep"></param>
    ''' <returns>replaced string</returns>
    Public Function replaceDelimsWithSpecialSep(theString As String, delimiter As String, quote As String, openBracket As String, closeBracket As String, specialSep As String) As String
        Dim openedBrackets As Long, quoteMode As Boolean
        Dim i As Long
        replaceDelimsWithSpecialSep = String.Empty
        Try
            For i = 1 To Len(theString)
                If Left$(Mid$(theString, i), Len(quote)) = quote And quote.Length > 0 And Not quoteMode Then
                    quoteMode = True
                Else
                    If quoteMode And Left$(Mid$(theString, i), Len(quote)) = quote And quote.Length > 0 Then quoteMode = False
                End If

                If Left$(Mid$(theString, i), Len(openBracket)) = openBracket And openBracket.Length > 0 And Not quoteMode Then
                    openedBrackets += 1
                End If
                If Left$(Mid$(theString, i), Len(closeBracket)) = closeBracket And closeBracket.Length > 0 And Not quoteMode Then
                    openedBrackets -= 1
                End If

                If Not (openedBrackets > 0 Or quoteMode) Then
                    If Left$(Mid$(theString, i), Len(delimiter)) = delimiter Then
                        replaceDelimsWithSpecialSep &= specialSep
                    Else
                        replaceDelimsWithSpecialSep &= Mid$(theString, i, 1)
                    End If
                Else
                    replaceDelimsWithSpecialSep &= Mid$(theString, i, 1)
                End If
            Next
        Catch ex As Exception
            LogError("Error: " & ex.Message)
        End Try
    End Function

    ''' <summary>changes theString by replacing substring starting after keystr and ending with separator with changed, case insensitive !!</summary>
    ''' <param name="theString"></param>
    ''' <param name="keystr"></param>
    ''' <param name="changed"></param>
    ''' <param name="separator"></param>
    ''' <returns>the changed string</returns>
    Public Function Change(ByVal theString As String, ByVal keystr As String, ByVal changed As String, ByVal separator As String) As String
        Dim replaceBeg, replaceEnd As Integer

        replaceBeg = InStr(1, UCase$(theString), UCase$(keystr))
        If replaceBeg = 0 Then
            Change = String.Empty
            Exit Function
        End If
        replaceEnd = InStr(replaceBeg, UCase$(theString), UCase$(separator))
        If replaceEnd = 0 Then replaceEnd = Len(theString) + 1
        Change = Left$(theString, replaceBeg - 1 + Len(keystr)) & changed & Right$(theString, Len(theString) - replaceEnd + 1)
    End Function

    ''' <summary>fetches substring starting after keystr and ending with separator from theString, case insensitive !! if separator is "" then fetch to end of string</summary>
    ''' <param name="theString"></param>
    ''' <param name="keystr"></param>
    ''' <param name="separator"></param>
    ''' <returns>the fetched substring</returns>
    Public Function fetch(ByVal theString As String, ByVal keystr As String, ByVal separator As String) As String
        Dim fetchBeg As Integer, fetchEnd As Integer

        fetchBeg = InStr(1, UCase$(theString), UCase$(keystr))
        If fetchBeg = 0 Then Return String.Empty
        fetchEnd = InStr(fetchBeg + Len(keystr), UCase$(theString), UCase$(separator))
        If fetchEnd = 0 Or separator.Length = 0 Then fetchEnd = Len(theString) + 1
        fetch = Mid$(theString, fetchBeg + Len(keystr), fetchEnd - (fetchBeg + Len(keystr)))
    End Function

    ''' <summary>checks whether worksheet called theName exists</summary>
    ''' <param name="theName"></param>
    ''' <returns>True if sheet exists</returns>
    Public Function existsSheet(ByRef theName As String) As Boolean
        existsSheet = True
        Try
            Dim dummy As String = ExcelDnaUtil.Application.Worksheets(theName).name
        Catch ex As Exception
            existsSheet = False
        End Try
    End Function

    ''' <summary>gets underlying DBtarget/DBsource Name from theRange</summary>
    ''' <param name="theRange"></param>
    ''' <returns>the retrieved name</returns>
    Public Function getDBunderlyingNameFromRange(theRange As Excel.Range) As String
        Dim nm As Excel.Name
        Dim rng, testRng As Excel.Range

        getDBunderlyingNameFromRange = ""
        Try
            For Each nm In theRange.Parent.Parent.Names
                rng = Nothing
                Try : rng = nm.RefersToRange : Catch ex As Exception : End Try
                If Not rng Is Nothing Then
                    testRng = Nothing
                    Try : testRng = ExcelDnaUtil.Application.Intersect(theRange, rng) : Catch ex As Exception : End Try
                    If Not IsNothing(testRng) And (InStr(1, nm.Name, "DBFtarget") >= 1 Or InStr(1, nm.Name, "DBFsource") >= 1) Then
                        getDBunderlyingNameFromRange = nm.Name
                        Exit Function
                    End If
                End If
            Next
        Catch ex As Exception
            LogError("Error: " & ex.Message)
        End Try
    End Function

    ''' <summary>check if multiple (hidden, containing DBtarget or DBsource) DB Function names exist in theRange</summary>
    ''' <param name="theRange"></param>
    ''' <returns>True if multiple names exist</returns>
    Public Function checkMultipleDBRangeNames(theRange As Excel.Range) As Boolean
        Dim nm As Excel.Name
        Dim rng, testRng As Excel.Range
        Dim foundNames As Integer = 0

        checkMultipleDBRangeNames = False
        Try
            For Each nm In theRange.Parent.Parent.Names
                rng = Nothing
                Try : rng = nm.RefersToRange : Catch ex As Exception : End Try
                If Not rng Is Nothing And Not (nm.Name Like "*ExterneDaten*" Or nm.Name Like "*_FilterDatabase") Then
                    testRng = Nothing
                    Try : testRng = ExcelDnaUtil.Application.Intersect(theRange, rng) : Catch ex As Exception : End Try
                    If Not IsNothing(testRng) And (InStr(1, nm.Name, "DBFtarget") >= 1 Or InStr(1, nm.Name, "DBFsource") >= 1) Then
                        foundNames += 1
                    End If
                End If
            Next
            If foundNames > 1 Then checkMultipleDBRangeNames = True
        Catch ex As Exception
            LogError("Error: " & ex.Message)
        End Try
    End Function

    ''' <summary>find out whether to recalc full (and do it), if we have DBFuncs in the workbook somewhere</summary>
    ''' <param name="Wb">workbook to refresh DB Functions in</param>
    ''' <param name="ignoreCalcMode">when calling refreshDBFunctions time delayed (when saving a workbook and DBFC* is set), need to trigger calculation regardless of calculation mode being manual, otherwise data is not refreshed</param>
    Public Sub refreshDBFunctions(Wb As Excel.Workbook, Optional ignoreCalcMode As Boolean = False)
        Dim searchCells As Excel.Range
        Dim ws As Excel.Worksheet
        Dim needRecalc As Boolean

        ' hidden workbooks produce an error when searching for cells, this is captured by 
        If TypeName(ExcelDnaUtil.Application.Calculation) = "Error" Then
            MsgBox("ExcelDnaUtil.Application.Calculation = Error, " & Wb.Path & "\" & Wb.Name & " (hidden workbooks produce calculation errors...)")
            Exit Sub
        End If
        Try
            needRecalc = False
            Dim cellcount As Long = 0
            For Each ws In Wb.Worksheets
                cellcount += ExcelDnaUtil.Application.WorksheetFunction.CountIf(ws.Range("1:" & ws.Rows.Count), "<>")
            Next
            If cellcount > CLng(fetchSetting("maxCellCount", "100000")) Then
                Dim retval As MsgBoxResult = MsgBox("This large workbook (" & cellcount.ToString() & " filled cells >" & CLng(fetchSetting("maxCellCount", "100000")) & ") might take long to search for DB functions to refresh, continue ?" & vbCrLf & "Click Cancel to add DBFskip to this Workbook, avoiding this search in the future...", vbQuestion + vbYesNoCancel, "Refresh DB functions")
                If retval <> vbYes Then
                    If retval = vbCancel Then
                        Try
                            Wb.CustomDocumentProperties.Add(Name:="DBFskip", LinkToContent:=False, Type:=MsoDocProperties.msoPropertyTypeBoolean, Value:=True)
                        Catch ex As Exception
                            LogWarn("Error when adding DBFskip to Workbook:" + ex.Message)
                        End Try
                    End If
                    Exit Sub
                End If
            End If
            For Each ws In Wb.Worksheets
                Dim theFunc As String
                For Each theFunc In {"DBListFetch(", "DBRowFetch(", "DBSetQuery("}
                    searchCells = ws.Cells.Find(What:=theFunc, After:=ws.Range("A1"), LookIn:=Excel.XlFindLookIn.xlFormulas, LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlNext, MatchCase:=False)
                    If Not (searchCells Is Nothing) Then
                        ' reset the cell find dialog....
                        searchCells = Nothing
                        searchCells = ws.Cells.Find(What:="", After:=ws.Range("A1"), LookIn:=Excel.XlFindLookIn.xlFormulas, LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlNext, MatchCase:=False)
                        needRecalc = True
                        GoTo done
                    End If
                Next
                ' reset the cell find dialog....
                searchCells = Nothing
                searchCells = ws.Cells.Find(What:="", After:=ws.Range("A1"), LookIn:=Excel.XlFindLookIn.xlFormulas, LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlNext, MatchCase:=False)
            Next
done:
            If needRecalc And (ExcelDnaUtil.Application.Calculation <> Excel.XlCalculation.xlCalculationManual Or ignoreCalcMode) Then
                queryCache.Clear()
                LogInfo("ExcelDnaUtil.Application.CalculateFull called " & Wb.Path & "\" & Wb.Name)
                ExcelDnaUtil.Application.CalculateFull()
            Else
                LogInfo("no dbfunc found... " & Wb.Path & "\" & Wb.Name)
            End If
        Catch ex As Exception
            LogError("Error: " & ex.Message & ", " & Wb.Path & "\" & Wb.Name)
        End Try
    End Sub

    ''' <summary>"repairs" legacy functions from old VB6-COM Addin by removing "DBAddin.Functions." before function name</summary>
    ''' <param name="showReponse">in case this is called interactively, provide a response in case of no legacy functions there</param>
    Public Sub repairLegacyFunctions(Optional showReponse As Boolean = False)
        Dim searchCell As Excel.Range
        Dim foundLegacyWS As Collection = New Collection
        Dim xlcalcmode As Long = ExcelDnaUtil.Application.Calculation
        Dim actWB As Excel.Workbook = ExcelDnaUtil.Application.ActiveWorkbook
        If IsNothing(actWB) Then
            LogWarn("no active workbook available !")
            Exit Sub
        End If
        Dim skipLegacyCheck As Boolean
        Try
            skipLegacyCheck = actWB.CustomDocumentProperties("DBFNoLegacyCheck").Value
        Catch ex As Exception
            skipLegacyCheck = False
        End Try
        If skipLegacyCheck Then Exit Sub
        Try
            ' count nonempty cells in workbook...
            Dim cellcount As Long = 0
            For Each ws In actWB.Worksheets
                cellcount += ExcelDnaUtil.Application.WorksheetFunction.CountIf(ws.Range("1:" & ws.Rows.Count), "<>")
            Next
            ' warn if above threshold
            If cellcount > CLng(fetchSetting("maxCellCount", "100000")) Then
                Dim retval As MsgBoxResult = MsgBox("This large workbook (" & cellcount.ToString() & " filled cells >" & CLng(fetchSetting("maxCellCount", "100000")) & ") might take long to search for legacy functions, continue ?" & vbCrLf & "Cancel to disable legacy function checking for this workbook ...", vbQuestion + MsgBoxStyle.YesNoCancel, "Legacy DBAddin functions")
                If retval <> vbYes Then
                    If retval = vbCancel Then
                        Try
                            actWB.CustomDocumentProperties.Add(Name:="DBFNoLegacyCheck", LinkToContent:=False, Type:=MsoDocProperties.msoPropertyTypeBoolean, Value:=True)
                        Catch ex As Exception
                            LogWarn("Error when adding NoLegacyCheck in workbook:" + ex.Message)
                        End Try
                    End If
                    Exit Sub
                End If
            End If
            For Each ws In actWB.Worksheets
                ' check whether legacy functions exist somewhere ...
                ExcelDnaUtil.Application.StatusBar = "checking for legacy DB functions in active workbook (ESC to stop)"
                searchCell = ws.Cells.Find(What:="DBAddin.Functions.", After:=ws.Range("A1"), LookIn:=Excel.XlFindLookIn.xlFormulas, LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlNext, MatchCase:=False)
                If Not (searchCell Is Nothing) Then foundLegacyWS.Add(ws)
            Next
            If foundLegacyWS.Count > 0 Then
                Dim retval As MsgBoxResult = MsgBox("Found legacy DBAddin functions in active workbook, should they be replaced with current addin functions (save workbook afterwards to persist) ?", vbQuestion + vbYesNo, "Legacy DBAddin functions")
                If retval = vbYes Then
                    ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual ' avoid recalculations during replace action
                    ExcelDnaUtil.Application.DisplayAlerts = False ' avoid warnings for sheet where "DBAddin.Functions." is not found
                    ' remove "DBAddin.Functions." in each sheet...
                    For Each ws In foundLegacyWS
                        ExcelDnaUtil.Application.StatusBar = "Replacing legacy DB functions in active workbook (ESC to stop)"
                        ws.Cells.Replace(What:="DBAddin.Functions.", Replacement:="", LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)
                    Next
                End If
            ElseIf showReponse Then
                MsgBox("No legacy DBAddin functions found in active workbook.", vbExclamation + vbOKOnly, "Legacy DBAddin functions")
            End If
            ' reset the cell find dialog....
            ExcelDnaUtil.Application.ActiveSheet.Cells.Find(What:="", After:=ExcelDnaUtil.Application.ActiveSheet.Range("A1"), LookIn:=Excel.XlFindLookIn.xlFormulas, LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlNext, MatchCase:=False)
        Catch ex As Exception
            LogError("Error occured in replacing legacy DB functions: " & ex.Message)
        End Try
        ExcelDnaUtil.Application.DisplayAlerts = True
        ' only set this back if it was changed to manual as otherwise it would change the (else unchanged) workbook, forcing a confirmation for saving...
        If ExcelDnaUtil.Application.Calculation <> xlcalcmode Then ExcelDnaUtil.Application.Calculation = xlcalcmode
        ExcelDnaUtil.Application.StatusBar = False
    End Sub

    ''' <summary>maintenance procedure to purge names used for dbfunctions from workbook</summary>
    Public Sub purgeNames()
        Dim resultingPurges As String = String.Empty
        Dim retval As MsgBoxResult = MsgBox("Should ExternalData names (from Queries) and names referring to missing references (thus containing #REF!) also be purged?", vbYesNoCancel + vbQuestion, "purge Names")
        If retval = vbCancel Then Exit Sub
        Dim calcMode = ExcelDnaUtil.Application.Calculation
        ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual
        Try
            Dim DBname As Excel.Name
            For Each DBname In ExcelDnaUtil.Application.ActiveWorkbook.Names
                If DBname.Name Like "*ExterneDaten*" Or DBname.Name Like "*ExternalData*" And retval = vbYes Then
                    'resultingPurges += DBname.Name + ", "
                    'DBname.Delete()
                ElseIf DBname.Name Like "DBFtarget*" Then
                    resultingPurges += DBname.Name + ", "
                    DBname.Delete()
                ElseIf DBname.Name Like "DBFsource*" Then
                    resultingPurges += DBname.Name + ", "
                    DBname.Delete()
                ElseIf InStr(1, DBname.RefersTo, "#REF!") > 0 And retval = vbYes Then
                    'resultingPurges += DBname.Name + "(refers to " + DBname.RefersTo + "), "
                    'DBname.Delete()
                End If
            Next
            If resultingPurges = String.Empty Then
                MsgBox("nothing purged...", vbOKOnly, "purge Names")
            Else
                MsgBox("removed " + resultingPurges, vbOKOnly, "purge Names")
                LogInfo("purgeNames removed " + resultingPurges)
            End If
        Catch ex As Exception
            LogError("Error: " & ex.Message)
        End Try
        ExcelDnaUtil.Application.Calculation = calcMode
    End Sub

End Module
