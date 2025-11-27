Imports ExcelDna.Integration
Imports Microsoft.Office.Interop

''' <summary>Global functions for DB Addin</summary>
Public Module Globals

    ''' <summary>splits theString into tokens delimited by delimiter, ignoring delimiters inside quotes and brackets</summary>
    ''' <param name="theString">string to be split into tokens, case insensitive !</param>
    ''' <param name="delimiter">delimiter that string is to be split by</param>
    ''' <param name="quote">quote character where delimiters should be ignored inside</param>
    ''' <param name="startStr">part of theString where splitting should start after, case insensitive !</param>
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
                UserMsg("couldn't produce balanced string from " + theString)
                functionSplit = Nothing
                Exit Function
            End If
            tempString = replaceDelimsWithSpecialSep(tempString, delimiter, quote, openBracket, closeBracket, vbTab)
            finalResult = Split(tempString, vbTab)
            functionSplit = finalResult
        Catch ex As Exception
            UserMsg("Exception: " + ex.Message, "function split into tokens")
            functionSplit = Nothing
        End Try
    End Function

    ''' <summary>returns the minimal bracket balancing string contained in theString, opening bracket defined in openBracket, closing bracket defined in closeBracket
    ''' disregarding quoted areas inside optionally given quote character/string</summary>
    ''' <param name="theString"></param>
    ''' <param name="openBracket"></param>
    ''' <param name="closeBracket"></param>
    ''' <param name="quote"></param>
    ''' <returns>the balanced string</returns>
    Public Function balancedString(theString As String, openBracket As String, closeBracket As String, Optional quote As String = "") As String
        Dim startBalance As Long, endBalance As Long, i As Long, countOpen As Long, countClose As Long
        balancedString = ""
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
            UserMsg("Exception: " + ex.Message, "get minimal balanced string")
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
        replaceDelimsWithSpecialSep = ""
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
                        replaceDelimsWithSpecialSep += specialSep
                    Else
                        replaceDelimsWithSpecialSep += Mid$(theString, i, 1)
                    End If
                Else
                    replaceDelimsWithSpecialSep += Mid$(theString, i, 1)
                End If
            Next
        Catch ex As Exception
            UserMsg("Exception: " + ex.Message, "replace delimiters with special separator")
        End Try
    End Function

    ''' <summary>changes theString to changedString by replacing substring starting AFTER keystr and ending with separator (so "(keystr)...;" will become "(keystr)(changedString);", case insensitive !!</summary>
    ''' <param name="theString"></param>
    ''' <param name="keystr"></param>
    ''' <param name="changedString"></param>
    ''' <param name="separator"></param>
    ''' <returns>the changed string</returns>
    Public Function Change(ByVal theString As String, ByVal keystr As String, ByVal changedString As String, ByVal separator As String) As String
        Dim replaceBeg, replaceEnd As Integer

        replaceBeg = InStr(1, UCase$(theString), UCase$(keystr))
        If replaceBeg = 0 Then
            Change = theString
            Exit Function
        End If
        replaceEnd = InStr(replaceBeg, UCase$(theString), UCase$(separator))
        If replaceEnd = 0 Then replaceEnd = Len(theString) + 1
        Change = Left$(theString, replaceBeg - 1 + Len(keystr)) + changedString + Right$(theString, Len(theString) - replaceEnd + 1)
    End Function

    ''' <summary>fetches substring starting after keystr and ending with separator from theString, case insensitive !! if separator is "" then fetch to end of string</summary>
    ''' <param name="theString">string to be searched</param>
    ''' <param name="keystr">string indicating the start of the substring combination</param>
    ''' <param name="separator">string ending the whole substring, not included in returned string!</param>
    ''' <param name="includeKeyStr">if includeKeyStr is set to true, include keystr in returned string</param>
    ''' <returns>the fetched substring</returns>
    Public Function fetchSubstr(ByVal theString As String, ByVal keystr As String, ByVal separator As String, Optional includeKeyStr As Boolean = False) As String
        Dim fetchBeg As Integer, fetchEnd As Integer

        fetchBeg = InStr(1, UCase$(theString), UCase$(keystr))
        If fetchBeg = 0 Then Return ""
        fetchEnd = InStr(fetchBeg + Len(keystr), UCase$(theString), UCase$(separator))
        If fetchEnd = 0 Or separator.Length = 0 Then fetchEnd = Len(theString) + 1
        fetchSubstr = Mid$(theString, fetchBeg + IIf(includeKeyStr, 0, Len(keystr)), fetchEnd - (fetchBeg + IIf(includeKeyStr, 0, Len(keystr))))
    End Function

    ''' <summary>checks whether worksheet called theName exists in workbook theWb</summary>
    ''' <param name="theName">string name of worksheet name</param>
    ''' <param name="theWb">given workbook</param>
    ''' <returns>True if sheet exists</returns>
    Public Function existsSheet(ByRef theName As String, theWb As Excel.Workbook) As Boolean
        existsSheet = True
        Try
            Dim dummy As String = theWb.Worksheets(theName).name
        Catch ex As Exception
            existsSheet = False
        End Try
    End Function

    ''' <summary>helper function for check whether name exists in active workbook</summary>
    ''' <param name="CheckForName">name to be checked</param>
    ''' <returns>true if name exists</returns>
    Public Function existsName(CheckForName As String) As Boolean
        existsName = False
        On Error GoTo Last
        If Len(ExcelDnaUtil.Application.ActiveWorkbook.Names(CheckForName).Name) <> 0 Then existsName = True
Last:
    End Function

    ''' <summary>checks whether theName exists as a name in Workbook theWb</summary>
    ''' <param name="theName">string name of range name</param>
    ''' <param name="theWb">given workbook</param>
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
    ''' <param name="theName">string name of range name</param>
    ''' <param name="theWs">given sheet</param>
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

    ''' <summary>gets underlying DBtarget/DBsource Name from theRange</summary>
    ''' <param name="theRange">given range</param>
    ''' <returns>the retrieved name</returns>
    Public Function getUnderlyingDBNameFromRange(theRange As Excel.Range) As String
        Dim nm As Excel.Name
        Dim rng, testRng As Excel.Range

        getUnderlyingDBNameFromRange = ""
        Try
            For Each nm In theRange.Parent.Parent.Names
                rng = Nothing
                Try : rng = nm.RefersToRange : Catch ex As Exception : End Try
                If rng IsNot Nothing Then
                    testRng = Nothing
                    Try : testRng = ExcelDnaUtil.Application.Intersect(theRange, rng) : Catch ex As Exception : End Try
                    If testRng IsNot Nothing And (InStr(nm.Name, "DBFtarget") > 0 Or InStr(nm.Name, "DBFsource") > 0) Then
                        Dim WbkSepPos As Integer = InStr(nm.Name, "!")
                        If WbkSepPos > 1 Then
                            getUnderlyingDBNameFromRange = Mid(nm.Name, WbkSepPos + 1)
                        Else
                            getUnderlyingDBNameFromRange = nm.Name
                        End If
                        Exit Function
                    End If
                End If
            Next
        Catch ex As Exception
            UserMsg("Exception: " + ex.Message, "get underlying DBFName from Range")
        End Try
    End Function

    ''' <summary>create a final connection string from passed String or number (environment), as well as a EnvPrefix for showing the environment (or set ConnString)</summary>
    ''' <param name="ConnString">passed connection string or environment number, resolved (=returned) to actual connection string</param>
    ''' <param name="EnvPrefix">prefix for showing environment (ConnString set if no environment)</param>
    Public Sub resolveConnstring(ByRef ConnString As Object, ByRef EnvPrefix As String, getConnStrForDBSet As Boolean)
        If Left(TypeName(ConnString), 10) = "ExcelError" Then Exit Sub
        If TypeName(ConnString) = "ExcelReference" Then ConnString = ConnString.Value
        If TypeName(ConnString) = "ExcelMissing" Then ConnString = ""
        If TypeName(ConnString) = "ExcelEmpty" Then ConnString = ""
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
            If ConnString.ToString() = "" Then ' no ConnString or environment number set: get connection string of currently selected environment
                EnvPrefix = "Env:" + fetchSetting("ConfigName" + env(), "")
                ConnString = fetchSetting("ConstConnString" + env(), "")
                If getConnStrForDBSet Then
                    ' if an alternate connection string is given, use this one...
                    Dim altConnString = fetchSetting("AltConnString" + env(), "")
                    If altConnString <> "" Then
                        ConnString = altConnString
                    Else
                        ' To get the connection string work also for SQLOLEDB provider for SQL Server, change to ODBC driver setting (this can be generally used to fix connection string problems with ListObjects)
                        ConnString = Replace(ConnString, fetchSetting("ConnStringSearch" + env(), "provider=SQLOLEDB"), fetchSetting("ConnStringReplace" + env(), "driver=SQL SERVER"))
                    End If
                End If
            Else
                EnvPrefix = "ConnString set"
            End If
        End If
    End Sub

    ''' <summary>recalculate fully the DB functions, if we have DBFuncs in the workbook somewhere</summary>
    ''' <param name="Wb">workbook to refresh DB Functions in</param>
    ''' <param name="ignoreCalcMode">when calling refreshDBFunctions time delayed (when saving a workbook and DBFC* is set), need to trigger calculation regardless of calculation mode being manual, otherwise data is not refreshed</param>
    Public Sub refreshDBFunctions(Wb As Excel.Workbook, Optional ignoreCalcMode As Boolean = False, Optional calledOnWBOpen As Boolean = False)
        Dim WbNames As Excel.Names
        Try : WbNames = Wb.Names
        Catch ex As Exception
            LogWarn("Exception when trying to get Workbook names: " + ex.Message + ", this might be due to errors in the VBA Macros (missing references)")
            Exit Sub
        End Try

        ' hidden workbooks produce an error when searching for cells, this is captured by 
        If TypeName(ExcelDnaUtil.Application.Calculation) = "Error" Then
            LogWarn("ExcelDnaUtil.Application.Calculation = Error, " + Wb.Path + "\" + Wb.Name + " (hidden workbooks produce calculation errors...)")
            Exit Sub
        End If
        DBModifHelper.preventChangeWhileFetching = True
        Dim calcMode As Long = ExcelDnaUtil.Application.Calculation
        Dim calcModeSet As Boolean = False
        Try
            ' walk through all DB functions (having hidden names DBFsource*) cells there to find DB Functions and change their formula, adding " " to trigger recalculation
            For Each DBname As Excel.Name In WbNames
                Dim DBFuncCell As Excel.Range = Nothing
                If DBname.Name Like "*DBFsource*" Then
                    ' some names might have lost their reference to the cell, so catch this here...
                    Try : DBFuncCell = DBname.RefersToRange : Catch ex As Exception : End Try
                End If
                If Not IsNothing(DBFuncCell) Then
                    If DBFuncCell.Parent.ProtectContents Then
                        UserMsg("Worksheet " + DBFuncCell.Parent.Name + " is content protected, can't refresh DB Functions !")
                        Continue For
                    End If
                    Dim callID As String = "" : Dim underlyingName As String = ""
                    If Not (DBFuncCell.Formula.ToString().ToUpper.Contains("DBLISTFETCH") Or DBFuncCell.Formula.ToString().ToUpper.Contains("DBROWFETCH") Or DBFuncCell.Formula.ToString().ToUpper.Contains("DBSETQUERY")) Then
                        LogWarn("Found former DB Function in Cell " + DBFuncCell.Parent.Name + "!" + DBFuncCell.Address + " that doesn't contain a DB Function anymore.")
                    Else
                        ' only if there are really DB Functions, set calculation to manual to prevent recalculation during "dirtying" the db functions
                        Try
                            ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual
                        Catch ex As Exception
                            UserMsg("Error when trying to change calculation to manual: " + ex.Message + ". This can occur " + vbCrLf + "1) when currently editing a cell (F2) or " + vbCrLf + "2) when workbooks are opened as 'hidden' (View/Window/Hide-Unhide) or" + vbCrLf + "3) when workbooks are opened using MS-office hyper-links (word, outlook, powerpoint)", "refresh DBFunctions")
                            Exit Sub
                        End Try
                        calcModeSet = True
                    End If
                    Try
                        ' repair DBSheet auto-filling lookup functionality, in case it was lost due to accidental editing of these cells.
                        underlyingName = Replace(DBname.Name, "DBFsource", "DBFtarget", 1, , vbTextCompare)
                        Dim DBTargetListObject As Excel.ListObject = Nothing
                        Dim DBSheetName As String = ""
                        Try : DBTargetListObject = ExcelDnaUtil.Application.Range(underlyingName).ListObject : Catch ex As Exception : End Try
                        Try : DBSheetName = Left(getDBModifNameFromRange(DBTargetListObject.Range), 8) : Catch ex As Exception : End Try
                        If Not IsNothing(DBTargetListObject) AndAlso DBSheetName = "DBMapper" Then
                            ' walk through all columns
                            For Each listcol As Excel.ListColumn In DBTargetListObject.ListColumns
                                Dim colFormula As String = ""
                                ' check for formula and store it
                                Try : colFormula = listcol.DataBodyRange.Cells(1, 1).Formula : Catch ex As Exception : End Try
                                If Left(colFormula, 1) = "=" Then
                                    DBModifHelper.preventChangeWhileFetching = True
                                    ' delete whole column
                                    listcol.DataBodyRange.Clear()
                                    ' re-insert the formula, this repairs the auto-filling functionality
                                    listcol.DataBodyRange.Formula = colFormula
                                    DBTargetListObject.QueryTable.PreserveColumnInfo = True ' if there is a lookup formula, always set this as it is required to fill it in automatically
                                    DBModifHelper.preventChangeWhileFetching = False
                                End If
                            Next
                        End If
                        ' get the callID of the underlying name of the target (key of the queryCache and StatusCollection)
                        callID = "[" + DBFuncCell.Parent.Parent.Name + "]" + DBFuncCell.Parent.Name + "!" + DBFuncCell.Address
                        ' remove query cache to force re-fetching
                        If queryCache.ContainsKey(callID) And Not calledOnWBOpen Then queryCache.Remove(callID)
                        ' trigger recalculation by changing formula of DB Function
                        If DBFuncCell.HasFormula Then
                            DBFuncCell.Formula += " "
                        Else
                            LogWarn("DB Function with callID (" + callID + ") of DB Function in Cell (" + DBFuncCell.Parent.Name + "!" + DBFuncCell.Address + ") is not a formula anymore, maybe it was commented out?")
                        End If
                    Catch ex As Exception
                        LogWarn("Exception when setting Formula or getting callID (" + callID + ") of DB Function in Cell (" + DBFuncCell.Parent.Name + "!" + DBFuncCell.Address + "): " + ex.Message + ", this might be due to former errors in the VBA Macros (missing references)")
                    End Try
                End If
            Next
            If ignoreCalcMode And ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual Then
                LogInfo("ignoreCalcMode = True and Application.Calculation = xlCalculationManual, Application.CalculateFull called " + Wb.Path + "\" + Wb.Name)
                ExcelDnaUtil.Application.CalculateFull()
            End If
        Catch ex As Exception
            UserMsg("Exception: " + ex.Message + ", " + Wb.Path + "\" + Wb.Name, "refresh DBFunctions")
        End Try
        ' after all db function cells have been "dirtied" set calculation mode to automatic again (if it was automatic)
        If calcModeSet Then ExcelDnaUtil.Application.Calculation = calcMode
        DBModifHelper.preventChangeWhileFetching = False
    End Sub

    ''' <summary>"OnTime" event function to "escape" current (main) thread: event procedure to re-fetch DB functions results after triggering a recalculation inside Application.WorkbookBeforeSave</summary>
    Public Sub refreshDBFuncLater()
        Dim previouslySaved As Boolean
        Dim actWb As Excel.Workbook = Nothing
        Try : actWb = ExcelDnaUtil.Application.ActiveWorkbook : Catch ex As Exception
            LogWarn("Exception when trying to get the active workbook for refreshing DBfunc later: " + ex.Message + ", this might be due to errors in the VBA Macros (missing references)")
        End Try
        Try
            If actWb IsNot Nothing Then
                previouslySaved = actWb.Saved
                LogInfo("clearing DBfunction targets: refreshDBFunctions after clearing")
                refreshDBFunctions(actWb, True)
                actWb.Saved = previouslySaved
            End If
        Catch ex As Exception
            UserMsg("Exception: " + ex.Message, "refresh DBFunc later")
        End Try
    End Sub

    ''' <summary>create a ListObject one cell to the right of TargetCell and insert a dummy cmd sql definition for the list-object table (to be an external source)</summary>
    ''' <param name="TargetCell">the reference cell for the ListObject (will be the source cell for the DBSetQuery function)</param>
    Public Function createListObject(TargetCell As Excel.Range) As Object
        Dim createdQueryTable As Object
        ' if an alternate connection string is given for List-object, use this one...
        Dim altConnString = fetchSetting("AltConnString" + env(), "")
        ' To get the connection string work also for SQLOLEDB provider for SQL Server, change to ODBC driver setting (this can be generally used to fix connection string problems with ListObjects)
        If altConnString = "" Then altConnString = "OLEDB;" + Replace(ConstConnString, fetchSetting("ConnStringSearch" + env(), "provider=SQLOLEDB"), fetchSetting("ConnStringReplace" + env(), "driver=SQL SERVER"))
        Try
            createdQueryTable = TargetCell.Parent.ListObjects.Add(SourceType:=Excel.XlListObjectSourceType.xlSrcQuery, Source:=altConnString, Destination:=TargetCell.Offset(0, 1)).QueryTable
            With createdQueryTable
                .CommandType = Excel.XlCmdType.xlCmdSql
                .CommandText = fetchSetting("listobjectCmdTextToSet" + env(), "select CURRENT_TIMESTAMP") ' this should be sufficient for all ansi sql compliant databases
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .BackgroundQuery = False
                .RefreshStyle = Excel.XlCellInsertionMode.xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = False
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .Refresh(BackgroundQuery:=False)
            End With
        Catch ex As Exception
            UserMsg("Exception adding list-object query table:" + ex.Message, "Create List Object")
            createListObject = Nothing
            Exit Function
        End Try
        ' turn off auto-filter can't be done here because it leads to memory corruption...
        createListObject = createdQueryTable.ListObject
    End Function

    ''' <summary>create a pivot table object one cell below TargetCell and insert a dummy cmd sql definition for the pivot-cache external query</summary>>
    ''' <param name="TargetCell">the reference cell for the pivot table (will be the source cell for the DBSetQuery function)</param>
    Public Sub createPivotTable(TargetCell As Excel.Range)
        Dim pivotcache As Excel.PivotCache
        Dim pivotTables As Excel.PivotTables
        ' if an alternate connection string is given for List-object, use this one...
        Dim altConnString = fetchSetting("AltConnString" + env(), "")
        ' for standard connection strings only OLEDB drivers seem to work with pivot tables...
        If altConnString = "" Then altConnString = "OLEDB;" + ConstConnString
        Dim ExcelVersionForPivot As Integer = -1 : Dim versionTooHigh As Boolean = False
        Do
            ExcelVersionForPivot += 1
            Try
                ' don't use TargetCell.Parent.Parent.PivotCaches().Add(Excel.XlPivotTableSourceType.xlExternal) as we can't set the Version there...
                Dim dummy As Excel.PivotCache = ExcelDnaUtil.Application.ActiveWorkbook.PivotCaches.Create(SourceType:=Excel.XlPivotTableSourceType.xlExternal, Version:=ExcelVersionForPivot)
            Catch ex As Exception
                versionTooHigh = True
            End Try
        Loop Until versionTooHigh
        ExcelVersionForPivot -= 1
        pivotcache = ExcelDnaUtil.Application.ActiveWorkbook.PivotCaches.Create(SourceType:=Excel.XlPivotTableSourceType.xlExternal, Version:=ExcelVersionForPivot)
        'LogInfo("created pivot cache with version: " + CStr(ExcelVersionForPivot))
        Try
            pivotcache.Connection = altConnString
            pivotcache.MaintainConnection = False
        Catch ex As Exception
            UserMsg("Exception setting connection string for pivot cache: " + ex.Message, "Create Pivot Table")
        End Try
        ' set a minimum command text that should be sufficient for the database engine
        Dim pivotTableCmdTextToSet As String = fetchSetting("pivotTableCmdTextToSet" + env(), "select 1")
        Try
            pivotcache.CommandText = pivotTableCmdTextToSet
            pivotcache.CommandType = Excel.XlCmdType.xlCmdSql
        Catch ex As Exception
            UserMsg("Exception setting CommandText '" + pivotTableCmdTextToSet + "' for pivot cache: " + ex.Message, "Create Pivot Table")
        End Try

        Try
            pivotTables = TargetCell.Parent.PivotTables()
            pivotTables.Add(PivotCache:=pivotcache, TableDestination:=TargetCell.Offset(1, 0), DefaultVersion:=ExcelVersionForPivot)
        Catch ex As Exception
            UserMsg("Exception adding pivot table: " + ex.Message, "Create Pivot Table")
            Exit Sub
        End Try
    End Sub

    ''' <summary>creates functions in target cells (relative to referenceCell) as defined in ItemLineDef</summary>
    ''' <param name="originCell">original reference Cell</param>
    ''' <param name="ItemLineDef">String array, pairwise containing relative cell addresses and the functions in those cells (= cell content)</param>
    Public Sub createFunctionsInCells(originCell As Excel.Range, ByRef ItemLineDef As Object)
        Dim cellToBeStoredAddress As String, cellToBeStoredContent As String
        ' disabling calculation is necessary to avoid object errors
        Dim calcMode As Long = ExcelDnaUtil.Application.Calculation
        Try
            ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual
        Catch ex As Exception
            UserMsg("The Calculation mode can't be set, maybe you are in the formula/cell editor?", "Create Function In Cell")
            Exit Sub
        End Try
        If Functions.preventRefreshFlag Or (Functions.preventRefreshFlagColl.ContainsKey(ExcelDnaUtil.Application.ActiveWorkbook.Name) AndAlso Functions.preventRefreshFlagColl(ExcelDnaUtil.Application.ActiveWorkbook.Name)) Then
            UserMsg("preventRefresh is currently set, this affects creation of DB Functions, therefore disabling it now", "Create Function In Cell")
            Functions.preventRefreshFlag = False
            If Functions.preventRefreshFlagColl.ContainsKey(ExcelDnaUtil.Application.ActiveWorkbook.Name) Then Functions.preventRefreshFlagColl(ExcelDnaUtil.Application.ActiveWorkbook.Name) = False
            theRibbon.InvalidateControl("preventRefresh")
        End If

        Dim i As Long
        ' for each defined cell address and content pair
        For i = 0 To UBound(ItemLineDef) Step 2
            cellToBeStoredAddress = ItemLineDef(i)
            cellToBeStoredContent = ItemLineDef(i + 1)

            ' get cell in relation to function target cell
            If cellToBeStoredAddress.Length > 0 Then
                ' if there is a reference to a different sheet in cellToBeStoredAddress (starts with '<sheetname>'! ) and this sheet doesn't exist, create it...
                If InStr(1, cellToBeStoredAddress, "!") > 0 Then
                    Dim theSheetName As String = Replace(Mid$(cellToBeStoredAddress, 1, InStr(1, cellToBeStoredAddress, "!") - 1), "'", "")
                    Try
                        Dim testSheetExist As String = ExcelDnaUtil.Application.Worksheets(theSheetName).name
                    Catch ex As Exception
                        With ExcelDnaUtil.Application.Worksheets.Add(After:=originCell.Parent)
                            .name = theSheetName
                        End With
                        originCell.Parent.Activate()
                    End Try
                End If

                ' get target cell respecting relative cellToBeStoredAddress starting from originCell
                Dim TargetCell As Excel.Range = Nothing
                If Not getRangeFromRelative(originCell, cellToBeStoredAddress, TargetCell) Then
                    UserMsg("Excel Borders would be violated by placing target cell (relative address:" + cellToBeStoredAddress + ")" + vbLf + "Cell content: " + cellToBeStoredContent + vbLf + "Please select different cell !!")
                End If

                ' finally fill function target cell with function text (relative cell references to target cell) or value
                Try
                    If Left$(cellToBeStoredContent, 1) = "=" Then
                        TargetCell.FormulaR1C1 = cellToBeStoredContent
                    Else
                        TargetCell.Value = Left(cellToBeStoredContent, 32767)
                    End If
                Catch ex As Exception
                    UserMsg("Error in setting Cell: " + ex.Message, "Create functions in cells")
                End Try
            End If
        Next
        ExcelDnaUtil.Application.Calculation = calcMode
    End Sub

    ''' <summary>gets target range in relation to origin range</summary>
    ''' <param name="originCell">the origin cell to be related to</param>
    ''' <param name="relAddress">the relative address of the target as an RC style reference</param>
    ''' <param name="theTargetRange">the returned resulting range</param>
    ''' <returns>True if boundaries are not violated, false otherwise</returns>
    Private Function getRangeFromRelative(originCell As Excel.Range, ByVal relAddress As String, ByRef theTargetRange As Excel.Range) As Boolean
        Dim theSheetName As String

        If InStr(1, relAddress, "!") = 0 Then
            theSheetName = originCell.Parent.Name
        Else
            theSheetName = Replace(Mid$(relAddress, 1, InStr(1, relAddress, "!") - 1), "'", "")
        End If
        ' parse row or column out of RC style reference addresses
        Dim startRow As Long = 0, startCol As Long = 0, endRow As Long = 0, endCol As Long = 0
        Dim begins As String
        Dim relAddressPart() As String = Split(relAddress, ":")

        ' get startRow and startCol from both multi and single cell range (without separation by ":")
        If InStr(1, relAddressPart(0), "R[") > 0 Then
            begins = Mid$(relAddressPart(0), InStr(1, relAddressPart(0), "R[") + 2)
            startRow = CLng(Mid$(begins, 1, InStr(1, begins, "]") - 1))
        End If
        If InStr(1, relAddressPart(0), "C[") > 0 Then
            begins = Mid$(relAddressPart(0), InStr(1, relAddressPart(0), "C[") + 2)
            startCol = CLng(Mid$(begins, 1, InStr(1, begins, "]") - 1))
        End If
        ' get endRow and endCol in case of multi cell range ((topleftAddress):(bottomrightAddress))
        If UBound(relAddressPart) = 1 Then
            If InStr(1, relAddressPart(1), "R[") > 0 Then
                begins = Mid$(relAddressPart(1), InStr(1, relAddressPart(1), "R[") + 2)
                endRow = CLng(Mid$(begins, 1, InStr(1, begins, "]") - 1))
            End If
            If InStr(1, relAddressPart(1), "C[") > 0 Then
                begins = Mid$(relAddressPart(1), InStr(1, relAddressPart(1), "C[") + 2)
                endCol = CLng(Mid$(begins, 1, InStr(1, begins, "]") - 1))
            End If
        End If
        ' check if resulting target range would violate excel sheets boundaries, if so, then return error (false)
        If originCell.Row + startRow > 0 And originCell.Row + startRow <= originCell.Parent.Rows.Count _
           And originCell.Column + startCol > 0 And originCell.Column + startCol <= originCell.Parent.Columns.Count Then
            If InStr(1, relAddress, ":") > 0 Then
                ' for multi cell relative ranges, final target offset is starting at the bottom right of relative range
                theTargetRange = ExcelDnaUtil.Application.Range(originCell, originCell.Offset(endRow - startRow, endCol - startCol))
            Else
                ' for single cell relative ranges, target range is just set to the offsetting row and column of the relative range.
                theTargetRange = originCell
            End If
            theTargetRange = ExcelDnaUtil.Application.Worksheets(theSheetName).Range(theTargetRange.Offset(startRow, startCol).Address)
            getRangeFromRelative = True
        Else
            theTargetRange = Nothing
            getRangeFromRelative = False
        End If
    End Function

    ''' <summary>get a boolean type custom property</summary>
    ''' <param name="name">name of the property</param>
    ''' <param name="Wb">workbook of the property</param>
    ''' <returns>the value of the custom property</returns>
    Public Function getCustPropertyBool(name As String, Wb As Excel.Workbook) As Boolean
        Try
            getCustPropertyBool = Wb.CustomDocumentProperties(name).Value
        Catch ex As Exception
            getCustPropertyBool = False
        End Try
    End Function

    ''' <summary>converts a passed object (reference, value) to a boolean</summary>
    ''' <param name="value">object to be converted</param>
    ''' <returns>boolean result</returns>
    Public Function convertToBool(value As Object) As Boolean
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

    ''' <summary>takes an OADate and formats it as a DB Compliant string, using formatting as formatting instruction</summary>
    ''' <param name="datVal">OADate (double) date parameter</param>
    ''' <param name="formatting">formatting flag (see DBDate for details)</param>
    ''' <returns>formatted Date string</returns>
    Public Function formatDBDate(datVal As Double, formatting As Integer) As String
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

End Module