Imports Microsoft.Office.Interop.Excel

''' <summary>various functions used in most parts of DBAddin</summary>
Public Module CommonFuncs
    ''' <summary>splits theString into tokens delimited by delimiter, ignoring delimiters inside quotes and brackets</summary>
    ''' <param name="theString"></param>
    ''' <param name="delimiter"></param>
    ''' <param name="quote"></param>
    ''' <param name="startStr"></param>
    ''' <param name="openBracket"></param>
    ''' <param name="closeBracket"></param>
    ''' <returns>the list of tokens</returns>
    ''' <remarks>theString is split starting from startStr up to the first balancing closing Bracket (as defined by openBracket and closeBracket)
    ''' startStr, openBracket and closeBracket are case insensitive for comparing with theString.
    ''' the tokens are not blank trimmed !!</remarks>
    Public Function functionSplit(ByVal theString As String, delimiter As String, quote As String, startStr As String, openBracket As String, closeBracket As String) As Object
        Dim tempString As String
        Dim finalResult

        Try
            ' skip until we found startStr
            tempString = Mid$(theString, InStr(1, UCase$(theString), UCase$(startStr)) + Len(startStr))
            ' rip out the balancing string now...
            tempString = balancedString(tempString, openBracket, closeBracket, quote)
            If tempString.Length = 0 Then
                LogError("couldn't produce balanced string from " & theString)
                functionSplit = Nothing
                Exit Function
            End If
            tempString = replaceDelimsWithSpecialSep(tempString, delimiter, quote, openBracket, closeBracket, vbTab)
            finalResult = Split(tempString, vbTab)
            functionSplit = finalResult
        Catch ex As Exception
            WriteToLog("Error: " & ex.Message & " in CommonFuncs.functionSplit", EventLogEntryType.Warning)
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
            WriteToLog("Error: " & ex.Message & " in CommonFuncs.balancedString in ", EventLogEntryType.Warning)
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
    Private Function replaceDelimsWithSpecialSep(theString As String, delimiter As String, quote As String, openBracket As String, closeBracket As String, specialSep As String) As String
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
            WriteToLog("Error: " & ex.Message & " in CommonFuncs.replaceDelimsWithSpecialSep", EventLogEntryType.Warning)
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
            Dim dummy As String = hostApp.Worksheets(theName).name
        Catch ex As Exception
            existsSheet = False
        End Try
    End Function

    ''' <summary>checks whether ADO type theType is a date or time type</summary>
    ''' <param name="theType"></param>
    ''' <returns>True if DateTime</returns>
    Public Function checkIsDateTime(theType As ADODB.DataTypeEnum) As Boolean
        checkIsDateTime = False
        If theType = ADODB.DataTypeEnum.adDate Or theType = ADODB.DataTypeEnum.adDBDate Or theType = ADODB.DataTypeEnum.adDBTime Or theType = ADODB.DataTypeEnum.adDBTimeStamp Then
            checkIsDateTime = True
        End If
    End Function

    ''' <summary>checks whether ADO type theType is a date type</summary>
    ''' <param name="theType"></param>
    ''' <returns>True if Date</returns>
    Public Function checkIsDate(theType As ADODB.DataTypeEnum) As Boolean
        checkIsDate = False
        If theType = ADODB.DataTypeEnum.adDate Or theType = ADODB.DataTypeEnum.adDBDate Then
            checkIsDate = True
        End If
    End Function

    ''' <summary>checks whether ADO type theType is a time type</summary>
    ''' <param name="theType"></param>
    ''' <returns>True if Time</returns>
    Public Function checkIsTime(theType As ADODB.DataTypeEnum) As Boolean
        checkIsTime = False
        If theType = ADODB.DataTypeEnum.adDBTime Or theType = ADODB.DataTypeEnum.adDBTimeStamp Then
            checkIsTime = True
        End If
    End Function

    ''' <summary>checks whether ADO type theType is a numeric type</summary>
    ''' <param name="theType"></param>
    ''' <returns>True if numeric</returns>
    Public Function checkIsNumeric(theType As ADODB.DataTypeEnum) As Boolean
        checkIsNumeric = False
        If theType = ADODB.DataTypeEnum.adNumeric Or theType = ADODB.DataTypeEnum.adInteger Or theType = ADODB.DataTypeEnum.adTinyInt Or theType = ADODB.DataTypeEnum.adSmallInt Or theType = ADODB.DataTypeEnum.adBigInt Or theType = ADODB.DataTypeEnum.adUnsignedInt Or theType = ADODB.DataTypeEnum.adUnsignedTinyInt Or theType = ADODB.DataTypeEnum.adUnsignedSmallInt Or theType = ADODB.DataTypeEnum.adDouble Or theType = ADODB.DataTypeEnum.adSingle Or theType = ADODB.DataTypeEnum.adCurrency Or theType = ADODB.DataTypeEnum.adUnsignedBigInt Then
            checkIsNumeric = True
        End If
    End Function

    ''' <summary>gets first underlying Name that contains DBtarget or DBsource in theRange in theWb</summary>
    ''' <param name="theRange"></param>
    ''' <returns>the retrieved name</returns>
    Public Function getDBRangeName(theRange As Range) As Name
        Dim nm As Name
        Dim rng As Range
        Dim testRng As Range
        getDBRangeName = Nothing
        Try
            For Each nm In theRange.Parent.Parent.Names
                rng = Nothing
                Try : rng = nm.RefersToRange : Catch ex As Exception : End Try
                If Not rng Is Nothing And Not (nm.Name Like "*ExterneDaten*" Or nm.Name Like "*_FilterDatabase") Then
                    testRng = Nothing
                    Try : testRng = hostApp.Intersect(theRange, rng) : Catch ex As Exception : End Try
                    If Not IsNothing(testRng) And (InStr(1, nm.Name, "DBFtarget") >= 1 Or InStr(1, nm.Name, "DBFsource") >= 1) Then
                        getDBRangeName = nm
                        Exit Function
                    End If
                End If
            Next
        Catch ex As Exception
            WriteToLog("Error: " & Err.Description & " in CommonFuncs.getRangeName", EventLogEntryType.Warning)
        End Try
    End Function

    ''' <summary>only recalc full if we have DBFuncs in the workbook somewhere</summary>
    ''' <param name="Wb"></param>
    ''' <param name="ignoreCalcMode"></param>
    Public Sub refreshDBFunctions(Wb As Workbook, Optional ignoreCalcMode As Boolean = False)
        Dim searchCells As Range
        Dim ws As Worksheet
        Dim needRecalc As Boolean
        Dim theFunc

        If TypeName(hostApp.Calculation) = "Error" Then
            WriteToLog("hostApp.Calculation = Error, " & Wb.Path & "\" & Wb.Name, EventLogEntryType.Warning)
            Exit Sub
        End If
        Try
            needRecalc = False
            For Each ws In Wb.Worksheets
                For Each theFunc In {"DBListFetch(", "DBRowFetch(", "DBSetQuery("}
                    searchCells = ws.Cells.Find(What:=theFunc, After:=ws.Range("A1"), LookIn:=XlFindLookIn.xlFormulas, LookAt:=XlLookAt.xlPart, SearchOrder:=XlSearchOrder.xlByRows, SearchDirection:=XlSearchDirection.xlNext, MatchCase:=False)
                    If Not (searchCells Is Nothing) Then
                        ' reset the cell find dialog....
                        searchCells = Nothing
                        searchCells = ws.Cells.Find(What:="", After:=ws.Range("A1"), LookIn:=XlFindLookIn.xlFormulas, LookAt:=XlLookAt.xlPart, SearchOrder:=XlSearchOrder.xlByRows, SearchDirection:=XlSearchDirection.xlNext, MatchCase:=False)
                        needRecalc = True
                        GoTo done
                    End If
                Next
                ' reset the cell find dialog....
                searchCells = Nothing
                searchCells = ws.Cells.Find(What:="", After:=ws.Range("A1"), LookIn:=XlFindLookIn.xlFormulas, LookAt:=XlLookAt.xlPart, SearchOrder:=XlSearchOrder.xlByRows, SearchDirection:=XlSearchDirection.xlNext, MatchCase:=False)
            Next
done:
            If needRecalc And (hostApp.Calculation <> XlCalculation.xlCalculationManual Or ignoreCalcMode) Then
                WriteToLog("hostApp.CalculateFull called" & Wb.Path & "\" & Wb.Name, EventLogEntryType.Information)
                hostApp.CalculateFull()
            End If
        Catch ex As Exception
            WriteToLog("Error: " & ex.Message & ", " & Wb.Path & "\" & Wb.Name, EventLogEntryType.Warning)
        End Try
    End Sub

    ''' <summary>check whether key with name "tblName" is contained in collection tblColl</summary>
    ''' <param name="tblName">key to be found</param>
    ''' <param name="tblColl">collection to be checked</param>
    ''' <returns>if name was found in collection</returns>
    Public Function existsInCollection(tblName As String, tblColl As Collection) As Boolean
        existsInCollection = True
        Try
            Dim dummy As Integer = tblColl(tblName)
        Catch ex As Exception
            existsInCollection = False
        End Try
    End Function

    ''' <summary>"repairs" legacy functions from old VB6-COM Addin by removing "DBAddin.Functions." before function name</summary>
    Public Sub repairLegacyFunctions()
        Dim searchCell As Range
        Dim foundLegacy As Boolean
        Try
            Dim xlcalcmode As Long = hostApp.Calculation
            For Each ws In hostApp.ActiveWorkbook.Worksheets
                ' check whether legacy functions exist somewhere ...
                searchCell = ws.Cells.Find(What:="DBAddin.Functions.", After:=ws.Range("A1"), LookIn:=XlFindLookIn.xlFormulas, LookAt:=XlLookAt.xlPart, SearchOrder:=XlSearchOrder.xlByRows, SearchDirection:=XlSearchDirection.xlNext, MatchCase:=False)
                If Not (searchCell Is Nothing) Then foundLegacy = True
            Next
            If foundLegacy Then
                Dim retval As MsgBoxResult = MsgBox("Found Legacy DBAddin functions in Workbook, should they be replaced with current Addin functions (save Workbook afterwards to persist) ?", vbQuestion + vbYesNo, "Legacy DBAddin functions")
                If retval = vbYes Then
                    hostApp.Calculation = XlCalculation.xlCalculationManual ' avoid recalculations during repair...
                    hostApp.DisplayAlerts = False ' avoid warnings for sheet where "DBAddin.Functions." is not found
                    ' remove "DBAddin.Functions." in each sheet...
                    For Each ws In hostApp.ActiveWorkbook.Worksheets
                        ws.Cells.Replace(What:="DBAddin.Functions.", Replacement:="", LookAt:=XlLookAt.xlPart, SearchOrder:=XlSearchOrder.xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)
                    Next
                    hostApp.DisplayAlerts = True
                    hostApp.Calculation = xlcalcmode
                End If
            End If
            ' reset the cell find dialog....
            hostApp.ActiveSheet.Cells.Find(What:="", After:=hostApp.ActiveSheet.Range("A1"), LookIn:=XlFindLookIn.xlFormulas, LookAt:=XlLookAt.xlPart, SearchOrder:=XlSearchOrder.xlByRows, SearchDirection:=XlSearchDirection.xlNext, MatchCase:=False)
        Catch ex As Exception
            LogError("Error occured in repairLegacyFunctions: " & ex.Message)
        End Try
    End Sub

    ''' <summary> maintenance procedure to purge names used for dbfunctions from workbook</summary>
    Public Sub purgeNames()
        Dim resultingPurges As String = String.Empty
        Try
            Dim DBname As Name
            For Each DBname In hostApp.ActiveWorkbook.Names
                If DBname.Name Like "*ExterneDaten*" Or DBname.Name Like "*ExternalData*" Then
                    resultingPurges += DBname.Name + ","
                    DBname.Delete()
                ElseIf DBname.Name Like "DBListArea*" Then
                    resultingPurges += DBname.Name + ","
                    DBname.Delete()
                ElseIf DBname.Name Like "DBFtarget*" Then
                    resultingPurges += DBname.Name + ","
                    DBname.Delete()
                ElseIf DBname.Name Like "DBFsource*" Then
                    resultingPurges += DBname.Name + ","
                    DBname.Delete()
                ElseIf InStr(1, DBname.RefersTo, "#REF!") > 0 Then
                    resultingPurges += DBname.Name + ","
                    DBname.Delete()
                End If
            Next
            If resultingPurges = String.Empty Then
                MsgBox("nothing purged...", vbOKOnly, "purge Names")
            Else
                MsgBox("removed " + resultingPurges)
                WriteToLog("purgeNames removed " + resultingPurges, EventLogEntryType.Information)
            End If
        Catch ex As Exception
            LogError("Error occured in purgeNames: " & ex.Message)
        End Try
    End Sub

End Module