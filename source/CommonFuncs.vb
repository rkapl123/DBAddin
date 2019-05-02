Imports Microsoft.Office.Interop.Excel
Imports ExcelDna.Integration
Imports ADODB

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

        On Error GoTo err1
        ' skip until we found startStr
        tempString = Mid$(theString, InStr(1, UCase$(theString), UCase$(startStr)) + Len(startStr))
        ' rip out the balancing string now...
        tempString = balancedString(tempString, openBracket, closeBracket, quote)
        If tempString.Length = 0 Then GoTo err0
        tempString = replaceDelimsWithSpecialSep(tempString, delimiter, quote, openBracket, closeBracket, vbTab)
        finalResult = Split(tempString, vbTab)

        functionSplit = finalResult
        Exit Function
err0:
        Err.Raise(1, , "couldn't produce balanced string from " & theString)
        On Error Resume Next
        Exit Function
err1:
        Dim errDesc As String = Err.Description
        LogToEventViewer("Error: " & errDesc & " in CommonFuncs.functionSplit in " & Erl(), EventLogEntryType.Error)
        functionSplit = Nothing
    End Function


    Sub testfunctionSplit()
        Dim check

        check = functionSplit("ignored, because it is before opener..,func(token3,'(', token4,internalfunc(next,next))&this is also ignored, because we have closed the bracket", ",", "'", "func", "(", ")")
        Debug.Print(check(0) = "token3")
        Debug.Print(check(1) = "'('")
        Debug.Print(check(2) = " token4")
        Debug.Print(check(3) = "internalfunc(next,next)")
        Debug.Print(UBound(check) = 3)
        Debug.Print("")

        ' watch out, startStr really searches for the first occurrence ("func") !!
        check = functionSplit("ignoredfunction(because,it,is,before,opener)&func(token3,'(', token4,internalfunc(next,next))&this is also ignored, because we have closed the bracket", ",", "'", "func", "(", ")")
        Debug.Print(check(0) <> "token3")
        Debug.Print(check(1) <> "'('")
        Debug.Print(check(2) <> " token4")
        Debug.Print(check(3) <> "internalfunc(next,next)")
        Debug.Print(UBound(check) <> 3)
        Debug.Print("")

        check = functionSplit("ignored(because,it,is,before,opener)&func(token3,'(', token4,internalfunc(next,next))&this is also ignored, because we have closed the bracket", ",", "'", "func", "(", ")")
        Debug.Print(check(0) = "token3")
        Debug.Print(check(1) = "'('")
        Debug.Print(check(2) = " token4")
        Debug.Print(check(3) = "internalfunc(next,next)")
        Debug.Print(UBound(check) = 3)
        Debug.Print("")

        check = functionSplit("func(token3,'(ignore,ignore),whatever is inside'&(still ignored, because in brackets), token4,internalfunc(arg1,anotherFunc(arg1,arg2),arg2))&this is also ignored, because we have closed the bracket", ",", "'", "func", "(", ")")
        Debug.Print(check(0) = "token3")
        Debug.Print(check(1) = "'(ignore,ignore),whatever is inside'&(still ignored, because in brackets)")
        Debug.Print(check(2) = " token4")
        Debug.Print(check(3) = "internalfunc(arg1,anotherFunc(arg1,arg2),arg2)")
        Debug.Print(UBound(check) = 3)
        Debug.Print("")

        ' a different quote and a different delimiter:
        check = functionSplit("=func(token1;token2;""ignoredcloseBracket)""; token3;""ignored1;ignored2"");ignored1;ignored2", ";", """", "func", "(", ")")
        Debug.Print(check(0) = "token1")
        Debug.Print(check(1) = "token2")
        Debug.Print(check(2) = """ignoredcloseBracket)""")
        Debug.Print(check(3) = " token3")
        Debug.Print(check(4) = """ignored1;ignored2""")
        Debug.Print(UBound(check) = 4)
        Debug.Print("")
    End Sub

    ''' <summary>returns the minimal bracket balancing string contained in theString, opening bracket defined in openBracket, closing bracket defined in closeBracket
    ''' disregarding quoted areas inside optionally given quote charachter/string</summary>
    ''' <param name="theString"></param>
    ''' <param name="openBracket"></param>
    ''' <param name="closeBracket"></param>
    ''' <param name="quote"></param>
    ''' <returns>the balanced string</returns>
    Private Function balancedString(theString As String, openBracket As String, closeBracket As String, Optional quote As String = "") As String
        Dim startBalance As Long, endBalance As Long, i As Long, countOpen As Long, countClose As Long

        Dim quoteMode As Boolean = False
        On Error GoTo err1
        startBalance = 0
        For i = 1 To Len(theString)
            If Left$(Mid$(theString, i), Len(quote)) = quote And quote.Length > 0 And Not quoteMode Then
                quoteMode = True
            Else
                If Not quoteMode Then
                    If Left$(Mid$(theString, i), Len(openBracket)) = openBracket Then
                        If startBalance = 0 Then startBalance = i
                        countOpen = countOpen + 1
                    End If
                    If startBalance <> 0 And Left$(Mid$(theString, i), Len(closeBracket)) = closeBracket Then countClose = countClose + 1
                Else
                    If Left$(Mid$(theString, i), Len(quote)) = quote And quote.Length > 0 Then quoteMode = False
                End If
            End If

            If countOpen = countClose And startBalance <> 0 Then
                endBalance = i - 1
                Exit For
            End If
        Next
        If endBalance = 0 Then
            balancedString = String.Empty
        Else
            balancedString = Mid$(theString, startBalance + 1, endBalance - startBalance)
        End If
        Exit Function
err1:
        If VBDEBUG Then Debug.Print("balancedString: " & Err.Description) : Stop : Resume
        LogToEventViewer("Error: " & Err.Description & " in CommonFuncs.balancedString in " & Erl(), EventLogEntryType.Error)
    End Function


    Private Sub testBalanced()
        Debug.Print(balancedString("ignored,(start,""ignore '(' , but include"",(go on, the end)),this should (all()) be excluded", "(", ")", """") = "start,""ignore '(' , but include"",(go on, the end)")
        Debug.Print(balancedString("""(ignored"",(start,""ignore '(' , but include"",(go on, the end)),this should (all) be excluded", "(", ")", """") = "start,""ignore '(' , but include"",(go on, the end)")
    End Sub

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
        replaceDelimsWithSpecialSep = String.Empty

        Dim i As Long
        On Error GoTo err1
        For i = 1 To Len(theString)
            If Left$(Mid$(theString, i), Len(quote)) = quote And quote.Length > 0 And Not quoteMode Then
                quoteMode = True
            Else
                If quoteMode And Left$(Mid$(theString, i), Len(quote)) = quote And quote.Length > 0 Then quoteMode = False
            End If

            If Left$(Mid$(theString, i), Len(openBracket)) = openBracket And openBracket.Length > 0 And Not quoteMode Then
                openedBrackets = openedBrackets + 1
            End If
            If Left$(Mid$(theString, i), Len(closeBracket)) = closeBracket And closeBracket.Length > 0 And Not quoteMode Then
                openedBrackets = openedBrackets - 1
            End If

            If Not (openedBrackets > 0 Or quoteMode) Then
                If Left$(Mid$(theString, i), Len(delimiter)) = delimiter Then
                    replaceDelimsWithSpecialSep = replaceDelimsWithSpecialSep & specialSep
                Else
                    replaceDelimsWithSpecialSep = replaceDelimsWithSpecialSep & Mid$(theString, i, 1)
                End If
            Else
                replaceDelimsWithSpecialSep = replaceDelimsWithSpecialSep & Mid$(theString, i, 1)
            End If
        Next
        Exit Function
err1:
        If VBDEBUG Then Debug.Print("replaceDelimsWithSpecialSep: " & Err.Description) : Stop : Resume
        LogToEventViewer("Error: " & Err.Description & " in CommonFuncs.replaceDelimsWithSpecialSep in " & Erl(), EventLogEntryType.Error)
    End Function


    ''' <summary>changes theString by replacing substring starting after keystr and ending with separator with changed, case insensitive !!</summary>
    ''' <param name="theString"></param>
    ''' <param name="keystr"></param>
    ''' <param name="changed"></param>
    ''' <param name="separator"></param>
    ''' <returns>the changed string</returns>
    Public Function Change(ByVal theString As String, ByVal keystr As String, ByVal changed As String, ByVal separator As String) As String
        Dim replaceBeg As Integer, replaceEnd As Integer

        replaceBeg = InStr(1, UCase$(theString), UCase$(keystr))
        If replaceBeg = 0 Then
            Change = String.Empty
            Exit Function
        End If
        replaceEnd = InStr(replaceBeg, UCase$(theString), UCase$(separator))
        If replaceEnd = 0 Then replaceEnd = Len(theString) + 1
        Change = Left$(theString, replaceBeg - 1 + Len(keystr)) & changed & Right$(theString, Len(theString) - replaceEnd + 1)
    End Function


    ''' <summary>fetches substring starting after keystr and ending with separator from theString, case insensitive !!
    ''' if separator is "" then fetch to end of string</summary>
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
        Dim dummy As String

        existsSheet = True
        On Error Resume Next
        Err.Clear()
        dummy = theHostApp.Worksheets(theName).name
        If Err.Number <> 0 Then existsSheet = False
    End Function

    ''' <summary>checks whether ADO type theType is a date or time type</summary>
    ''' <param name="theType"></param>
    ''' <returns>True if DateTime</returns>
    Public Function checkIsDateTime(theType As ADODB.DataTypeEnum) As Boolean
        If theType = ADODB.DataTypeEnum.adDate Or theType = ADODB.DataTypeEnum.adDBDate Or theType = ADODB.DataTypeEnum.adDBTime Or theType = ADODB.DataTypeEnum.adDBTimeStamp Then
            checkIsDateTime = True
        Else
            checkIsDateTime = False
        End If
    End Function

    ''' <summary>checks whether ADO type theType is a date type</summary>
    ''' <param name="theType"></param>
    ''' <returns>True if Date</returns>
    Public Function checkIsDate(theType As ADODB.DataTypeEnum) As Boolean
        If theType = ADODB.DataTypeEnum.adDate Or theType = ADODB.DataTypeEnum.adDBDate Then
            checkIsDate = True
        Else
            checkIsDate = False
        End If
    End Function

    ''' <summary>checks whether ADO type theType is a time type</summary>
    ''' <param name="theType"></param>
    ''' <returns>True if Time</returns>
    Public Function checkIsTime(theType As ADODB.DataTypeEnum) As Boolean
        If theType = ADODB.DataTypeEnum.adDBTime Or theType = ADODB.DataTypeEnum.adDBTimeStamp Then
            checkIsTime = True
        Else
            checkIsTime = False
        End If
    End Function

    ''' <summary>checks whether ADO type theType is a numeric type</summary>
    ''' <param name="theType"></param>
    ''' <returns>True if numeric</returns>
    Public Function checkIsNumeric(theType As ADODB.DataTypeEnum) As Boolean
        If theType = ADODB.DataTypeEnum.adNumeric Or theType = ADODB.DataTypeEnum.adInteger Or theType = ADODB.DataTypeEnum.adTinyInt Or theType = ADODB.DataTypeEnum.adSmallInt Or theType = ADODB.DataTypeEnum.adBigInt Or theType = ADODB.DataTypeEnum.adUnsignedInt Or theType = ADODB.DataTypeEnum.adUnsignedTinyInt Or theType = ADODB.DataTypeEnum.adUnsignedSmallInt Or theType = ADODB.DataTypeEnum.adDouble Or theType = ADODB.DataTypeEnum.adSingle Or theType = ADODB.DataTypeEnum.adCurrency Or theType = ADODB.DataTypeEnum.adUnsignedBigInt Then
            checkIsNumeric = True
        Else
            checkIsNumeric = False
        End If
    End Function

    ''' <summary>gets first underlying Name that contains DBtarget or DBsource in theRange in theWb</summary>
    ''' <param name="theRange"></param>
    ''' <returns>the retrieved name</returns>
    Public Function getDBRangeName(theRange As Range) As Name
        Dim nm As Name
        Dim rng As Range
        Dim testRng As Range

        On Error GoTo err1
        For Each nm In theRange.Parent.Parent.Names
            rng = Nothing
            On Error Resume Next
            rng = nm.RefersToRange
            On Error GoTo err1
            If Not rng Is Nothing And Not (nm.Name Like "*ExterneDaten*" Or nm.Name Like "*_FilterDatabase") Then
                On Error Resume Next
                testRng = theHostApp.Intersect(theRange, rng)
                If Err.Number = 0 And Not testRng Is Nothing And (InStr(1, nm.Name, "DBFtarget") >= 1 Or InStr(1, nm.Name, "DBFsource") >= 1) Then
                    getDBRangeName = nm
                    Err.Clear()
                    Exit Function
                End If
                On Error GoTo err1
            End If
        Next
        getDBRangeName = Nothing
        Exit Function
err1:
        If VBDEBUG Then Debug.Print("getRangeName: " & Err.Description) : Stop : Resume
        LogToEventViewer("Error: " & Err.Description & " in CommonFuncs.getRangeName in " & Erl(), EventLogEntryType.Error)
    End Function

    ''' <summary>only recalc full if we have DBFuncs in the workbook somewhere</summary>
    ''' <param name="Wb"></param>
    ''' <param name="ignoreCalcMode"></param>
    Public Sub refreshDBFunctions(Wb As Workbook, Optional ignoreCalcMode As Boolean = False)
        Dim searchCells As Range
        Dim ws As Worksheet
        Dim needRecalc As Boolean
        Dim theFunc

        If TypeName(theHostApp.Calculation) = "Error" Then
            LogToEventViewer("refreshDBFunctions: theHostApp.Calculation = Error, " & Wb.Path & "\" & Wb.Name, EventLogEntryType.Information)
            Exit Sub
        End If
        On Error GoTo err_1
        needRecalc = False
        For Each ws In Wb.Worksheets
            For Each theFunc In {"DBListFetch(", "DBRowFetch(", "DBCellFetch(", "DBMakeControl(", "DBSetQuery("}
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
        If needRecalc And (theHostApp.Calculation <> XlCalculation.xlCalculationManual Or ignoreCalcMode) Then
            LogToEventViewer("refreshDBFunctions: theHostApp.CalculateFull called" & Wb.Path & "\" & Wb.Name, EventLogEntryType.Information)
            theHostApp.CalculateFull
        End If
        Exit Sub
err_1:
        LogToEventViewer("Error: " & Err.Description & " in CommonFuncs.refreshDBFunctions in " & Erl() & ", " & Wb.Path & "\" & Wb.Name, EventLogEntryType.Error)
    End Sub

    ''' <summary>formats theVal to fit the type of column theHead in recordset checkrst</summary>
    ''' <param name="theVal"></param>
    ''' <param name="theHead"></param>
    ''' <param name="checkrst"></param>
    ''' <returns>formatted value</returns>
    Public Function dbformat(ByVal theVal As Object, ByVal theHead As String, checkrst As Recordset) As String
        ' build where clause criteria..
        If checkIsNumeric(checkrst.Fields(theHead).Type) Then
            dbformat = Replace(CStr(theVal), ",", ".")
        ElseIf checkIsDate(checkrst.Fields(theHead).Type) Then
            dbformat = "'" & Format$(theVal, "YYYYMMDD") & "'"
        ElseIf checkIsTime(checkrst.Fields(theHead).Type) Then
            dbformat = "'" & Format$(theVal, "YYYYMMDD HH:MM:SS") & "'"
        ElseIf TypeName(theVal) = "Boolean" Then
            dbformat = IIf(theVal, 1, 0)
        Else
            ' quote strings
            theVal = Replace(theVal, "'", "''")
            dbformat = "'" & theVal & "'"
        End If
    End Function

End Module