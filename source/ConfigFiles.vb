Imports Microsoft.Office.Interop
Imports ExcelDna.Integration
''
'  procedures used for loading config files (storing DBFuncs and general sheet content)
Public Module ConfigFiles
    Public referenceCell As Excel.Range

    ''
    ' get the current reference sheet (during display of the form/building db item)
    ' @return the current reference sheet
    ' @remarks
    Function referenceSheet() As Excel.Worksheet
        Return referenceCell.Parent
    End Function

    ''' <summary>loads config from file given in theFileName</summary>
    ''' <param name="theFileName">the File name of the config file</param>
    Public Sub loadConfig(theFileName As String)
        Dim ItemLine As String
        Dim retval As Integer

        On Error GoTo err1
        retval = MsgBox("Inserting contents configured in " & theFileName, vbInformation + vbOKCancel, "DBAddin: Inserting Configuration...") 'Ctrl.Parameter
        If retval = vbCancel Then Exit Sub
        If theHostApp.ActiveWorkbook Is Nothing Then theHostApp.Workbooks.Add

        ' open file for reading
        Dim fileReader As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(theFileName)
        Do
            ItemLine = fileReader.ReadLine()
            ' now insert the parsed information
            createFunctionsInCells(theHostApp.ActiveCell, Split(ItemLine, vbTab))
        Loop Until EOF(1)
        fileReader.Close()
        Exit Sub

err1:
        If VBDEBUG Then Debug.Print("Error (" & Err.Description & ") in ConfigFiles.loadConfig") : Stop : Resume
        LogToEventViewer("Error (" & Err.Description & ") using filename '" & theFileName & "' in ConfigFiles.loadConfig" & " in " & Erl(), EventLogEntryType.Error)
    End Sub


    ''' <summary>creates functions in target cells (relative to referenceCell) as defined in ItemLineDef</summary>
    ''' <param name="referenceCell">reference Cell where all functions relative addresses are related to</param>
    ''' <param name="ItemLineDef">String array, pairwise containing relative cell addresses and the functions in those cells (= cell content)</param>
    Public Sub createFunctionsInCells(referenceCell As Excel.Range, ByRef ItemLineDef As Object)
        On Error GoTo err1

        Dim cellToBeStoredAddress As String, cellToBeStoredContent As String
        ' disabling calculation is necessary to avoid object errors
        Dim calcMode As Long : calcMode = theHostApp.Calculation
        theHostApp.Calculation = Excel.XlCalculation.xlCalculationManual
        Dim i As Long

        ' for each defined cell address and content pair
        For i = 0 To UBound(ItemLineDef) Step 2
            cellToBeStoredAddress = ItemLineDef(i)
            cellToBeStoredContent = ItemLineDef(i + 1)

            ' get cell in relation to function target cell
            If cellToBeStoredAddress.Length > 0 Then
                'if targetsheet for the cell doesn't exist, create it...
                createSheetForTarget(cellToBeStoredAddress)
                Err.Clear()

                ' finally fill function target cell with function text (relative cell references) or value
                Dim TargetCell As Excel.Range
                TargetCell = Nothing

                If Not getRangeFromRelative(referenceCell, cellToBeStoredAddress, TargetCell) Then
                    LogWarn("Excel Borders would be violated by placing target cell (relative address:" & cellToBeStoredAddress & ")" & vbLf & "Cell content: " & cellToBeStoredContent & vbLf & "Please select different cell !!", 1)
                    GoTo cleanup
                End If
                On Error Resume Next

                If Left$(cellToBeStoredContent, 1) = "=" Then
                    TargetCell.FormulaR1C1 = cellToBeStoredContent
                Else
                    TargetCell.Value = cellToBeStoredContent
                End If

                If Err.Number <> 0 Then
                    LogWarn("Error in setting Cell: " & Err.Description, 1)
                    GoTo cleanup
                End If

                ' for dbcellfetch wraptext makes sense !!
                If InStr(1, UCase$(cellToBeStoredContent), "DBCELLFETCH(") > 0 Then
                    TargetCell.WrapText = True
                End If
            End If
        Next
cleanup:
        theHostApp.Calculation = calcMode
        Exit Sub
err1:
        If VBDEBUG Then Debug.Print("Error (" & Err.Description & ") in ConfigFiles.createFunctionsInCells") : Stop : Resume
        LogError(Err.Description & " in ConfigFiles.createFunctionsInCells" & Erl(), , , 1)
    End Sub


    ''' <summary>creates a sheet if theTarget is specifying to be in a different worksheet (theTarget starts with '(sheetname)'! )</summary>
    ''' <param name="theTarget"></param>
    Private Sub createSheetForTarget(ByVal theTarget As String)
        Dim theSheetName As String
        Dim testSheetExist As String

        If InStr(1, theTarget, "!") = 0 Then Exit Sub
        theSheetName = Replace(Mid$(theTarget, 1, InStr(1, theTarget, "!") - 1), "'", String.Empty)
        On Error Resume Next
        testSheetExist = theHostApp.Worksheets(theSheetName).name
        If Err.Number <> 0 Then
            With theHostApp.Worksheets.Add(After:=referenceSheet())
                .name = theSheetName
            End With
            referenceSheet().Activate()
        End If
    End Sub

    ''' <summary>gets range in relation to another (originRange)</summary>
    ''' <param name="originRange">the origin to be related to</param>
    ''' <param name="relAddress">the relative address of the target</param>
    ''' <param name="theTargetRange">the returned range</param>
    ''' <returns>True if no errors, false otherwise</returns>
    Private Function getRangeFromRelative(originRange As Excel.Range, ByVal relAddress As String, ByRef theTargetRange As Excel.Range) As Boolean
        Dim theSheetName As String

        If InStr(1, relAddress, "!") = 0 Then
            theSheetName = referenceSheet().Name
        Else
            theSheetName = Replace(Mid$(relAddress, 1, InStr(1, relAddress, "!") - 1), "'", String.Empty)
        End If
        Dim startRow As Long, startCol As Long
        startRow = getRowOrCol(relAddress, True)
        startCol = getRowOrCol(relAddress, False)
        If originRange.Row + startRow > 0 And originRange.Row + startRow <= referenceSheet().Rows.Count _
           And originRange.Column + startCol > 0 And originRange.Column + startCol <= referenceSheet().Columns.Count Then
            If InStr(1, relAddress, ":") > 0 Then
                Dim endRow As Long, endCol As Long
                endRow = getRowOrCol(relAddress, True, True)
                endCol = getRowOrCol(relAddress, False, True)
                ' extend origin range to size of relAddress (being then set to theTargetRange)
                theTargetRange = theHostApp.Range(originRange, originRange.Offset(endRow - startRow, endCol - startCol))
            Else
                theTargetRange = originRange
            End If
            theTargetRange = theHostApp.Worksheets(theSheetName).Range(theTargetRange.Offset(startRow, startCol).Address)
            getRangeFromRelative = True
        Else
            theTargetRange = Nothing
            getRangeFromRelative = False
        End If
    End Function

    ''' <summary>parse row or column out of RC style reference adresses</summary>
    ''' <param name="relAddr">RC style reference adresses</param>
    ''' <param name="getRow">get the row (true) or column (false)</param>
    ''' <param name="getBottomRight">if we have a multi cell range ((topleftAddress):(bottomrightAddress)) then get the row or column from the bottomright part</param>
    ''' <returns>parsed row (getRow = true) or column (getRow = false) from address</returns>
    Function getRowOrCol(relAddr As String, getRow As Boolean, Optional getBottomRight As Boolean = False) As Long
        Dim beg As String, srchSubStr As String, srchBeg As Integer

        srchSubStr = IIf(getRow, "R[", "C[")
        srchBeg = 0
        getRowOrCol = 0
        If getBottomRight Then
            srchBeg = InStr(1, relAddr, ":")
            If srchBeg = 0 Then Exit Function
        Else
            If InStr(1, relAddr, srchSubStr) > InStr(1, relAddr, ":") And InStr(1, relAddr, ":") > 0 Then Exit Function
        End If
        If InStr(srchBeg + 1, relAddr, srchSubStr) = 0 Then
            Exit Function
        Else
            beg = Mid$(relAddr, InStr(srchBeg + 1, relAddr, srchSubStr) + 2)
            getRowOrCol = CLng(Mid$(beg, 1, InStr(1, beg, "]") - 1))
        End If
    End Function

End Module
