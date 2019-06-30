Imports Microsoft.Office.Interop
Imports System.IO ' for getting config files for menu

'''<summary>procedures used for loading config files (containing DBFunctions and general sheet content) and building the config menu</summary>
Public Module ConfigFiles
    Public referenceCell As Excel.Range

    '''<summary>get the current reference sheet (during display of the form/building db item)</summary>
    '''<returns>the current reference sheet</returns>
    Function referenceSheet() As Excel.Worksheet
        Return referenceCell.Parent
    End Function

    ''' <summary>loads config from file given in theFileName</summary>
    ''' <param name="theFileName">the File name of the config file</param>
    Public Sub loadConfig(theFileName As String)
        Dim ItemLine As String
        Dim retval As Integer

        On Error GoTo err1
        retval = MsgBox("Inserting contents configured in " & theFileName, vbInformation + vbOKCancel, "DBAddin: Inserting Configuration...")
        If retval = vbCancel Then Exit Sub
        If theHostApp.ActiveWorkbook Is Nothing Then theHostApp.Workbooks.Add
        ConfigFiles.referenceCell = theHostApp.ActiveCell

        ' open file for reading
        Dim fileReader As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(theFileName)
        Do
            ItemLine = fileReader.ReadLine()
            ' now insert the parsed information
            createFunctionsInCells(theHostApp.ActiveCell, Split(ItemLine, vbTab))
        Loop Until fileReader.EndOfStream
        fileReader.Close()
        Exit Sub

err1:
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
        LogError(Err.Description & " in ConfigFiles.createFunctionsInCells" & Erl())
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

    ''' <summary>used to create menu and button ids</summary>
    Private menuID As Integer
    ''' <summary>tree menu stored here</summary>
    Public ConfigMenuXML As String = vbNullString
    ''' <summary>max depth limitation by Ribbon: 5 levels: 1 top level, 1 folder level (Database foldername) -> 3 left</summary>
    Const specialFolderMaxDepth As Integer = 3
    ''' <summary>store found submenus in this collection</summary>
    Private specialConfigFoldersTempColl As Collection
    ''' <summary>for correct display of menu</summary>
    Private ReadOnly xnspace As XNamespace = "http://schemas.microsoft.com/office/2009/07/customui"

    Public Sub createConfigTreeMenu()
        Dim currentBar, button As XElement

        If Not Directory.Exists(ConfigStoreFolder) Then
            LogError("No predefined config store folder '" & ConfigStoreFolder & "' found, please correct setting and refresh!")
            ConfigMenuXML = "<menu xmlns='http://schemas.microsoft.com/office/2009/07/customui'><button id='refreshDBConfig' label='refresh DBConfig Tree' imageMso='Refresh' onAction='refreshDBConfigTree'/></menu>"
        Else
            ' top level menu
            currentBar = New XElement(xnspace + "menu")
            ' refresh button
            button = New XElement(xnspace + "button")
            button.SetAttributeValue("id", "refreshConfig")
            button.SetAttributeValue("label", "refresh DBConfig Tree")
            button.SetAttributeValue("imageMso", "Refresh")
            button.SetAttributeValue("onAction", "refreshDBConfigTree")
            currentBar.Add(button)
            specialConfigFoldersTempColl = New Collection
            menuID = 0
            readAllFiles(ConfigStoreFolder, currentBar)
            specialConfigFoldersTempColl = Nothing
            hostApp.StatusBar = String.Empty
            currentBar.SetAttributeValue("xmlns", "http://schemas.microsoft.com/office/2009/07/customui")
            ConfigMenuXML = currentBar.ToString()
        End If
    End Sub

    ''' <summary>reads all files contained in rootPath and its subfolders (recursively) and adds them to the DBConfig menu (sub)structure (recursively). For folders contained in specialConfigStoreFolders, apply further structuring by splitting names on camelcase or specialConfigStoreSeparator</summary>
    ''' <param name="rootPath"></param>
    ''' <param name="currentBar"></param>
    Private Sub readAllFiles(rootPath As String, ByRef currentBar As XElement)
        Dim newBar As XElement = Nothing
        Dim i As Long

        On Error GoTo err1
        ' read all leaf node entries (files) and sort them by name to create action menus
        Dim di As DirectoryInfo = New DirectoryInfo(rootPath)
        Dim fileList() As FileSystemInfo = di.GetFileSystemInfos("*.xcl").OrderBy(Function(fi) fi.Name).ToArray()
        If fileList.Length > 0 Then

            ' for special folders split further into camelcase (or other special ) separated names
            Dim aFolder : Dim spclFolder As String : spclFolder = String.Empty
            Dim theFolder As String
            theFolder = Mid$(rootPath, InStrRev(rootPath, "\") + 1)
            For Each aFolder In specialConfigStoreFolders
                If UCase$(theFolder) = UCase$(aFolder) Then
                    spclFolder = aFolder
                    Exit For
                End If
            Next
            If spclFolder.Length > 0 Then
                Dim firstCharLevel As Boolean = CBool(fetchSetting(spclFolder & "FirstLetterLevel", "False"))
                Dim specialConfigStoreSeparator As String = fetchSetting(spclFolder & "Separator", String.Empty)
                Dim nameParts As String
                For i = 0 To UBound(fileList)
                    ' is current entry contained in next entry then revert order to allow for containment in next entry's hierarchy..
                    If i < UBound(fileList) Then
                        If InStr(1, Left$(fileList(i + 1).Name, Len(fileList(i + 1).Name) - 4), Left$(fileList(i).Name, Len(fileList(i).Name) - 4)) > 0 Then
                            nameParts = stringParts(IIf(firstCharLevel, Left$(fileList(i + 1).Name, 1) & " ", String.Empty) &
                                                            Left$(fileList(i + 1).Name, Len(fileList(i + 1).Name) - 4), specialConfigStoreSeparator)
                            buildFileSepMenuCtrl(nameParts, currentBar, rootPath & "\" & fileList(i + 1).Name, spclFolder, specialFolderMaxDepth)
                            nameParts = stringParts(IIf(firstCharLevel, Left$(fileList(i).Name, 1) & " ", String.Empty) &
                                                            Left$(fileList(i).Name, Len(fileList(i).Name) - 4), specialConfigStoreSeparator)
                            buildFileSepMenuCtrl(nameParts, currentBar, rootPath & "\" & fileList(i).Name, spclFolder, specialFolderMaxDepth)
                            i += 2
                            If i > UBound(fileList) Then Exit For
                        End If
                    End If
                    nameParts = stringParts(IIf(firstCharLevel, Left$(fileList(i).Name, 1) & " ", String.Empty) &
                            Left$(fileList(i).Name, Len(fileList(i).Name) - 4), specialConfigStoreSeparator)
                    buildFileSepMenuCtrl(nameParts, currentBar, rootPath & "\" & fileList(i).Name, spclFolder, specialFolderMaxDepth)
                Next
                ' normal case: just follow the path and enter all entries as buttons
            Else
                For i = 0 To UBound(fileList)
                    newBar = New XElement(xnspace + "button")
                    menuID += 1
                    newBar.SetAttributeValue("id", "m" + menuID.ToString())
                    newBar.SetAttributeValue("tag", rootPath + "\" & fileList(i).Name)
                    newBar.SetAttributeValue("label", Left$(fileList(i).Name, Len(fileList(i).Name) - 4))
                    newBar.SetAttributeValue("onAction", "getConfig")
                    currentBar.Add(newBar)
                Next
            End If
        End If

        ' read all folder xcl entries and sort them by name
        Dim DirList() As DirectoryInfo = di.GetDirectories().OrderBy(Function(fi) fi.Name).ToArray()
        If DirList.Length = 0 Then Exit Sub
        ' recursively build branched menu structure from dirEntries
        For i = 0 To UBound(DirList)
            hostApp.StatusBar = "Filling DBConfigs Menu: " & rootPath & "\" & DirList(i).Name
            newBar = New XElement(xnspace + "menu")
            menuID += 1
            newBar.SetAttributeValue("id", "m" + menuID.ToString())
            newBar.SetAttributeValue("label", DirList(i).Name)
            currentBar.Add(newBar)
            readAllFiles(rootPath & "\" & DirList(i).Name, newBar)
        Next
        Exit Sub
err1:
        LogToEventViewer("Error (" & Err.Description & ") in MenuHandler.readAllFiles in " & Erl(), EventLogEntryType.Error)
    End Sub

    ''' <summary>parses Substrings contained in nameParts (recursively) and adds them to currentBar and submenus (recursively)</summary>
    ''' <param name="nameParts" >tokenized string (separated by space)</param>
    ''' <param name="currentBar"></param>
    ''' <param name="fullPathName"></param>
    ''' <param name="newRootName"></param>
    ''' <param name="specialFolderMaxDepth"></param>
    Private Sub buildFileSepMenuCtrl(nameParts As String, ByRef currentBar As XElement,
                                         fullPathName As String, newRootName As String, specialFolderMaxDepth As Integer)
        Dim newBar As XElement
        Static currentDepth As Integer

        On Error GoTo buildFileSepMenuCtrl_Err
        ' end node: add callable entry (= button)
        If InStr(1, nameParts, " ") = 0 Or currentDepth > specialFolderMaxDepth - 1 Then
            Dim entryName As String = Mid$(fullPathName, InStrRev(fullPathName, "\") + 1)
            newBar = New XElement(xnspace + "button")
            menuID += 1
            newBar.SetAttributeValue("id", "m" + menuID.ToString())
            newBar.SetAttributeValue("label", Left$(entryName, Len(entryName) - 4))
            newBar.SetAttributeValue("tag", fullPathName)
            newBar.SetAttributeValue("onAction", "getConfig")
            currentBar.Add(newBar)
        Else  ' branch node: add new menu, recursively descend
            Dim newName As String = Left$(nameParts, InStr(1, nameParts, " ") - 1)
            ' prefix already exists: put new submenu below already existing prefix
            If specialConfigFoldersTempColl.Contains(newRootName & newName) Then
                newBar = specialConfigFoldersTempColl(newRootName & newName)
            Else
                newBar = New XElement(xnspace + "menu")
                menuID += 1
                newBar.SetAttributeValue("id", "m" + menuID.ToString())
                newBar.SetAttributeValue("label", newName)
                specialConfigFoldersTempColl.Add(newBar, newRootName & newName)
                currentBar.Add(newBar)
            End If
            currentDepth += 1
            buildFileSepMenuCtrl(Mid$(nameParts, InStr(1, nameParts, " ") + 1), newBar, fullPathName, newRootName & newName, specialFolderMaxDepth)
            currentDepth -= 1
        End If
        Exit Sub

buildFileSepMenuCtrl_Err:
        LogToEventViewer("Error (" & Err.Description & ") in MenuHandler.buildFileSepMenuCtrl in " & Erl(), EventLogEntryType.Error)
    End Sub

    ''' <summary>returns string in space separated parts (tokenize String following CamelCase switch or when given specialConfigStoreSeparator occurs)</summary>
    Private Function stringParts(theString As String, specialConfigStoreSeparator As String) As String
        stringParts = String.Empty
        ' specialConfigStoreSeparator given: split by it
        If specialConfigStoreSeparator.Length > 0 Then
            stringParts = Join(Split(theString, specialConfigStoreSeparator), " ")
        Else ' walk through string, separating by camelcase switch
            Dim CamelCaseStrLen As Integer = Len(theString)
            Dim i As Integer
            For i = 1 To CamelCaseStrLen
                Dim aChar As String = Mid$(theString, i, 1)
                Dim charAsc As Integer = Asc(aChar)

                If i > 1 Then
                    ' character before current character
                    Dim pre As Integer = Asc(Mid$(theString, i - 1, 1))
                    ' underscore also separates camelcase, except preceded by $, - or another underscore
                    If charAsc = 95 Then
                        If Not (pre = 36 Or pre = 45 Or pre = 95) _
                            Then stringParts &= " "
                    End If
                    ' Uppercase characters separate unless the are preceding
                    If (charAsc >= 65 And charAsc <= 90) Then
                        If Not (pre >= 65 And pre <= 90) _
                           And Not (pre = 36 Or pre = 45 Or pre = 95) _
                           And Not (pre >= 48 And pre <= 57) _
                           Then stringParts &= " "
                    End If
                End If
                stringParts &= aChar
            Next
            stringParts = LTrim$(Replace(Replace(stringParts, "   ", " "), "  ", " "))
        End If
    End Function

End Module
