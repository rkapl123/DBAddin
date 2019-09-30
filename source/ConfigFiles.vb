Imports Microsoft.Office.Interop
Imports System.IO ' for getting config files for menu
Imports System.Xml.Linq
Imports System.Linq

'''<summary>procedures used for loading config files (containing DBFunctions and general sheet content) and building the config menu</summary>
Public Module ConfigFiles
    ''' <summary>the folder used to store predefined DB item definitions</summary>
    Public ConfigStoreFolder As String
    ''' <summary>Array of special ConfigStoreFolders for non default treatment of Name Separation (Camelcase) and max depth</summary>
    Public specialConfigStoreFolders() As String

    ''' <summary>loads config from file given in theFileName</summary>
    ''' <param name="theFileName">the File name of the config file</param>
    Public Sub loadConfig(theFileName As String)
        Dim ItemLine As String
        Dim retval As Integer

        Dim srchdFunc As String = ""
        ' check whether there is any existing db function other than DBListFetch inside active cell
        For Each srchdFunc In {"DBSETQUERY", "DBROWFETCH"}
            If Left(UCase(hostApp.ActiveCell.Formula), Len(srchdFunc) + 2) = "=" & srchdFunc & "(" Then
                Exit For
            Else
                srchdFunc = ""
            End If
        Next

        retval = MsgBox("Inserting contents configured in " & theFileName, vbInformation + vbOKCancel, "DBAddin: Inserting Configuration...")
        If retval = vbCancel Then Exit Sub
        If hostApp.ActiveWorkbook Is Nothing Then hostApp.Workbooks.Add

        ' open file for reading
        Try
            Dim fileReader As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(theFileName)
            Do
                ItemLine = fileReader.ReadLine()
                ' for existing dbfunction replace querystring in existing formula of active cell
                If srchdFunc <> "" Then
                    hostApp.ActiveCell.Formula = replaceQueryInFormula(Split(ItemLine, vbTab)(1), srchdFunc, hostApp.ActiveCell.Formula)
                Else ' for other cells simply insert the parsed information
                    createFunctionsInCells(hostApp.ActiveCell, Split(ItemLine, vbTab))
                End If
            Loop Until fileReader.EndOfStream
            fileReader.Close()
        Catch ex As Exception
            ErrorMsg("Error (" & ex.Message & ") during filling items from config file '" & theFileName & "' in ConfigFiles.loadConfig")
        End Try
    End Sub

    ''' <summary>replace query given in theQueryFormula inside sourceFormula containing DB Function "theFunction"</summary>
    ''' <param name="theQueryFormula"></param>
    ''' <param name="theFunction"></param>
    ''' <param name="sourceFormula"></param>
    ''' <returns></returns>
    Private Function replaceQueryInFormula(theQueryFormula As String, theFunction As String, sourceFormula As Object) As String
        Dim queryString As String = functionSplit(theQueryFormula, ",", """", "DBListFetch", "(", ")")(0)
        Dim formulaBody As String = Mid$(sourceFormula, Len(theFunction) + 3)
        formulaBody = Left(formulaBody, Len(formulaBody) - 1)
        Dim tempFormula As String = replaceDelimsWithSpecialSep(formulaBody, ",", """", "(", ")", vbTab)
        Dim restFormula As String = Mid$(tempFormula, InStr(tempFormula, vbTab))
        ' for existing DB Functions DBSetQuery or DBRowFetch...
        ' replace querystring in existing formula of active cell
        replaceQueryInFormula = "=" & theFunction & "(" & queryString & Replace(restFormula, vbTab, ",")
    End Function

    ''' <summary>creates functions in target cells (relative to referenceCell) as defined in ItemLineDef</summary>
    ''' <param name="originCell">original reference Cell</param>
    ''' <param name="ItemLineDef">String array, pairwise containing relative cell addresses and the functions in those cells (= cell content)</param>
    Public Sub createFunctionsInCells(originCell As Excel.Range, ByRef ItemLineDef As Object)
        Dim cellToBeStoredAddress As String, cellToBeStoredContent As String
        ' disabling calculation is necessary to avoid object errors
        Dim calcMode As Long = hostApp.Calculation
        hostApp.Calculation = Excel.XlCalculation.xlCalculationManual
        Dim i As Long

        ' for each defined cell address and content pair
        For i = 0 To UBound(ItemLineDef) Step 2
            cellToBeStoredAddress = ItemLineDef(i)
            cellToBeStoredContent = ItemLineDef(i + 1)

            ' get cell in relation to function target cell
            If cellToBeStoredAddress.Length > 0 Then
                ' if there is a reference to a different sheet in cellToBeStoredAddress (starts with '<sheetname>'! ) and this sheet doesn't exist, create it...
                If InStr(1, cellToBeStoredAddress, "!") > 0 Then
                    Dim theSheetName As String = Replace(Mid$(cellToBeStoredAddress, 1, InStr(1, cellToBeStoredAddress, "!") - 1), "'", String.Empty)
                    Try
                        Dim testSheetExist As String = hostApp.Worksheets(theSheetName).name
                    Catch ex As Exception
                        With hostApp.Worksheets.Add(After:=originCell.Parent)
                            .name = theSheetName
                        End With
                        originCell.Parent.Activate()
                    End Try
                End If

                ' get target cell respecting relative cellToBeStoredAddress starting from originCell
                Dim TargetCell As Excel.Range = Nothing
                If Not getRangeFromRelative(originCell, cellToBeStoredAddress, TargetCell) Then
                    ErrorMsg("Excel Borders would be violated by placing target cell (relative address:" & cellToBeStoredAddress & ")" & vbLf & "Cell content: " & cellToBeStoredContent & vbLf & "Please select different cell !!")
                End If

                ' finally fill function target cell with function text (relative cell references to target cell) or value
                Try
                    If Left$(cellToBeStoredContent, 1) = "=" Then
                        TargetCell.FormulaR1C1 = cellToBeStoredContent
                    Else
                        TargetCell.Value = cellToBeStoredContent
                    End If
                Catch ex As Exception
                    MsgBox("Error in setting Cell: " & ex.Message)
                End Try
            End If
        Next
        hostApp.Calculation = calcMode
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
            theSheetName = Replace(Mid$(relAddress, 1, InStr(1, relAddress, "!") - 1), "'", String.Empty)
        End If
        ' parse row or column out of RC style reference adresses
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
                theTargetRange = hostApp.Range(originCell, originCell.Offset(endRow - startRow, endCol - startCol))
            Else
                ' for single cell relative ranges, target range is just set to the offsetting row and column of the relative range.
                theTargetRange = originCell
            End If
            theTargetRange = hostApp.Worksheets(theSheetName).Range(theTargetRange.Offset(startRow, startCol).Address)
            getRangeFromRelative = True
        Else
            theTargetRange = Nothing
            getRangeFromRelative = False
        End If
    End Function


    ''' <summary>used to create menu and button ids</summary>
    Private menuID As Integer
    ''' <summary>tree menu stored here</summary>
    Public ConfigMenuXML As String = vbNullString
    ''' <summary>max depth limitation by Ribbon: 5 levels: 1 top level, 1 folder level (Database foldername) -> 3 left</summary>
    Public specialFolderMaxDepth As Integer
    ''' <summary>store found submenus in this collection</summary>
    Private specialConfigFoldersTempColl As Collection
    ''' <summary>for correct display of menu</summary>
    Private ReadOnly xnspace As XNamespace = "http://schemas.microsoft.com/office/2009/07/customui"

    ''' <summary>creates the Config tree menu by reading the menu elements from the config store folder files/subfolders</summary>
    Public Sub createConfigTreeMenu()
        Dim currentBar, button As XElement

        If Not Directory.Exists(ConfigStoreFolder) Then
            ErrorMsg("No predefined config store folder '" & ConfigStoreFolder & "' found, please correct setting and refresh!")
            ConfigMenuXML = "<menu xmlns='http://schemas.microsoft.com/office/2009/07/customui'><button id='refreshDBConfig' label='refresh DBConfig Tree' imageMso='Refresh' onAction='refreshDBConfigTree'/></menu>"
        Else
            ' top level menu
            currentBar = New XElement(xnspace + "menu")
            ' add refresh button to top level
            button = New XElement(xnspace + "button")
            button.SetAttributeValue("id", "refreshConfig")
            button.SetAttributeValue("label", "refresh DBConfig Tree")
            button.SetAttributeValue("imageMso", "Refresh")
            button.SetAttributeValue("onAction", "refreshDBConfigTree")
            currentBar.Add(button)
            ' collect all config files recursively, creating submenus for the structure (see readAllFiles) and buttons for the final config files.
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
    ''' <param name="rootPath">root folder to be searched for config files</param>
    ''' <param name="currentBar">current menu element, where submenus and buttons are added</param>
    Private Sub readAllFiles(rootPath As String, ByRef currentBar As XElement)
        Try
            Dim newBar As XElement = Nothing
            Dim i As Long
            ' read all leaf node entries (files) and sort them by name to create action menus
            Dim di As DirectoryInfo = New DirectoryInfo(rootPath)
            Dim fileList() As FileSystemInfo = di.GetFileSystemInfos("*.xcl").OrderBy(Function(fi) fi.Name).ToArray()
            If fileList.Length > 0 Then
                ' for special folders split menu further into camelcase (or other special) separated names
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
                    ' max depth limitation by Ribbon: 5 levels: 1 top level, 1 folder level (Database foldername) -> 3 left
                    specialFolderMaxDepth = IIf(fetchSetting(spclFolder & "MaxDepth", 1) <= 3, fetchSetting(spclFolder & "MaxDepth", 1), 3)
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
                        nameParts = stringParts(IIf(firstCharLevel, Left$(fileList(i).Name, 1) & " ", String.Empty) & Left$(fileList(i).Name, Len(fileList(i).Name) - 4), specialConfigStoreSeparator)
                        buildFileSepMenuCtrl(nameParts, currentBar, rootPath & "\" & fileList(i).Name, spclFolder, specialFolderMaxDepth)
                    Next
                    ' normal case: just follow the path and enter all entries as buttons
                Else
                    For i = 0 To UBound(fileList)
                        newBar = New XElement(xnspace + "button")
                        menuID += 1
                        newBar.SetAttributeValue("id", "m" + menuID.ToString())
                        newBar.SetAttributeValue("screentip", "click to insert DBListFetch for " & Left$(fileList(i).Name, Len(fileList(i).Name) - 4) & " in active cell")
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
        Catch ex As Exception
            ErrorMsg("Error (" & ex.Message & ") in MenuHandler.readAllFiles")
        End Try
    End Sub

    ''' <summary>for special Folders, parses Substrings contained in nameParts (recursively) of passed config file (fullPathName) and adds them to currentBar and submenus (recursively)</summary>
    ''' <param name="nameParts">tokenized string (separated by space)</param>
    ''' <param name="currentBar">current menu element, where submenus and buttons are added</param>
    ''' <param name="fullPathName">full path name to config file</param>
    ''' <param name="newRootName">the new root name for the menu, used avoid multiple placement of buttons in submenus</param>
    ''' <param name="specialFolderMaxDepth">limit for menu depth (required technically - depth limitation by Ribbon: 5 levels - and sometimes practically)</param>
    Private Sub buildFileSepMenuCtrl(nameParts As String, ByRef currentBar As XElement, fullPathName As String, newRootName As String, specialFolderMaxDepth As Integer)
        Try
            Dim newBar As XElement
            Static currentDepth As Integer
            ' end node: add callable entry (= button)
            If InStr(1, nameParts, " ") = 0 Or currentDepth > specialFolderMaxDepth - 1 Then
                Dim entryName As String = Mid$(fullPathName, InStrRev(fullPathName, "\") + 1)
                newBar = New XElement(xnspace + "button")
                menuID += 1
                newBar.SetAttributeValue("id", "m" + menuID.ToString())
                newBar.SetAttributeValue("screentip", "click to insert DBListFetch for " & Left$(entryName, Len(entryName) - 4) & " in active cell")
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
        Catch ex As Exception
            ErrorMsg("Error (" & ex.Message & ") in MenuHandler.buildFileSepMenuCtrl")
        End Try
    End Sub

    ''' <summary>returns string in space separated parts (tokenize String following CamelCase switch or when given specialConfigStoreSeparator occurs)</summary>
    ''' <param name="theString">string to be tokenized</param>
    ''' <param name="specialConfigStoreSeparator">if not empty, tokenize theString by this separator, else tokenize by camel case</param>
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
