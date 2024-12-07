Imports ExcelDna.Integration
Imports System.IO ' for getting config files for menu
Imports System.Linq ' to enhance arrays with useful methods (count, orderby)
Imports System.Xml.Linq ' XNamespace and XElement for constructing the ConfigMenuXML
Imports System.Collections.Generic 'ConfigDocCollection

'''<summary>procedures used for loading config files (containing DBFunctions and general sheet content) and building the config menu</summary>
Public Module ConfigFiles

    ''' <summary>the folder used to store predefined DB item definitions</summary>
    Public ConfigStoreFolder As String
    ''' <summary>Array of special ConfigStoreFolders for non default treatment of Name Separation (Camel-case) and max depth</summary>
    Public specialConfigStoreFolders() As String
    ''' <summary>fixed max Depth for Ribbon</summary>
    Const maxMenuDepth As Integer = 5
    ''' <summary>fixed max size for menu XML</summary>
    Const maxSizeRibbonMenu = 320000
    ''' <summary>used to create menu and button ids</summary>
    Private menuID As Integer
    ''' <summary>tree menu stored here</summary>
    Public ConfigMenuXML As String = vbNullString
    ''' <summary>individual limitation of grouping of entries in special folders (set by _DBname_MaxDepth)</summary>
    Public specialFolderMaxDepth As Integer
    ''' <summary>store found sub-menus in this collection</summary>
    Private specialConfigFoldersTempColl As Collection
    ''' <summary>for correct display of menu</summary>
    Private ReadOnly xnspace As XNamespace = "http://schemas.microsoft.com/office/2009/07/customui"
    ''' <summary>Documentation Collection for Config Objects (to be displayed with Ctrl or Shift)</summary>
    Public ConfigDocCollection As Dictionary(Of String, String)

    ''' <summary>loads config from file given in theFileName</summary>
    ''' <param name="theFileName">the File name of the config file</param>
    Public Sub loadConfig(theFileName As String)
        Dim ItemLine As String
        Dim retval As Integer

        Dim srchdFunc As String = ""
        ' check whether there is any existing db function other than DBListFetch inside active cell
        For Each srchdFunc In {"DBSETQUERY", "DBROWFETCH"}
            If Left(UCase(ExcelDnaUtil.Application.ActiveCell.Formula), Len(srchdFunc) + 2) = "=" + srchdFunc + "(" Then
                Exit For
            Else
                srchdFunc = ""
            End If
        Next

        retval = QuestionMsg("Inserting contents configured in " + theFileName, MsgBoxStyle.OkCancel, "DBAddin: Inserting Configuration...", MsgBoxStyle.Information)
        If retval = vbCancel Then Exit Sub
        If ExcelDnaUtil.Application.ActiveWorkbook Is Nothing Then ExcelDnaUtil.Application.Workbooks.Add

        ' open file for reading
        Try
            Dim fileReader As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(theFileName, Text.Encoding.Default)
            Do
                ItemLine = fileReader.ReadLine()
                ' ConfigArray: Configs are tab separated pairs of <RC location vbTab function formula> vbTab <...> vbTab...
                Dim ConfigArray As String() = Split(ItemLine, vbTab)
                ' if there is a ConfigSelect setting use it to replace the query with the template, replacing the contained table with the FROM <table>...
                ' also regard the possibility to have a preference for a specific ConfigSelect(1, 2, or any other postfix being available in settings)
                Dim ConfigSelect As String = fetchSetting("ConfigSelect" + fetchSetting("ConfigSelectPreference", ""), "")
                If ConfigSelect = "" Then ConfigSelect = fetchSetting("ConfigSelect", "") ' if nothing found under given ConfigSelectPreference, fall back to standard ConfigSelect
                ' replace query in function formula in second part of pairs with ConfigSelect template. 
                ' This works only for templates with actual query string as first argument (not having reference(s) to cell(s) with query string(s))
                ' also only works for single pair config templates
                If ConfigSelect <> "" And ConfigArray.Count() = 2 Then
                    If Left(UCase(ExcelDnaUtil.Application.ActiveCell.Formula), Len(srchdFunc) + 2) = "=DBLISTFETCH(" Then
                        ConfigArray(1) = replaceConfigSelectInFormula(ConfigArray(1), ConfigSelect)
                    End If
                End If
                ' for existing dbfunction replace query-string in existing formula of active cell, only works for single pair config templates
                If srchdFunc <> "" And ConfigArray.Count() = 2 Then
                    ExcelDnaUtil.Application.ActiveCell.Formula = replaceQueryInFormula(ConfigArray(1), srchdFunc, ExcelDnaUtil.Application.ActiveCell.Formula.ToString())
                Else ' for other cells simply insert the ConfigArray
                    createFunctionsInCells(ExcelDnaUtil.Application.ActiveCell, ConfigArray)
                End If
            Loop Until fileReader.EndOfStream
            fileReader.Close()
        Catch ex As Exception
            UserMsg("Error (" + ex.Message + ") during filling items from config file '" + theFileName + "' in ConfigFiles.loadConfig")
        End Try
    End Sub

    ''' <summary>replace query given in theQueryFormula with template query in ConfigSelect</summary>
    ''' <param name="dbFunctionFormula"></param>
    ''' <param name="ConfigSelect"></param>
    ''' <returns></returns>
    Private Function replaceConfigSelectInFormula(dbFunctionFormula As String, ConfigSelect As String) As String
        ' get the query from the config templates function formula (standard templates are created with DBListFetch)
        Dim queryString As String
        Dim functionParts As String() = functionSplit(dbFunctionFormula, ",", """", "DBListFetch", "(", ")")
        If functionParts IsNot Nothing Then
            queryString = functionParts(0)
            ' fetch table-name from query string
            Dim tableName As String = Mid$(queryString, InStr(queryString.ToUpper, "FROM ") + 5)
            ' remove last quoting...
            tableName = Left(tableName, Len(tableName) - 1)
            ' now replace table template with actual table name
            queryString = ConfigSelect.Replace("!Table!", tableName)
            ' reconstruct the rest of the db function formula
            Dim formulaParams As String = Mid$(dbFunctionFormula, Len("DBListFetch") + 3)
            formulaParams = Left(formulaParams, Len(formulaParams) - 1)
            Dim tempFormula As String = replaceDelimsWithSpecialSep(formulaParams, ",", """", "(", ")", vbTab)
            Dim restFormula As String = Mid$(tempFormula, InStr(tempFormula, vbTab))
            ' replace query-string in existing formula
            replaceConfigSelectInFormula = "=DBListFetch(""" + queryString + """" + Replace(restFormula, vbTab, ",") + ")"
        Else
            ' when problems occurred, leave everything as is
            replaceConfigSelectInFormula = dbFunctionFormula
        End If
    End Function

    ''' <summary>replace query given in dbFunctionFormula inside targetFormula containing DB Function "theFunction"</summary>
    ''' <param name="dbFunctionFormula">passed config templates function formula</param>
    ''' <param name="theFunction">db function in targetFormula</param>
    ''' <param name="targetFormula">passed ActiveCell.Formula</param>
    ''' <returns></returns>
    Private Function replaceQueryInFormula(dbFunctionFormula As String, theFunction As String, targetFormula As String) As String
        ' get the query from the config templates function formula (standard templates are created with DBListFetch)
        Dim queryString As String
        Dim functionParts As String() = functionSplit(dbFunctionFormula, ",", """", "DBListFetch", "(", ")")
        If functionParts IsNot Nothing Then
            queryString = functionParts(0)
            ' get the parts of the targeted function formula
            Dim formulaParams As String = Mid$(targetFormula, Len(theFunction) + 3)
            formulaParams = Left(formulaParams, Len(formulaParams) - 1)
            Dim tempFormula As String = replaceDelimsWithSpecialSep(formulaParams, ",", """", "(", ")", vbTab)
            Dim restFormula As String = Mid$(tempFormula, InStr(tempFormula, vbTab))
            ' for existing theFunction (DBSetQuery or DBRowFetch)...
            ' replace query-string in existing formula and pass as result
            replaceQueryInFormula = "=" + theFunction + "(" + queryString + Replace(restFormula, vbTab, ",") + ")"
        Else
            ' when problems occurred, leave everything as is
            replaceQueryInFormula = targetFormula
        End If
    End Function

    ''' <summary>creates the Config tree menu by reading the menu elements from the config store folder files/sub-folders</summary>
    Public Sub createConfigTreeMenu()
        Dim currentBar, button As XElement

        ' also get the documentation that was provided in setting ConfigDocQuery into ConfigDocCollection (used in config menu when clicking entry + Ctrl/Shift)
        Dim ConfigDocQuery As String = fetchSetting("ConfigDocQuery" + env(), fetchSetting("ConfigDocQuery", ""))
        If ConfigDocQuery <> "" Then ConfigDocCollection = getConfigDocCollection(ConfigDocQuery)

        ' get the .xcl config files from the folders beneath ConfigStoreFolder
        If Not Directory.Exists(ConfigStoreFolder) Then
            UserMsg("No predefined config store folder '" + ConfigStoreFolder + "' found, please correct setting and refresh!")
            ConfigMenuXML = "<menu xmlns='" + xnspace.ToString() + "'><button id='refreshDBConfig' label='refresh DBConfig Tree' imageMso='Refresh' onAction='refreshDBConfigTree'/></menu>"
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
            ' collect all config files recursively, creating sub-menus for the structure (see readAllFiles) and buttons for the final config files.
            specialConfigFoldersTempColl = New Collection
            menuID = 0
            readAllFiles(ConfigStoreFolder, currentBar)
            specialConfigFoldersTempColl = Nothing
            ExcelDnaUtil.Application.StatusBar = ""
            currentBar.SetAttributeValue("xmlns", xnspace)
            ' avoid exception in ribbon...
            ConfigMenuXML = currentBar.ToString()
            If ConfigMenuXML.Length > maxSizeRibbonMenu Then
                UserMsg("Too many entries in " + ConfigStoreFolder + ", can't display them in a ribbon menu ..")
                ConfigMenuXML = "<menu xmlns='" + xnspace.ToString() + "'><button id='refreshDBConfig' label='refresh DBConfig Tree' imageMso='Refresh' onAction='refreshDBConfigTree'/></menu>"
            End If
        End If
    End Sub

    ''' <summary>reads all files contained in rootPath and its sub-folders (recursively) and adds them to the DBConfig menu (sub)structure (recursively). For folders contained in specialConfigStoreFolders, apply further structuring by splitting names on camel-case or specialConfigStoreSeparator</summary>
    ''' <param name="rootPath">root folder to be searched for config files</param>
    ''' <param name="currentBar">current menu element, where sub-menus and buttons are added</param>
    ''' <param name="Folderpath">for sub menus path of current folder is passed (recursively)</param>
    Private Sub readAllFiles(rootPath As String, ByRef currentBar As XElement, Optional Folderpath As String = vbNullString)
        Try
            Dim newBar As XElement = Nothing
            Static MenuFolderDepth As Integer = 1 ' needed to not exceed max. menu depth (currently 5)

            ' read all leaf node entries (files) and sort them by name to create action menus
            Dim di As New DirectoryInfo(rootPath)
            Dim fileList() As FileSystemInfo = di.GetFileSystemInfos("*.xcl").OrderBy(Function(fi) fi.Name).ToArray()
            If fileList.Length > 0 Then
                ' for special folders split menu further into camel-case (or other special) separated names
                Dim spclFolder As String = ""
                For Each aFolder As String In specialConfigStoreFolders
                    ' is current folder contained in special config folders?
                    If UCase$(Mid$(rootPath, InStrRev(rootPath, "\") + 1)) = UCase$(aFolder) Then
                        spclFolder = aFolder
                        Exit For
                    End If
                Next
                If spclFolder <> "" And MenuFolderDepth < maxMenuDepth Then
                    Dim firstCharLevel As Boolean = fetchSettingBool(spclFolder + "FirstLetterLevel", "False")
                    Dim specialConfigStoreSeparator As String = fetchSetting(spclFolder + "Separator", "")
                    specialFolderMaxDepth = fetchSettingInt(spclFolder + "MaxDepth", "4")
                    Dim nameParts As String
                    For i As Long = 0 To UBound(fileList)
                        ' is current entry contained in next entry then revert order to allow for containment in next entry's hierarchy..
                        ' e.g. SpecialTable and SpecialTableDetails (and afterwards SpecialTableMoreDetails) -> SpecialTable opens hierarchy
                        If i < UBound(fileList) Then
                            Dim nextEntry As String = Left(fileList(i + 1).Name, Len(fileList(i + 1).Name) - 4)
                            Dim thisEntry As String = Left(fileList(i).Name, Len(fileList(i).Name) - 4)
                            Dim firstCharNextEntry As String = Left$(fileList(i + 1).Name, 1)
                            Dim firstCharThisEntry As String = Left$(fileList(i).Name, 1)
                            If InStr(1, nextEntry, thisEntry) > 0 Then
                                ' first process NEXT alphabetically ordered entry, returning nextLevel as new command bar element (menu or button)
                                nameParts = stringParts(IIf(firstCharLevel, firstCharNextEntry + " ", "") + nextEntry, specialConfigStoreSeparator)
                                Dim nextLevel As XElement = buildFileSepMenuCtrl(nameParts, currentBar, rootPath + "\" + fileList(i + 1).Name, spclFolder, Folderpath, MenuFolderDepth, specialFolderMaxDepth)
                                ' only if a menu was created...
                                If Right(nextLevel.Name.ToString(), 4) = "menu" Then
                                    ' ... process THIS entry and insert in to nextLevel
                                    nameParts = stringParts(IIf(firstCharLevel, firstCharThisEntry + " ", "") + thisEntry, specialConfigStoreSeparator)
                                    buildFileSepMenuCtrl(nameParts, nextLevel, rootPath + "\" + fileList(i).Name, spclFolder, Folderpath, MenuFolderDepth, specialFolderMaxDepth)
                                Else
                                    ' otherwise insert THIS entry in the same level (currentBar)
                                    buildFileSepMenuCtrl(nameParts, currentBar, rootPath + "\" + fileList(i).Name, spclFolder, Folderpath, MenuFolderDepth, specialFolderMaxDepth)
                                End If
                                ' skip this and next one
                                i += 2
                                If i > UBound(fileList) Then Exit For
                            End If
                        End If
                        nameParts = stringParts(IIf(firstCharLevel, Left$(fileList(i).Name, 1) + " ", "") + Left$(fileList(i).Name, Len(fileList(i).Name) - 4), specialConfigStoreSeparator)
                        buildFileSepMenuCtrl(nameParts, currentBar, rootPath + "\" + fileList(i).Name, spclFolder, Folderpath, MenuFolderDepth, specialFolderMaxDepth)
                    Next
                    ' normal case or max menu depth branch: just follow the path and enter all entries as buttons
                Else
                    For i = 0 To UBound(fileList)
                        newBar = New XElement(xnspace + "button")
                        menuID += 1
                        newBar.SetAttributeValue("id", "m" + menuID.ToString())
                        newBar.SetAttributeValue("screentip", "click to insert DBListFetch for " + Left$(fileList(i).Name, Len(fileList(i).Name) - 4) + " in active cell. Ctrl or Shift + click to display documentation for config if existing.")
                        newBar.SetAttributeValue("tag", rootPath + "\" + fileList(i).Name)
                        newBar.SetAttributeValue("label", Folderpath + Left$(fileList(i).Name, Len(fileList(i).Name) - 4))
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
                ExcelDnaUtil.Application.StatusBar = "Filling DBConfigs Menu: " + rootPath + "\" + DirList(i).Name
                ' only add new menu element if below max. menu depth for ribbons
                If MenuFolderDepth < maxMenuDepth Then
                    newBar = New XElement(xnspace + "menu")
                    menuID += 1
                    newBar.SetAttributeValue("id", "m" + menuID.ToString())
                    newBar.SetAttributeValue("label", DirList(i).Name)
                    currentBar.Add(newBar)
                    MenuFolderDepth += 1
                    readAllFiles(rootPath + "\" + DirList(i).Name, newBar, Folderpath + DirList(i).Name + "\")
                    MenuFolderDepth -= 1
                Else
                    newBar = currentBar
                    readAllFiles(rootPath + "\" + DirList(i).Name, newBar, Folderpath + DirList(i).Name + "\")
                End If
            Next
        Catch ex As Exception
            UserMsg("Error (" + ex.Message + ") in MenuHandler.readAllFiles")
        End Try
    End Sub

    ''' <summary>parses Substrings (filenames in special Folders) contained in nameParts (recursively) of passed xcl config file-path (fullPathName) and adds them to currentBar and sub-menus (recursively)</summary>
    ''' <param name="nameParts">tokenized string (separated by space)</param>
    ''' <param name="currentBar">current menu element, where sub-menus and buttons are added</param>
    ''' <param name="fullPathName">full path name to xcl config file</param>
    ''' <param name="newRootName">the new root name for the menu, used avoid multiple placement of buttons in sub-menus</param>
    ''' <param name="Folderpath">Path of enclosing Folder(s)</param>
    ''' <param name="MenuFolderDepth">required for keeping maxMenuDepth limit</param>
    ''' <returns>new bar as Xelement (for containment)</returns>
    Private Function buildFileSepMenuCtrl(nameParts As String, ByRef currentBar As XElement, fullPathName As String, newRootName As String, Folderpath As String, MenuFolderDepth As Integer, specialFolderMaxDepth As Integer) As XElement
        Static MenuDepth As Integer = 0
        Try
            Dim newBar As XElement
            ' end node: add callable entry (= button)
            If InStr(1, nameParts, " ") = 0 Or MenuDepth >= specialFolderMaxDepth Or MenuDepth + MenuFolderDepth >= maxMenuDepth Then
                Dim entryName As String = Mid$(fullPathName, InStrRev(fullPathName, "\") + 1)
                newBar = New XElement(xnspace + "button")
                menuID += 1
                newBar.SetAttributeValue("id", "m" + menuID.ToString())
                newBar.SetAttributeValue("screentip", "click to insert DBListFetch for " + Left$(entryName, Len(entryName) - 4) + " in active cell. Ctrl or Shift + click to display documentation for config if existing.")
                newBar.SetAttributeValue("label", Left$(entryName, Len(entryName) - 4))
                newBar.SetAttributeValue("tag", fullPathName)
                newBar.SetAttributeValue("onAction", "getConfig")
                currentBar.Add(newBar)
                buildFileSepMenuCtrl = newBar
            Else  ' branch node: add new menu, recursively descend
                Dim newName As String = Left$(nameParts, InStr(1, nameParts, " ") - 1)
                ' prefix already exists: put new sub-menu below already existing prefix
                If specialConfigFoldersTempColl.Contains(newRootName + newName) Then
                    newBar = specialConfigFoldersTempColl(newRootName + newName)
                Else
                    newBar = New XElement(xnspace + "menu")
                    menuID += 1
                    newBar.SetAttributeValue("id", "m" + menuID.ToString())
                    newBar.SetAttributeValue("label", newName)
                    specialConfigFoldersTempColl.Add(newBar, newRootName + newName)
                    currentBar.Add(newBar)
                End If
                MenuDepth += 1
                buildFileSepMenuCtrl(Mid$(nameParts, InStr(1, nameParts, " ") + 1), newBar, fullPathName, newRootName + newName, Folderpath, MenuFolderDepth, specialFolderMaxDepth)
                MenuDepth -= 1
                buildFileSepMenuCtrl = newBar
            End If
        Catch ex As Exception
            UserMsg("Error (" + ex.Message + ") in MenuHandler.buildFileSepMenuCtrl")
            buildFileSepMenuCtrl = Nothing
        End Try
    End Function

    ''' <summary>returns string in space separated parts (tokenize String following CamelCase switch or when given specialConfigStoreSeparator occurs)</summary>
    ''' <param name="theString">string to be tokenized</param>
    ''' <param name="specialConfigStoreSeparator">if not empty, tokenize theString by this separator, else tokenize by camel case</param>
    Private Function stringParts(theString As String, specialConfigStoreSeparator As String) As String
        stringParts = ""
        ' specialConfigStoreSeparator given: split by it
        If specialConfigStoreSeparator.Length > 0 Then
            stringParts = Join(Split(theString, specialConfigStoreSeparator), " ")
        Else ' walk through string, separating by camel-case switch
            Dim CamelCaseStrLen As Integer = Len(theString)
            Dim i As Integer
            For i = 1 To CamelCaseStrLen
                Dim aChar As String = Mid$(theString, i, 1)
                Dim charAsc As Integer = Asc(aChar)

                If i > 1 Then
                    ' character before current character
                    Dim pre As Integer = Asc(Mid$(theString, i - 1, 1))
                    ' underscore also separates camel-case, except preceded by $, - or another underscore
                    If charAsc = 95 Then
                        If Not (pre = 36 Or pre = 45 Or pre = 95) _
                            Then stringParts += " "
                    End If
                    ' Uppercase characters separate unless they are preceded by other uppercase characters 
                    ' also numbers can precede: And Not (pre >= 48 And pre <= 57) _
                    If (charAsc >= 65 And charAsc <= 90) Then
                        If Not (pre >= 65 And pre <= 90) And Not (pre = 36 Or pre = 45 Or pre = 95) Then stringParts += " "
                    End If
                End If
                stringParts += aChar
            Next
            stringParts = LTrim$(Replace(Replace(stringParts, "   ", " "), "  ", " "))
        End If
    End Function
End Module