Imports ExcelDna.Integration
Imports ExcelDna.Integration.CustomUI
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports Microsoft.Office.Interop

''' <summary>handles all Menu related aspects (context menu for building/refreshing, "DBAddin"/"Load Config" tree menu for retrieving stored configuration files</summary>
<ComVisible(True)>
Public Class MenuHandler
    Inherits ExcelRibbon

    ''' <summary>callback after Excel loaded the Ribbon, used to initialize data for the Ribbon</summary>
    Public Sub ribbonLoaded(theRibbon As ExcelDna.Integration.CustomUI.IRibbonUI)
        DBAddin.theRibbon = theRibbon
        hostApp = ExcelDnaUtil.Application
    End Sub

    ''' <summary>creates the Ribbon (only at startup). any changes to the ribbon can only be done via dynamic menus</summary>
    Public Overrides Function GetCustomUI(RibbonID As String) As String
        Dim customUIXml As String = "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='ribbonLoaded' ><ribbon><tabs><tab id='DBaddinTab' label='DB Addin'>" +
            "<group id='DBAddinGroup' label='General settings'>" +
              "<dropDown id='envDropDown' label='Environment:' sizeString='12345678901234567890' getSelectedItemIndex='GetSelItem' getItemCount='GetItemCount' getItemID='GetItemID' getItemLabel='GetItemLabel' onAction='selectItem'/>" +
              "<dynamicMenu id='DBConfigs' size='normal' label='DB Configs' imageMso='Refresh' screentip='DB Function Configuration Files quick access' getContent='getDBConfigMenu'/>" +
              "<dialogBoxLauncher><button id='dialog' label='About DBAddin' onAction='showAbout' tag='3' screentip='Show Aboutbox and refresh configs if wanted'/></dialogBoxLauncher></group>" +
              "<group id='DBMapperGroup' label='Store Data defined with saveRangeToDB'>"
        For i As Integer = 0 To 10
            customUIXml = customUIXml + "<dynamicMenu id='ID" + i.ToString() + "' " +
                                            "size='normal' getLabel='getSheetLabel' imageMso='SignatureLineInsert' " +
                                            "screentip='Select DBMapper range to store' " +
                                            "getContent='getDynMenContent' getVisible='getDynMenVisible'/>"
        Next
        customUIXml = customUIXml + "</group></tab></tabs></ribbon>" +
         "<contextMenus><contextMenu idMso='ContextMenuCell'>" +
         "<button id='refreshData' label='refresh data (Ctrl-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
         "<button id='gotoDBFunc' label='jump to DBFunc/target (Ctrl-J)' imageMso='ConvertTextToTable' onAction='clickjumpButton' insertBeforeMso='Cut'/>" +
         "<menuSeparator id='MySeparator' insertBeforeMso='Cut'/>" +
         "</contextMenu></contextMenus></customUI>"
        Return customUIXml
    End Function

    ''' <summary>for environment dropdown to get the total number of the entries</summary>
    Public Function GetItemCount(control As IRibbonControl) As Integer
        Return environdefs.Length
    End Function

    ''' <summary>for environment dropdown to get the label of the entries</summary>
    Public Function GetItemLabel(control As IRibbonControl, index As Integer) As String
        Return environdefs(index)
    End Function

    ''' <summary>for environment dropdown to get the ID of the entries</summary>
    Public Function GetItemID(control As IRibbonControl, index As Integer) As String
        Return environdefs(index)
    End Function

    ''' <summary>after selection of environment (using selectItem) used to return the selected environment</summary>
    Public Function GetSelItem(control As IRibbonControl) As Integer
        Return selectedEnvironment
    End Function

    ''' <summary>Choose environment (configured in registry with ConstConnString(N), ConfigStoreFolder(N))</summary>
    Public Sub selectItem(control As IRibbonControl, id As String, index As Integer)
        selectedEnvironment = index
        Dim env As String = (index + 1).ToString()

        If GetSetting("DBAddin", "Settings", "DontChangeEnvironment", String.Empty) = "Y" Then
            MsgBox("Setting DontChangeEnvironment is set to Y, therefore changing the Environment is prevented !")
            Exit Sub
        End If
        storeSetting("ConstConnString", fetchSetting("ConstConnString" & env, String.Empty))
        storeSetting("ConfigStoreFolder", fetchSetting("ConfigStoreFolder" & env, String.Empty))
        storeSetting("ConfigName", fetchSetting("ConfigName" & env, String.Empty))
        initSettings()
        MsgBox("ConstConnString:" & ConstConnString & vbCrLf & "ConfigStoreFolder:" & ConfigStoreFolder & vbCrLf & vbCrLf & "Please refresh DBSheets or DBFuncs to see effects...", vbOKOnly, "set defaults to: ")
        ' provide a chance to reconnect when switching environment...
        theDBFuncEventHandler.cnn = Nothing
        dontTryConnection = False
    End Sub

    ''' <summary>dialogBoxLauncher of leftmost group: activate about box</summary>
    Public Sub showAbout(control As IRibbonControl)
        Dim myAbout As AboutBox1 = New AboutBox1
        myAbout.ShowDialog()
    End Sub

    ''' <summary>store found submenus in this collection</summary>
    Private specialConfigFoldersTempColl As Collection
    ''' <summary>on demand, refresh the DB Config tree</summary>
    Public Sub refreshDBConfigTree(control As IRibbonControl)
        initSettings()
        createConfigTreeMenu()
        MsgBox("refreshed DB Config Tree Menu", vbInformation + vbOKOnly, "DBAddin: refresh Config tree...")
        theRibbon.Invalidate()
    End Sub

    ''' <summary>used to create menu and button ids</summary>
    Private menuID As Integer
    ''' <summary>temporary store menu here</summary>
    Private menuXML As String = vbNullString
    ''' <summary>max depth limitation by Ribbon: 5 levels: 1 top level, 1 folder level (Database foldername) -> 3 left</summary>
    Const specialFolderMaxDepth As Integer = 3
    Private ReadOnly xnspace As XNamespace = "http://schemas.microsoft.com/office/2009/07/customui"

    ''' <summary>get DB Config Menu from File</summary>
    Public Function getDBConfigMenu(control As IRibbonControl) As String
        If menuXML = vbNullString Then createConfigTreeMenu()
        Return menuXML
    End Function

    Private Sub createConfigTreeMenu()
        Dim currentBar, button As XElement

        If Not Directory.Exists(ConfigStoreFolder) Then
            MsgBox("No predefined config store folder '" & ConfigStoreFolder & "' found, please correct setting and refresh!", vbOKOnly + vbCritical, "DBAddin: No config store folder!")
            menuXML = "<menu xmlns='http://schemas.microsoft.com/office/2009/07/customui'><button id='refreshDBConfig' label='refresh DBConfig Tree' imageMso='Refresh' onAction='refreshDBConfigTree'/></menu>"
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
            menuXML = currentBar.ToString()
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

    ''' <summary>load config if config tree menu has been activated (path to config xcl file is in control.Tag)</summary>
    Public Sub getConfig(control As IRibbonControl)
        loadConfig(control.Tag)
    End Sub

    ''' <summary>set the name of the WB/sheet dropdown to the sheet name (for the WB dropdown this is the WB name)</summary>
    Public Function getSheetLabel(control As IRibbonControl) As String
        getSheetLabel = vbNullString
        If DBMapperDefMap.ContainsKey(control.Id) Then getSheetLabel = DBMapperDefMap(control.Id)
    End Function

    ''' <summary>create the buttons in the WB/sheet dropdown</summary>
    Public Function getDynMenContent(control As IRibbonControl) As String
        Dim xmlString As String = "<menu xmlns='http://schemas.microsoft.com/office/2009/07/customui'>"

        If Not DBMapperDefMap.ContainsKey(control.Id) Then Return ""

        Dim currentSheet As String = DBMapperDefMap(control.Id)
        For Each nodeName As String In DBMapperDefColl(currentSheet).Keys
            xmlString = xmlString + "<button id='" + nodeName + "' label='store " + nodeName + "' imageMso='SignatureLineInsert' onAction='saveRangeToDBClick' tag ='" + currentSheet + "' screentip='store " + nodeName + "' supertip='stores data defined in " + nodeName + " Mapper range on sheet " + currentSheet + "' />"
        Next
        xmlString += "</menu>"
        Return xmlString
    End Function

    ''' <summary>shows the sheet button only if it was collected...</summary>
    Public Function getDynMenVisible(control As IRibbonControl) As Boolean
        Return DBMapperDefMap.ContainsKey(control.Id)
    End Function

    Public Shared Sub saveRangeToDBClick(control As IRibbonControl)
        If saveRangeToDB(hostApp.ActiveCell) Then MsgBox("hooray!")
    End Sub

    Public Sub saveAllWSheetRangesToDBClick(control As IRibbonControl)

    End Sub

    Public Sub saveAllWBookRangesToDBClick(control As IRibbonControl)

    End Sub

    ''' <summary>context menu entry refreshData: refresh Data in db function (if area or cell selected) or all db functions</summary>
    Public Sub clickrefreshData(control As IRibbonControl)
        refreshData()
    End Sub

    ''' <summary>context menu entry gotoDBFunc: jumps from DB function to data area and back</summary>
    Public Sub clickjumpButton(control As IRibbonControl)
        jumpButton()
    End Sub

End Class
