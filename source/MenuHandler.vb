Imports Microsoft.Office.Interop
Imports ExcelDna.Integration
Imports ExcelDna.Integration.CustomUI
Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel

Public Module MenuCommands

    Public defnames As String() = {}
    Public defsheetMap As Dictionary(Of String, String)
    Public defsheetColl As Dictionary(Of String, Dictionary(Of String, Range))

    <ExcelCommand(Name:="saveRangeToDB", ShortCut:="^S")>
    Public Sub saveRangeToDBClick()
        If saveRangeToDB(theHostApp.ActiveCell) Then MsgBox("hooray!")
    End Sub

    Public Sub saveAllWSheetRangesToDBClick()

    End Sub

    Public Sub saveAllWBookRangesToDBClick()

    End Sub

    '  context menu "refresh data" was clicked, do a refresh of db functions (shortcut CTRL-e)
    <ExcelCommand(Name:="refreshData", ShortCut:="^e")>
    Public Sub refreshData()
        initSettings()

        ' enable events in case there were some problems in procedure with EnableEvents = false
        On Error Resume Next
        theHostApp.EnableEvents = True
        If Err.Number <> 0 Then
            LogError("Can't refresh data while lookup dropdown is open !!")
            Exit Sub
        End If

        ' also reset the database connection in case of errors...
        theDBFuncEventHandler.cnn.Close()
        theDBFuncEventHandler.cnn = Nothing

        dontTryConnection = False
        On Error GoTo err1

        ' now for DBListfetch/DBRowfetch resetting
        allCalcContainers = Nothing
        Dim underlyingName As Excel.Name
        underlyingName = getDBRangeName(theHostApp.ActiveCell)
        theHostApp.ScreenUpdating = True
        If underlyingName Is Nothing Then
            ' reset query cache, so we really get new data !
            theDBFuncEventHandler.queryCache = New Collection
            refreshDBFunctions(theHostApp.ActiveWorkbook)
            ' general refresh: also refresh all embedded queries and pivot tables..
            'On Error Resume Next
            'Dim ws     As Excel.Worksheet
            'Dim qrytbl As Excel.QueryTable
            'Dim pivottbl As Excel.PivotTable

            'For Each ws In theHostApp.ActiveWorkbook.Worksheets
            '    For Each qrytbl In ws.QueryTables
            '       qrytbl.Refresh
            '    Next
            '    For Each pivottbl In ws.PivotTables
            '        pivottbl.PivotCache.Refresh
            '    Next
            'Next
            'On Error GoTo err1
        Else
            ' reset query cache, so we really get new data !
            theDBFuncEventHandler.queryCache = New Collection

            Dim jumpName As String
            jumpName = underlyingName.Name
            ' because of a stupid excel behaviour (Range.Dirty only works if the parent sheet of Range is active)
            ' we have to jump to the sheet containing the dbfunction and then activate back...
            theDBFuncEventHandler.origWS = Nothing
            ' this is switched back in DBFuncEventHandler.Calculate event,
            ' where we also select back the original active worksheet

            ' we're being called on a target (addtional) functions area
            If Left$(jumpName, 10) = "DBFtargetF" Then
                jumpName = Replace(jumpName, "DBFtargetF", "DBFsource", 1, , vbTextCompare)

                If Not theHostApp.Range(jumpName).Parent Is theHostApp.ActiveSheet Then
                    theHostApp.ScreenUpdating = False
                    theDBFuncEventHandler.origWS = theHostApp.ActiveSheet
                    On Error Resume Next
                    theHostApp.Range(jumpName).Parent.Select
                    On Error GoTo err1
                End If
                theHostApp.Range(jumpName).Dirty
                ' we're being called on a target area
            ElseIf Left$(jumpName, 9) = "DBFtarget" Then
                jumpName = Replace(jumpName, "DBFtarget", "DBFsource", 1, , vbTextCompare)

                If Not theHostApp.Range(jumpName).Parent Is theHostApp.ActiveSheet Then
                    theHostApp.ScreenUpdating = False
                    theDBFuncEventHandler.origWS = theHostApp.ActiveSheet
                    On Error Resume Next
                    theHostApp.Range(jumpName).Parent.Select
                    On Error GoTo err1
                End If
                theHostApp.Range(jumpName).Dirty
                ' we're being called on a source (invoking function) cell
            ElseIf Left$(jumpName, 9) = "DBFsource" Then
                On Error Resume Next
                theHostApp.Range(jumpName).Dirty
                On Error GoTo err1
            Else
                refreshDBFunctions(theHostApp.ActiveWorkbook)
            End If
        End If

        Exit Sub
err1:
        LogToEventViewer("Error (" & Err.Description & ") in MenuHandler.refreshData in " & Erl(), EventLogEntryType.Error)
    End Sub

    ''
    '  jumps from DB function to data area and back
    Public Sub jumpButton()
        Dim underlyingName As Excel.Name
        underlyingName = getDBRangeName(theHostApp.ActiveCell)

        If underlyingName Is Nothing Then Exit Sub
        Dim jumpName As String
        jumpName = underlyingName.Name
        If Left$(jumpName, 10) = "DBFtargetF" Then
            jumpName = Replace(jumpName, "DBFtargetF", "DBFsource", 1, , vbTextCompare)
        ElseIf Left$(jumpName, 9) = "DBFtarget" Then
            jumpName = Replace(jumpName, "DBFtarget", "DBFsource", 1, , vbTextCompare)
        Else
            jumpName = Replace(jumpName, "DBFsource", "DBFtarget", 1, , vbTextCompare)
        End If
        On Error Resume Next
        theHostApp.Range(jumpName).Parent.Select
        theHostApp.Range(jumpName).Select
        If Err.Number <> 0 Then LogWarn("Can't jump to target/source, corresponding workbook open? " & Err.Description, 1)
        Err.Clear()
    End Sub

    ' gets defined named ranges for DBMapper invocation in the current workbook 
    Public Function getRNames() As String
        ReDim Preserve defnames(-1)
        defsheetColl = New Dictionary(Of String, Dictionary(Of String, Range))
        defsheetMap = New Dictionary(Of String, String)
        Dim i As Integer = 0
        For Each namedrange As Name In theHostApp.ActiveWorkbook.Names
            Dim cleanname As String = Replace(namedrange.Name, namedrange.Parent.Name & "!", "")
            If Left(cleanname, 8) = "DBMapper" Then
                If InStr(namedrange.RefersTo, "#REF!") > 0 Then Return "DBMapper definitions range " + namedrange.Parent.name + "!" + namedrange.Name + " contains #REF!"
                ' final name of entry is without DBMapper and !
                Dim finalname As String = Replace(Replace(namedrange.Name, "DBMapper", ""), "!", "")
                Dim nodeName As String = Replace(Replace(namedrange.Name, "DBMapper", ""), namedrange.Parent.Name & "!", "")
                If nodeName = "" Then nodeName = "MainDBMapper"
                If Not InStr(namedrange.Name, "!") > 0 Then
                    finalname = theHostApp.ActiveWorkbook.Name + finalname
                End If

                Dim defColl As Dictionary(Of String, Range)
                If Not defsheetColl.ContainsKey(namedrange.Parent.Name) Then
                    ' add to new sheet "menu"
                    defColl = New Dictionary(Of String, Range)
                    defColl.Add(nodeName, namedrange.RefersToRange)
                    defsheetColl.Add(namedrange.Parent.Name, defColl)
                    defsheetMap.Add("ID" + i.ToString(), namedrange.Parent.Name)
                    i = i + 1
                Else
                    ' add definition to existing sheet "menu"
                    defColl = defsheetColl(namedrange.Parent.Name)
                    defColl.Add(nodeName, namedrange.RefersToRange)
                End If
            End If
        Next
        Return vbNullString
    End Function

    Private specialConfigFoldersTempColl As Collection

    ''
    ' create the config tree menu
    Public Sub createConfigTreeMenu(Optional fillConfigTreeFirstRun As Boolean = False)
        Dim DBConfigB As CommandBarPopup
        Dim cbar As CommandBar

        On Error GoTo err1
        cbar = theHostApp.CommandBars("DBAddin")
        DBConfigB = theHostApp.CommandBars.FindControl(Tag:=gsDBConfigB_TAG)
        If DBConfigB Is Nothing Then
            fillConfigTreeFirstRun = True
            DBConfigB = cbar.Controls.Add(controlType:=MsoControlType.msoControlPopup, Id:=2, Before:=1, Parameter:=1, Temporary:=False)
            DBConfigB.Caption = "DBConfigs"
            DBConfigB.Tag = gsDBConfigB_TAG
        End If
        DBConfigB.TooltipText = "DB Function Configuration Files quick access"
        If fillConfigTreeFirstRun Then
            If Not File.Exists(ConfigStoreFolder) Then
                DBConfigB.Caption = "No Config Store !!"
                DBConfigB.TooltipText = "Couldn't find predefined config store folder '" & ConfigStoreFolder & "', please check registry setting for config store folder location and refresh !"
                Exit Sub
            End If
            specialConfigFoldersTempColl = New Collection
            readAllFiles(ConfigStoreFolder, DBConfigB)
            specialConfigFoldersTempColl = Nothing
            theHostApp.StatusBar = String.Empty
        End If
        Exit Sub

err1:
        LogToEventViewer("Error (" & Err.Description & ") in MenuHandler.createTreeMenu in " & Erl(), EventLogEntryType.Error)
    End Sub

    ''
    ' reads all files contained in rootPath and its subfolders (recursively)
    '             and adds them to currentBar and submenus (recursively)
    ' @param rootPath
    ' @param currentBar
    Sub readAllFiles(rootPath As String, currentBar As CommandBarControl)
        Dim newBar As CommandBarControl
        Dim configType As String, entry As String
        Dim i As Long
        Dim DirList() As String, fileList() As String
        Dim specialFolderMaxDepth As Integer
        Dim specialConfigStoreSeparator As String

        On Error GoTo err1
        configType = "XCL"

        ' read all leaf node entries (files) to create action menus
        entry = Dir(rootPath & "\*." & configType, vbNormal)
        i = 0 : ReDim fileList(i)
        Do While entry.Length > 0
            ReDim Preserve fileList(i)
            fileList(i) = entry
            i = i + 1
            entry = Dir()
        Loop

        If i > 0 Then
            If sortConfigStoreFolders Then QuickSort(fileList, LBound(fileList), UBound(fileList))

            ' for special folders split further into camelcase (or other) separated names
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
                Dim firstCharLevel As Boolean
                firstCharLevel = CBool(fetchSetting(spclFolder & "FirstLetterLevel", "False"))
                specialFolderMaxDepth = fetchSetting(spclFolder & "MaxDepth", 10000)
                specialConfigStoreSeparator = fetchSetting(spclFolder & "Separator", String.Empty)
                For i = 0 To UBound(fileList)
                    ' is current entry contained in next entry then revert order to allow for containment in next entry's hierarchy..
                    If i < UBound(fileList) Then
                        If InStr(1, Left$(fileList(i + 1), Len(fileList(i + 1)) - 4), Left$(fileList(i), Len(fileList(i)) - 4)) > 0 Then
                            buildFileSepMenuCtrl(stringParts(IIf(firstCharLevel, Left$(fileList(i + 1), 1) & " ", String.Empty) &
                                                            Left$(fileList(i + 1), Len(fileList(i + 1)) - 4), specialConfigStoreSeparator), currentBar, rootPath & "\" & fileList(i + 1), spclFolder, specialFolderMaxDepth)
                            buildFileSepMenuCtrl(stringParts(IIf(firstCharLevel, Left$(fileList(i), 1) & " ", String.Empty) &
                                                            Left$(fileList(i), Len(fileList(i)) - 4), specialConfigStoreSeparator), currentBar, rootPath & "\" & fileList(i), spclFolder, specialFolderMaxDepth)
                            i = i + 2
                            If i > UBound(fileList) Then Exit For
                        End If
                    End If
                    buildFileSepMenuCtrl(stringParts(IIf(firstCharLevel, Left$(fileList(i), 1) & " ", String.Empty) &
                            Left$(fileList(i), Len(fileList(i)) - 4), specialConfigStoreSeparator), currentBar, rootPath & "\" & fileList(i), spclFolder, specialFolderMaxDepth)
                Next
                ' normal case: just follow the path and enter all
            Else
                For i = 0 To UBound(fileList)
                    'newBar = currentBar.Controls.Add(Type:=msoControlButton)
                    'newBar.Caption = Left$(fileList(i), Len(fileList(i)) - 4)
                    'newBar.Parameter = rootPath & "\" & fileList(i)
                    'newBar.Tag = gsITEMLOADPREPARED_TAG
                Next
            End If
        End If

        ' read all dir entries
        entry = Dir(rootPath & "\", vbDirectory)
        i = 0 : ReDim DirList(i)
        Do While entry.Length > 0
            If entry <> "." And entry <> ".." And (GetAttr(rootPath & "\" & entry) And vbDirectory) = vbDirectory Then
                ReDim Preserve DirList(i)
                DirList(i) = entry
                i = i + 1
            End If
            entry = Dir()
        Loop
        If i = 0 Then Exit Sub
        If sortConfigStoreFolders Then QuickSort(DirList, LBound(DirList), UBound(DirList))

        ' recursively build branched menu structure from dirEntries
        For i = 0 To UBound(DirList)
            theHostApp.StatusBar = "Filling DBConfigs Menu: " & rootPath & "\" & DirList(i)
            'newBar = currentBar.Controls.Add(Type:=msoControlPopup)
            'newBar.Caption = DirList(i)
            'newBar.Tag = DirList(i)
            'readAllFiles(rootPath & "\" & DirList(i), newBar)
        Next
        Exit Sub
err1:
        LogToEventViewer("Error (" & Err.Description & ") in MenuHandler.readAllFiles in " & Erl(), EventLogEntryType.Error)
        Resume Next
    End Sub

    'TODO: convert to ribbon
    ''
    ' parses Substrings contained in nameParts (recursively)
    '             and adds them to currentBar and submenus (recursively)
    ' @param nameParts
    ' @param currentBar
    ' @param fullPathName
    ' @param newRootName
    ' @param specialFolderMaxDepth
    Sub buildFileSepMenuCtrl(nameParts As String, currentBar As CommandBarControl, fullPathName As String, newRootName As String, specialFolderMaxDepth As Integer)
        Dim newBar As CommandBarControl
        Dim newSubBar As CommandBarControl
        Static currentDepth As Integer

        On Error GoTo buildFileSepMenuCtrl_Err
        ' end node: add callable cmdbar entry
        If InStr(1, nameParts, " ") = 0 Or currentDepth > specialFolderMaxDepth - 1 Then
            On Error Resume Next
            newBar = specialConfigFoldersTempColl(newRootName & nameParts)
            If Err.Number <> 0 Then newSubBar = currentBar
            Err.Clear()
            'newBar = newSubBar.Controls.Add(Type:=msoControlButton)
            Dim entryName : entryName = Mid$(fullPathName, InStrRev(fullPathName, "\") + 1)
            newBar.Caption = Left$(entryName, Len(entryName) - 4)
            'newBar.Parameter = fullPathName
            newBar.Tag = gsITEMLOADPREPARED_TAG
        Else  ' leaf node: add popup menu entry
            Dim newName As String
            newName = Left$(nameParts, InStr(1, nameParts, " ") - 1)
            On Error Resume Next
            newBar = specialConfigFoldersTempColl(newRootName & newName)
            If Err.Number <> 0 Then
                'newBar = currentBar.Controls.Add(Type:=msoControlPopup)
                newBar.Caption = newName
                specialConfigFoldersTempColl.Add(newBar, newRootName & newName)
            End If
            Err.Clear()
            currentDepth = currentDepth + 1
            buildFileSepMenuCtrl(Mid$(nameParts, InStr(1, nameParts, " ") + 1), newBar, fullPathName, newRootName & newName, specialFolderMaxDepth)
            currentDepth = currentDepth - 1
        End If
        Exit Sub

buildFileSepMenuCtrl_Err:
        LogToEventViewer("Error (" & Err.Description & ") in MenuHandler.buildFileSepMenuCtrl in " & Erl(), EventLogEntryType.Error)
    End Sub

    ''
    ' return parts of a CamelCase string
    Private Function stringParts(theString As String, specialConfigStoreSeparator As String) As String
        Dim CamelCaseStrLen As Integer
        Dim i As Integer
        Dim aChar As String
        Dim charAsc As Integer
        Dim pre As Integer

        stringParts = String.Empty
        If specialConfigStoreSeparator.Length > 0 Then
            stringParts = Join(Split(theString, specialConfigStoreSeparator), " ")
        Else
            CamelCaseStrLen = Len(theString)
            For i = 1 To CamelCaseStrLen
                aChar = Mid$(theString, i, 1)
                charAsc = Asc(aChar)

                If i > 1 Then
                    pre = Asc(Mid$(theString, i - 1, 1))
                    If charAsc = 95 Then
                        If Not (pre = 36 Or pre = 45 Or pre = 95) _
                            Then stringParts = stringParts & " "
                    End If
                    If (charAsc >= 65 And charAsc <= 90) Then      'Uppercase characters
                        If Not (pre >= 65 And pre <= 90) _
                           And Not (pre = 36 Or pre = 45 Or pre = 95) _
                           And Not (pre >= 48 And pre <= 57) _
                           Then stringParts = stringParts & " "
                    End If
                End If
                stringParts = stringParts & aChar
            Next
            stringParts = LTrim$(Replace(Replace(stringParts, "   ", " "), "  ", " "))
        End If
    End Function

    ''
    ' Do string Quicksort of array sortList
    Private Sub QuickSort(ByRef sortList As Object, ByVal LB As Long, ByVal UB As Long)
        Dim P1 As Long, P2 As Long, Ref As String, temp As String

        P1 = LB
        P2 = UB
        Ref = sortList((P1 + P2) / 2)

        Do
            Do While (sortList(P1) < Ref)
                P1 = P1 + 1
            Loop

            Do While (sortList(P2) > Ref)
                P2 = P2 - 1
            Loop

            If P1 <= P2 Then
                temp = sortList(P1)
                sortList(P1) = sortList(P2)
                sortList(P2) = temp

                P1 = P1 + 1
                P2 = P2 - 1
            End If
        Loop Until (P1 > P2)

        If LB < P2 Then Call QuickSort(sortList, LB, P2)
        If P1 < UB Then Call QuickSort(sortList, P1, UB)
    End Sub

End Module

''
'  handles all Menu related aspects (context menu for building/refreshing,
'             "DBAddin"/"Load Config" tree menu for retrieving stored configuration files
<ComVisible(True)>
Public Class MenuHandler
    Inherits ExcelRibbon

    Private specialConfigFoldersTempColl As Collection
    Private selectedEnvironment As Integer

    Public Sub ribbonLoaded(theRibbon As ExcelDna.Integration.CustomUI.IRibbonUI)
        Globals.theRibbon = theRibbon
        ' load environments
        Dim i As Integer = 1
        Dim ConfigName As String
        Do
            ConfigName = fetchSetting("ConfigName" + i, vbNullString)
            If Len(ConfigName) > 0 Then
                ReDim Preserve defnames(defnames.Length)
                defnames(defnames.Length - 1) = ConfigName + " - " + i.ToString()
            End If
            ' set selectedEnvironment
            If fetchSetting("ConstConnString" & i, vbNullString) = ConstConnString Then
                selectedEnvironment = i
                storeSetting("ConfigName", ConfigName)
            End If
            i = i + 1
        Loop Until Len(ConfigName) = 0
    End Sub

    Public Sub showAbout(control As IRibbonControl)
        Dim myAbout As AboutBox1 = New AboutBox1
        myAbout.ShowDialog()
    End Sub

    Public Function GetItemCount(control As IRibbonControl) As Integer
        Return defnames.Length
    End Function

    Public Function GetItemLabel(control As IRibbonControl, index As Integer) As String
        Return defnames(index)
    End Function

    Public Function GetItemID(control As IRibbonControl, index As Integer) As String
        Return defnames(index)
    End Function

    Public Function GetSelItem(control As IRibbonControl) As Integer
        Return selectedEnvironment
    End Function
    ' Choose environment (configured in registry with ConfigName<N>, ConstConnString<N>, ConfigStoreFolder<N>)
    Public Sub selectItem(control As IRibbonControl, id As String, index As Integer)
        selectedEnvironment = index + 1

        If GetSetting("DBAddin", "Settings", "DontChangeEnvironment", String.Empty) = "Y" Then
            MsgBox("Setting DontChangeEnvironment is set to Y, therefore changing the Environment is prevented !")
            Exit Sub
        End If
        storeSetting("ConstConnString", fetchSetting("ConstConnString" & selectedEnvironment.ToString(), String.Empty))
        storeSetting("ConfigStoreFolder", fetchSetting("ConfigStoreFolder" & selectedEnvironment.ToString(), String.Empty))
        storeSetting("ConfigName", fetchSetting("ConfigName" & selectedEnvironment.ToString(), String.Empty))

        initSettings()
        dontTryConnection = False  ' provide a chance to reconnect when switching environment...
    End Sub


    ' creates the Ribbon <buttonGroup id='buttonGroup'> <box id='box2' boxStyle='horizontal'>
    Public Overrides Function GetCustomUI(RibbonID As String) As String
        Dim customUIXml As String = "<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad='ribbonLoaded' ><ribbon><tabs><tab id='RaddinTab' label='R Addin'>" +
            "<group id='DBaddinGroup' label='General settings'>" +
              "<dropDown id='envDropDown' label='Environment:' sizeString='12345678901234567890' getSelectedItemIndex='GetSelItem' getItemCount='GetItemCount' getItemID='GetItemID' getItemLabel='GetItemLabel' onAction='selectItem'/>" +
              "" +
              "<dialogBoxLauncher><button id='dialog' label='About DBAddin' onAction='showAbout' tag='3' screentip='Show Aboutbox and refresh configs if wanted'/></dialogBoxLauncher></group>" +
              "<group id='RscriptsGroup' label='Store Data defined with saveRangeToDB'>"
        For i As Integer = 0 To 3
            customUIXml = customUIXml + "<dynamicMenu id='ID" + i.ToString() + "' " +
                                            "size='normal' getLabel='getSheetLabel' imageMso='SignatureLineInsert' " +
                                            "screentip='Select script to run' " +
                                            "getContent='getDynMenContent' getVisible='getDynMenVisible'/>"
        Next
        customUIXml = customUIXml + "</group></tab></tabs></ribbon></customUI>"
        Return customUIXml
    End Function

    ' set the name of the WB/sheet dropdown to the sheet name (for the WB dropdown this is the WB name) 
    Public Function getSheetLabel(control As IRibbonControl) As String
        getSheetLabel = vbNullString
        If defsheetMap.ContainsKey(control.Id) Then getSheetLabel = defsheetMap(control.Id)
    End Function

    ' create the buttons in the WB/sheet dropdown
    Public Function getDynMenContent(control As IRibbonControl) As String
        Dim xmlString As String = "<menu xmlns='http://schemas.microsoft.com/office/2009/07/customui'>"
        Dim currentSheet As String = defsheetMap(control.Id)
        For Each nodeName As String In defsheetColl(currentSheet).Keys
            xmlString = xmlString + "<button id='" + nodeName + "' label='run " + nodeName + "' imageMso='SignatureLineInsert' onAction='startRprocess' tag ='" + currentSheet + "' screentip='store " + nodeName + "' supertip='stores data defined in " + nodeName + " Mapper range on sheet " + currentSheet + "' />"
        Next
        xmlString = xmlString + "</menu>"
        Return xmlString
    End Function

    ' shows the sheet button only if it was collected...
    Public Function getDynMenVisible(control As IRibbonControl) As Boolean
        Return defsheetMap.ContainsKey(control.Id)
    End Function


    'TODO: convert to ribbon
    ''
    ' load config if config tree menu has been activated (name stored in Ctrl.Parameter)
    ' @param Ctrl
    ' @param CancelDefault
    Private Sub mDBConfigPreparedButton_Click(ByVal Ctrl As CommandBarButton, CancelDefault As Boolean)
        loadConfig(Ctrl.Tag)
    End Sub


    'TODO: convert to ribbon
    ''
    '  sets up the "DBAddin" Cmd menu bar
    Private Sub createCommandBar()
        '        Dim cbar As CommandBar
        '        Dim newBtn As CommandBarButton
        '        Dim cbpop As CommandBarControl

        '        On Error GoTo err1
        '        disableBar = False
        '        If existsCommandBar() Then
        '            cbar = theHostApp.CommandBars("DBAddin")
        '        Else
        '            cbar = theHostApp.CommandBars.Add(name:="DBAddin")

        '            With cbar
        '                .Visible = True
        '            End With
        '        End If
        '        If theHostApp.CommandBars.FindControl(Tag:=gsABOUT_TAG) Is Nothing Then
        '            ' Create "About" control button on the main menu bar
        '            newBtn = cbar.Controls.Add(controlType:=MsoControlType.msoControlButton,1,)
        '            newBtn.Caption = "About"
        '            newBtn.Tag = gsABOUT_TAG
        '            newBtn.Style = MsoButtonStyle.msoButtonIconAndCaption
        '            newBtn.FaceId = 984
        '        End If
        '        theHostApp.CommandBars.FindControl(Tag:=gsABOUT_TAG).ToolTipText = "General DBAddin Information/Help"

        '        ' Create "Connections" control popup on the main menu bar
        '        Dim connBtn As CommandBarPopup
        '        On Error Resume Next
        '        theHostApp.CommandBars.FindControl(Tag:=gsCONSTCONN_TAG).Delete
        '        On Error GoTo err1
        '        connBtn = cbar.Controls.Add(Type:=msoControlPopup)
        '        connBtn.Tag = gsCONSTCONN_TAG
        '        connBtn.ToolTipText = "Select Connection Definitions"
        '        Dim i As Long, ConfigName As String
        '        i = 1
        '        Do
        '            ConfigName = fetchSetting("ConfigName" & i, String.Empty)
        '            If ConfigName.Length > 0 Then
        '                With connBtn.Controls.Add(controlType:=MsoControlType.msoControlButton)
        '                    .Caption = ConfigName & " - " & i
        '                    .Style = msoButtonCaption
        '                    .Parameter = i
        '                    .Tag = gsCONSTCONNACTION_TAG

        '                    If fetchSetting("ConstConnString" & i, String.Empty) = ConstConnString Then
        '                        .State = -1
        '                        connBtn.Caption = "Env: " & ConfigName
        '                        storeSetting("ConfigName", ConfigName)
        '                    Else
        '                        .State = 0
        '                    End If
        '                End With
        '            End If
        '            i = i + 1
        '        Loop Until ConfigName.Length = 0

        '        mEvironSelButton = theHostApp.CommandBars.FindControl(Tag:=gsCONSTCONNACTION_TAG)
        '        theHostApp.CommandBars.FindControl(Tag:=gsDBSheetParametersB_TAG).Enabled = False
        '        createConfigTreeMenu()
        '        Exit Sub

        'err1:
        '        LogToEventViewer("Error (" & Err.Description & ") in MenuHandler.createCommandBar in " & Erl(), EventLogEntryType.Error, 1)
    End Sub


End Class
