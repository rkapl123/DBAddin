Imports Microsoft.Office.Interop
Imports ExcelDna.Integration
Imports ExcelDna.Integration.CustomUI
Imports System.IO
Imports System.Runtime.InteropServices

''
'  handles all Menu related aspects (context menu for building/refreshing,
'             "DBAddin"/"Load Config" tree menu for retrieving stored configuration files
<ComVisible(True)>
Public Class MenuHandler
    Inherits ExcelRibbon

    Private specialConfigFoldersTempColl As Collection

    Public Sub ribbonLoaded(myribbon As ExcelDna.Integration.CustomUI.IRibbonUI)
        Globals.theRibbon = myribbon
    End Sub

    '  context menu "refresh data" was clicked, do a refresh of db functions (shortcut CTRL-e)
    <ExcelCommand(ShortCut:="^e")>
    Public Sub refreshData(control As IRibbonControl)
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
    Public Sub jumpButton(control As IRibbonControl)
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

    'TODO: convert to ribbon
    ''
    ' Choose environment (configured in registry with ConfigName<N>, ConstConnString<N>, ConfigStoreFolder<N>)
    ' @param barEnabled whether bar parts should be enabled or not
    Private Sub mEvironSelButton_Click(ByVal Ctrl As CommandBarButton, CancelDefault As Boolean)
        Dim constConnSels, constConnSel
        Dim env As String = "1"

        If GetSetting("DBAddin", "Settings", "DontChangeEnvironment", String.Empty) = "Y" Then
            MsgBox("Setting DontChangeEnvironment is set to Y, therefore changing the Environment is prevented !")
            Exit Sub
        End If
        storeSetting("ConstConnString", fetchSetting("ConstConnString" & env, String.Empty))
        storeSetting("ConfigStoreFolder", fetchSetting("ConfigStoreFolder" & env, String.Empty))
        storeSetting("ConfigName", fetchSetting("ConfigName" & env, String.Empty))

        constConnSels = theHostApp.CommandBars.FindControl(Tag:=gsCONSTCONN_TAG).Controls
        For Each constConnSel In constConnSels
            constConnSel.State = 0
        Next

        initSettings()
        dontTryConnection = False  ' provide a chance to reconnect when switching environment...
        MsgBox("ConstConnString:" & ConstConnString & vbCrLf & "ConfigStoreFolder:" & ConfigStoreFolder & vbCrLf & vbCrLf & "Please refresh DBSheets or DBFuncs to see effects...", vbOKOnly, "set defaults to: ")
        theHostApp.CommandBars.FindControl(Tag:=gsCONSTCONN_TAG).caption = "Env: " & fetchSetting("ConfigName" & env, String.Empty)
    End Sub

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
    ' refreshes the DBConfig tree menu (automatic refreshing on open or context menu refresh turned out to be too lengthy)
    ' @param Ctrl
    ' @param CancelDefault
    Private Sub mDBConfigRefreshButton_Click(ByVal Ctrl As CommandBarButton, CancelDefault As Boolean)
        initSettings()
        createConfigTreeMenu(True)
        MsgBox("refreshed ConfigTreeMenu and restarted theDBSheetAppHandler", vbInformation + vbOKOnly, "DBAddin: refresh Config tree...")
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
        '        If VBDEBUG Then Debug.Print("Error (" & Err.Description & ") in MenuHandler.createCommandBar") : Stop : Resume
        '        LogToEventViewer("Error (" & Err.Description & ") in MenuHandler.createCommandBar in " & Erl(), EventLogEntryType.Error, 1)
    End Sub

    'TODO: convert to ribbon
    ''
    ' create the config tree menu
    ' @param customContext whether we should create a separate Word customization context or not (if invoked separately)
    Private Sub createConfigTreeMenu(Optional fillConfigTreeFirstRun As Boolean = False)
        '        Dim DBConfigB As CommandBarPopup
        '        Dim cbar As CommandBar

        '        On Error GoTo err1
        '        cbar = theHostApp.CommandBars("DBAddin")
        '        DBConfigB = theHostApp.CommandBars.FindControl(Tag:=gsDBConfigB_TAG)
        '        If DBConfigB Is Nothing Then
        '            fillConfigTreeFirstRun = True
        '            DBConfigB = cbar.Controls.Add(controlType:=MsoControlType.msoControlPopup, Id:=2, Before:=1, Parameter:=1, Temporary:=False)
        '            DBConfigB.caption = "DBConfigs"
        '            DBConfigB.Tag = gsDBConfigB_TAG
        '        End If
        '        DBConfigB.ToolTipText = "DB Function Configuration Files quick access"
        '        If fillConfigTreeFirstRun Then
        '            Dim btn As Object
        '            For Each btn In DBConfigB.Controls
        '                On Error Resume Next
        '                btn.Delete
        '                On Error GoTo err1
        '            Next
        '            On Error Resume Next
        '            If Not File.Exists(ConfigStoreFolder) Then
        '                DBConfigB.caption = "No Config Store !!"
        '                DBConfigB.ToolTipText = "Couldn't find predefined config store folder '" & ConfigStoreFolder & "', please check registry setting for config store folder location and refresh !"
        '                disableBar = True
        '                Exit Sub
        '            End If
        '            On Error GoTo err1
        '            specialConfigFoldersTempColl = New Collection
        '            If Not disableBar Then readAllFiles(ConfigStoreFolder, DBConfigB)
        '            specialConfigFoldersTempColl = Nothing
        '            theHostApp.StatusBar = String.Empty
        '        End If
        '        mDBConfigPreparedButton = theHostApp.CommandBars.FindControl(Tag:=gsITEMLOADPREPARED_TAG)
        '        mDBConfigRefreshButton = theHostApp.CommandBars.FindControl(Tag:=gsDBConfigRefreshB_TAG)
        '        If mDBConfigRefreshButton Is Nothing Then
        '            mDBConfigRefreshButton = DBConfigB.Controls.Add(Type:=msoControlButton, Before:=1)
        '            mDBConfigRefreshButton.Caption = "refresh DBConfigs"
        '            mDBConfigRefreshButton.Tag = gsDBConfigRefreshB_TAG
        '            mDBConfigRefreshButton.FaceId = 1020
        '            mDBConfigRefreshButton.TooltipText = "refresh DB Function Configuration Files Tree"
        '        End If
        '        Exit Sub

        'err1:
        '        If VBDEBUG Then Debug.Print("Error (" & Err.Description & ") in MenuHandler.createTreeMenu") : Stop : Resume
        '        LogToEventViewer("Error (" & Err.Description & ") in MenuHandler.createTreeMenu in " & Erl(), EventLogEntryType.Error, 1)
    End Sub

    'TODO: convert to ribbon
    ''
    ' reads all files contained in rootPath and its subfolders (recursively)
    '             and adds them to currentBar and submenus (recursively)
    ' @param rootPath
    ' @param currentBar
    Sub readAllFiles(rootPath As String, currentBar As CommandBarControl)
        '        Dim newBar As CommandBarControl
        '        Dim configType As String, entry As String
        '        Dim i As Long
        '        Dim DirList() As String, fileList() As String
        '        Dim specialFolderMaxDepth As Integer
        '        Dim specialConfigStoreSeparator As String

        '        On Error GoTo err1
        '        configType = "XCL"

        '        ' read all leaf node entries (files) to create action menus
        '        entry = Dir(rootPath & "\*." & configType, vbNormal)
        '        i = 0 : ReDim fileList(i)
        '        Do While entry.Length > 0
        '            ReDim Preserve fileList(i)
        '            fileList(i) = entry
        '            i = i + 1
        '            entry = Dir()
        '        Loop

        '        If i > 0 Then
        '            If sortConfigStoreFolders Then QuickSort(fileList, LBound(fileList), UBound(fileList))

        '            ' for special folders split further into camelcase (or other) separated names
        '            Dim aFolder : Dim spclFolder As String : spclFolder = String.Empty
        '            Dim theFolder As String
        '            theFolder = Mid$(rootPath, InStrRev(rootPath, "\") + 1)
        '            For Each aFolder In specialConfigStoreFolders
        '                If UCase$(theFolder) = UCase$(aFolder) Then
        '                    spclFolder = aFolder
        '                    Exit For
        '                End If
        '            Next

        '            If spclFolder.Length > 0 Then
        '                Dim firstCharLevel As Boolean
        '                firstCharLevel = CBool(fetchSetting(spclFolder & "FirstLetterLevel", "False"))
        '                specialFolderMaxDepth = fetchSetting(spclFolder & "MaxDepth", 10000)
        '                specialConfigStoreSeparator = fetchSetting(spclFolder & "Separator", String.Empty)
        '                For i = 0 To UBound(fileList)
        '                    ' is current entry contained in next entry then revert order to allow for containment in next entry's hierarchy..
        '                    If i < UBound(fileList) Then
        '                        If InStr(1, Left$(fileList(i + 1), Len(fileList(i + 1)) - 4), Left$(fileList(i), Len(fileList(i)) - 4)) > 0 Then
        '                            buildFileSepMenuCtrl(stringParts(IIf(firstCharLevel, Left$(fileList(i + 1), 1) & " ", String.Empty) &
        '                                                            Left$(fileList(i + 1), Len(fileList(i + 1)) - 4), specialConfigStoreSeparator), currentBar, rootPath & "\" & fileList(i + 1), spclFolder, specialFolderMaxDepth)
        '                            buildFileSepMenuCtrl(stringParts(IIf(firstCharLevel, Left$(fileList(i), 1) & " ", String.Empty) &
        '                                                            Left$(fileList(i), Len(fileList(i)) - 4), specialConfigStoreSeparator), currentBar, rootPath & "\" & fileList(i), spclFolder, specialFolderMaxDepth)
        '                            i = i + 2
        '                            If i > UBound(fileList) Then Exit For
        '                        End If
        '                    End If
        '                    buildFileSepMenuCtrl(stringParts(IIf(firstCharLevel, Left$(fileList(i), 1) & " ", String.Empty) &
        '                            Left$(fileList(i), Len(fileList(i)) - 4), specialConfigStoreSeparator), currentBar, rootPath & "\" & fileList(i), spclFolder, specialFolderMaxDepth)
        '                Next
        '                ' normal case: just follow the path and enter all
        '            Else
        '                For i = 0 To UBound(fileList)
        '                    newBar = currentBar.Controls.Add(Type:=msoControlButton)
        '                    newBar.caption = Left$(fileList(i), Len(fileList(i)) - 4)
        '                    newBar.Parameter = rootPath & "\" & fileList(i)
        '                    newBar.Tag = gsITEMLOADPREPARED_TAG
        '                Next
        '            End If
        '        End If

        '        ' read all dir entries
        '        entry = Dir(rootPath & "\", vbDirectory)
        '        i = 0 : ReDim DirList(i)
        '        Do While entry.Length > 0
        '            If entry <> "." And entry <> ".." And (GetAttr(rootPath & "\" & entry) And vbDirectory) = vbDirectory Then
        '                ReDim Preserve DirList(i)
        '                DirList(i) = entry
        '                i = i + 1
        '            End If
        '            entry = Dir()
        '        Loop
        '        If i = 0 Then Exit Sub
        '        If sortConfigStoreFolders Then QuickSort(DirList, LBound(DirList), UBound(DirList))

        '        ' recursively build branched menu structure from dirEntries
        '        For i = 0 To UBound(DirList)
        '            theHostApp.StatusBar = "Filling DBConfigs Menu: " & rootPath & "\" & DirList(i)
        '            newBar = currentBar.Controls.Add(Type:=msoControlPopup)
        '            newBar.caption = DirList(i)
        '            newBar.Tag = DirList(i)
        '            readAllFiles(rootPath & "\" & DirList(i), newBar)
        '        Next
        '        Exit Sub
        'err1:
        '        Debug.Print("Error (" & Err.Description & ") in MenuHandler.readAllFiles")
        '        LogToEventViewer("Error (" & Err.Description & ") in MenuHandler.readAllFiles in " & Erl(), EventLogEntryType.Error, 1)
        '        Resume Next
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
        '        Dim newBar As CommandBarControl
        '        Dim newSubBar As CommandBarControl
        '        Static currentDepth As Integer

        '        On Error GoTo buildFileSepMenuCtrl_Err
        '        ' end node: add callable cmdbar entry
        '        If InStr(1, nameParts, " ") = 0 Or currentDepth > specialFolderMaxDepth - 1 Then
        '            On Error Resume Next
        '            newBar = specialConfigFoldersTempColl(newRootName & nameParts)
        '            If Err.Number <> 0 Then newSubBar = currentBar
        '            Err.Clear()
        '            newBar = newSubBar.Controls.Add(Type:=msoControlButton)
        '            Dim entryName : entryName = Mid$(fullPathName, InStrRev(fullPathName, "\") + 1)
        '            newBar.caption = Left$(entryName, Len(entryName) - 4)
        '            newBar.Parameter = fullPathName
        '            newBar.Tag = gsITEMLOADPREPARED_TAG
        '            ' leaf node: add popup menu entry
        '        Else
        '            Dim newName As String
        '            newName = Left$(nameParts, InStr(1, nameParts, " ") - 1)
        '            On Error Resume Next
        '            newBar = specialConfigFoldersTempColl(newRootName & newName)
        '            If Err.Number <> 0 Then
        '                newBar = currentBar.Controls.Add(Type:=msoControlPopup)
        '                newBar.caption = newName
        '                specialConfigFoldersTempColl.Add(newBar, newRootName & newName)
        '            End If
        '            Err.Clear()
        '            currentDepth = currentDepth + 1
        '            buildFileSepMenuCtrl(Mid$(nameParts, InStr(1, nameParts, " ") + 1), newBar, fullPathName, newRootName & newName, specialFolderMaxDepth)
        '            currentDepth = currentDepth - 1
        '        End If
        '        Exit Sub

        'buildFileSepMenuCtrl_Err:
        '        If VBDEBUG Then Debug.Print("Error (" & Err.Description & ") in MenuHandler.buildFileSepMenuCtrl") : Stop : Resume
        '        LogToEventViewer("Error (" & Err.Description & ") in MenuHandler.buildFileSepMenuCtrl in " & Erl(), EventLogEntryType.Error, 1)
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

End Class
