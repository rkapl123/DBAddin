Imports ExcelDna.Integration
Imports ExcelDna.Registration
Imports System.IO ' needed for logfile
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

''' <summary>All global Variables for DBFuncBuilder and Functions and some global accessible functions</summary>
Public Module Globals
    ''' <summary>logfile for messages</summary>
    Public logfile As StreamWriter

    ''' <summary>encapsulates setting fetching (currently registry)</summary>
    ''' <param name="Key"></param>
    ''' <param name="defaultValue"></param>
    ''' <returns>the setting value</returns>
    Public Function fetchSetting(Key As String, defaultValue As Object) As Object
        fetchSetting = GetSetting("DBAddin", "Settings", Key, defaultValue)
    End Function

    ''' <summary>encapsulates setting storing (currently registry)</summary>
    ''' <param name="Key"></param>
    ''' <param name="Value"></param>
    Public Sub storeSetting(Key As String, Value As Object)
        SaveSetting("DBAddin", "Settings", Key, Value)
    End Sub

    ''' <summary>initializes global configuration variables from registry</summary>
    Public Sub initSettings()
        DebugAddin = fetchSetting("DebugAddin", False)
        ConstConnString = fetchSetting("ConstConnString", String.Empty)
        DBidentifierCCS = fetchSetting("DBidentifierCCS", "Database=")
        DBidentifierODBC = fetchSetting("DBidentifierODBC", "Database=")
        CnnTimeout = CInt(fetchSetting("CnnTimeout", "15"))
        CmdTimeout = CInt(fetchSetting("CmdTimeout", "60"))
        ConfigStoreFolder = fetchSetting("ConfigStoreFolder", String.Empty)
        specialConfigStoreFolders = Split(fetchSetting("specialConfigStoreFolders", String.Empty), ":")
        DefaultDBDateFormatting = CInt(fetchSetting("DefaultDBDateFormatting", "0"))
    End Sub

    ''' <summary>Logs sErrMsg of eEventType to Logfile</summary>
    ''' <param name="sErrMsg"></param>
    ''' <param name="eEventType"></param>
    ''' <returns></returns>
    Public Function LogToEventViewer(sErrMsg As String, eEventType As EventLogEntryType) As Boolean
        Try
            logfile.WriteLine(Now().ToString() & vbTab & IIf(eEventType = EventLogEntryType.Error, "ERROR", IIf(eEventType = EventLogEntryType.Information, "INFO", "WARNING")) & vbTab & sErrMsg)
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    ''' <summary>Logs error messages</summary>
    ''' <param name="LogMessage"></param>
    ''' <param name="includeMsg"></param>
    ''' <param name="exitMe"></param>
    ''' <param name="category"></param>
    Public Sub LogError(LogMessage As String, Optional includeMsg As Boolean = True, Optional ByRef exitMe As Boolean = False, Optional category As Long = 2)
        Dim retval As Integer

        LogToEventViewer(LogMessage, EventLogEntryType.Error)
        'If Not automatedMapper Is Nothing Then automatedMapper.returnedErrorMessages = automatedMapper.returnedErrorMessages & LogMessage & vbCrLf
        'If includeMsg And automatedMapper Is Nothing Then retval = MsgBox(LogMessage, vbCritical + IIf(exitMe, vbOKCancel, vbOKOnly), "DBAddin: Internal Error !! ")
        If retval = vbCancel Then
            exitMe = True
        Else
            exitMe = False
        End If
    End Sub

    ''' <summary>Logs warning messages</summary>
    ''' <param name="LogMessage"></param>
    ''' <param name="exitMe"></param>
    ''' <param name="category"></param>
    ''' <param name="includeMsg"></param>
    Public Sub LogWarn(LogMessage As String, Optional ByRef exitMe As Boolean = False, Optional category As Long = 2, Optional includeMsg As Boolean = True)
        Dim retval As Integer

        LogToEventViewer(LogMessage, EventLogEntryType.Warning)
        'If Not automatedMapper Is Nothing Then automatedMapper.returnedErrorMessages = automatedMapper.returnedErrorMessages & LogMessage & vbCrLf
        'If includeMsg And automatedMapper Is Nothing Then retval = MsgBox(LogMessage, vbCritical + IIf(exitMe, vbOKCancel, vbOKOnly), "DBAddin Error")
        If retval = vbCancel Then
            exitMe = True
        Else
            exitMe = False
        End If
    End Sub

    ''' <summary>Logs informational messages</summary>
    ''' <param name="LogMessage"></param>
    ''' <param name="category"></param>
    Public Sub LogInfo(LogMessage As String, Optional category As Long = 2)
        If DebugAddin Then LogToEventViewer(LogMessage, EventLogEntryType.Information)
    End Sub

    ' general Global objects/variables
    ''' <summary>Application object used for referencing objects</summary>
    Public theHostApp As Object
    ''' <summary>ribbon menu handler</summary>
    Public theMenuHandler As MenuHandler
    ''' <summary>for interrupting long running operations with Ctl-Break</summary>
    Public Interrupted As Boolean

    ''' <summary>the environment (for Mapper special cases "Test", "Development" or String.Empty (prod))</summary>
    Public env As String

    ' Global settings
    Public DebugAddin As Boolean
    ''' <summary>Default ConnectionString, if no connection string is given by user....</summary>
    Public ConstConnString As String
    ''' <summary>the tag used to identify the Database name within the ConstConnString</summary>
    Public DBidentifierCCS As String
    ''' <summary>the tag used to identify the Database name within the connection string returned by MSQuery</summary>
    Public DBidentifierODBC As String
    ''' <summary>the folder used to store predefined DB item definitions</summary>
    Public ConfigStoreFolder As String
    ''' <summary>Array of special ConfigStoreFolders for non default treatment of Name Separation (Camelcase) and max depth</summary>
    Public specialConfigStoreFolders() As String
    ''' <summary>should config stores be sorted alphabetically</summary>
    Public sortConfigStoreFolders As Boolean
    ''' <summary>global connection timeout (can't be set in DB functions)</summary>
    Public CnnTimeout As Integer
    ''' <summary>global command timeout (can't be set in DB functions)</summary>
    Public CmdTimeout As Integer
    ''' <summary>default formatting style used in DBDate</summary>
    Public DefaultDBDateFormatting As Integer

    ' Global flags
    ''' <summary>prevent multiple connection retries for each function in case of error</summary>
    Public dontTryConnection As Boolean
    ''' <summary>avoid entering dblistfetch function during clearing of listfetch areas (before saving)</summary>
    Public dontCalcWhileClearing As Boolean

    ' Global objects/variables for DBFuncs
    ''' <summary>store target filter in case of empty data lists</summary>
    Public targetFilterCont As Collection
    ''' <summary>global event class, mainly for calc event procedure</summary>
    Public theDBFuncEventHandler As DBFuncEventHandler
    ''' <summary>global collection of information transport containers between function and calc event procedure</summary>
    Public allCalcContainers As Collection
    ''' <summary>global collection of information transport containers between function and calc event procedure</summary>
    Public allStatusContainers As Collection
End Module

''' <summary>Connection class handling basic Events from Excel (Open, Close)</summary>
Public Class AddInEvents
    Implements IExcelAddIn

    WithEvents Application As Excel.Application

    ''' <summary>connect to Excel when opening Addin</summary>
    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        ' necessary for ExplicitRegistration of param arrays (https://groups.google.com/forum/#!topic/exceldna/kf76nqAqDUo)
        ExcelRegistration.GetExcelFunctions().ProcessParamsRegistrations().RegisterFunctions()
        Application = ExcelDnaUtil.Application
        theHostApp = ExcelDnaUtil.Application
        Dim logfilename As String = "C:\\DBAddinlogs\\" + DateTime.Today.ToString("yyyyMMdd") + ".log"
        If Not Directory.Exists("C:\\DBAddinlogs") Then MkDir("C:\\DBAddinlogs")
        Try
            logfile = New StreamWriter(logfilename, True, System.Text.Encoding.GetEncoding(1252))
            logfile.AutoFlush = True
        Catch ex As Exception
            MsgBox("Exception occured when trying to create logfile " + logfilename + ": " + ex.Message)
        End Try
        LogToEventViewer("starting DBAddin", EventLogEntryType.Information)
        initSettings()
        theMenuHandler = New MenuHandler
        theDBFuncEventHandler = New DBFuncEventHandler
    End Sub

    ''' <summary>AutoClose cleans up after finishing addin</summary>
    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        On Error Resume Next
        theMenuHandler = Nothing
        theHostApp = Nothing
        theDBFuncEventHandler = Nothing
    End Sub
    Private Sub Workbook_Save(Wb As Excel.Workbook, ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeSave
    End Sub
    Private Sub Workbook_Open(Wb As Excel.Workbook) Handles Application.WorkbookOpen
        ' ribbon invalidation is being treated in WorkbookActivate...
    End Sub
    Private Sub Workbook_Activate(Wb As Excel.Workbook) Handles Application.WorkbookActivate
        ' load DBMapper definitions
        getDBMapperNames()
        DBAddin.theRibbon.Invalidate()
    End Sub
End Class

Public Module DBAddin

    ''' <summary>reference object for the Addins ribbon</summary>
    Public theRibbon As ExcelDna.Integration.CustomUI.IRibbonUI
    ''' <summary>Application object used for referencing objects</summary>
    Public hostApp As Object

    Public defnames As String() = {}
    Public defsheetMap As Dictionary(Of String, String)
    Public defsheetColl As Dictionary(Of String, Dictionary(Of String, Range))

    Public Sub doRefreshData()
        initSettings()

        ' enable events in case there were some problems in procedure with EnableEvents = false
        On Error Resume Next
        hostApp.EnableEvents = True
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
        underlyingName = getDBRangeName(hostApp.ActiveCell)
        hostApp.ScreenUpdating = True
        If underlyingName Is Nothing Then
            ' reset query cache, so we really get new data !
            theDBFuncEventHandler.queryCache = New Collection
            refreshDBFunctions(hostApp.ActiveWorkbook)
            ' general refresh: also refresh all embedded queries and pivot tables..
            'On Error Resume Next
            'Dim ws     As Excel.Worksheet
            'Dim qrytbl As Excel.QueryTable
            'Dim pivottbl As Excel.PivotTable

            'For Each ws In hostApp.ActiveWorkbook.Worksheets
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

                If Not hostApp.Range(jumpName).Parent Is hostApp.ActiveSheet Then
                    hostApp.ScreenUpdating = False
                    theDBFuncEventHandler.origWS = hostApp.ActiveSheet
                    On Error Resume Next
                    hostApp.Range(jumpName).Parent.Select
                    On Error GoTo err1
                End If
                hostApp.Range(jumpName).Dirty
                ' we're being called on a target area
            ElseIf Left$(jumpName, 9) = "DBFtarget" Then
                jumpName = Replace(jumpName, "DBFtarget", "DBFsource", 1, , vbTextCompare)

                If Not hostApp.Range(jumpName).Parent Is hostApp.ActiveSheet Then
                    hostApp.ScreenUpdating = False
                    theDBFuncEventHandler.origWS = hostApp.ActiveSheet
                    On Error Resume Next
                    hostApp.Range(jumpName).Parent.Select
                    On Error GoTo err1
                End If
                hostApp.Range(jumpName).Dirty
                ' we're being called on a source (invoking function) cell
            ElseIf Left$(jumpName, 9) = "DBFsource" Then
                On Error Resume Next
                hostApp.Range(jumpName).Dirty
                On Error GoTo err1
            Else
                refreshDBFunctions(hostApp.ActiveWorkbook)
            End If
        End If

        Exit Sub
err1:
        LogToEventViewer("Error (" & Err.Description & ") in MenuHandler.refreshData in " & Erl(), EventLogEntryType.Error)
    End Sub

    ''
    '  jumps from DB function to data area and back
    Public Sub doJumpButton()
        Dim underlyingName As Excel.Name
        underlyingName = getDBRangeName(hostApp.ActiveCell)

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
        hostApp.Range(jumpName).Parent.Select
        hostApp.Range(jumpName).Select
        If Err.Number <> 0 Then LogWarn("Can't jump to target/source, corresponding workbook open? " & Err.Description, 1)
        Err.Clear()
    End Sub

    ' gets defined named ranges for DBMapper invocation in the current workbook 
    Public Function getDBMapperNames() As String
        ReDim Preserve defnames(-1)
        defsheetColl = New Dictionary(Of String, Dictionary(Of String, Range))
        defsheetMap = New Dictionary(Of String, String)
        Dim i As Integer = 0
        For Each namedrange As Name In hostApp.ActiveWorkbook.Names
            Dim cleanname As String = Replace(namedrange.Name, namedrange.Parent.Name & "!", "")
            If Left(cleanname, 8) = "DBMapper" Then
                If InStr(namedrange.RefersTo, "#REF!") > 0 Then Return "DBMapper definitions range " + namedrange.Parent.name + "!" + namedrange.Name + " contains #REF!"
                ' final name of entry is without DBMapper and !
                Dim finalname As String = Replace(Replace(namedrange.Name, "DBMapper", ""), "!", "")
                Dim nodeName As String = Replace(Replace(namedrange.Name, "DBMapper", ""), namedrange.Parent.Name & "!", "")
                If nodeName = "" Then nodeName = "MainDBMapper"
                If Not InStr(namedrange.Name, "!") > 0 Then
                    finalname = hostApp.ActiveWorkbook.Name + finalname
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

    Public Const menufilename As String = "C:\\DBAddinlogs\\menufile.xml"

    Private specialConfigFoldersTempColl As Collection
    ' Limitation by Ribbon: 5 levels: 1 top level, 1 folder level (Database foldername) -> 3 left
    Const specialFolderMaxDepth As Integer = 3
    Public xnspace As XNamespace = "http://schemas.microsoft.com/office/2009/07/customui"
    ''
    ' create the config tree menu
    Public Sub createConfigTreeMenu()
        Dim currentBar, button As XElement
        Dim menuXML As String

        If Not Directory.Exists(ConfigStoreFolder) Then
            menuXML = "<menu xmlns='http://schemas.microsoft.com/office/2009/07/customui' id='DBConfigs' label='No Config Store !!' screentip='Couldn't find predefined config store folder [" & ConfigStoreFolder & "], please check registry setting for config store folder location and refresh below!'>" +
                "<button id='refreshConfig' label='refresh DBConfig Tree' imageMso='Refresh' onAction='refreshDBConfigTree'/></menu>"
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
            readAllFiles(ConfigStoreFolder, currentBar)
            specialConfigFoldersTempColl = Nothing
            hostApp.StatusBar = String.Empty
            currentBar.SetAttributeValue("xmlns", "http://schemas.microsoft.com/office/2009/07/customui")
            menuXML = currentBar.ToString()
        End If
        If Not Directory.Exists("C:\\DBAddinlogs") Then MkDir("C:\\DBAddinlogs")
        Dim menufile As StreamWriter = Nothing
        Try
            menufile = New StreamWriter(menufilename, False, System.Text.Encoding.GetEncoding(1252))
        Catch ex As Exception
            MsgBox("Exception occured when trying to create menufile " + menufilename + ": " + ex.Message)
        End Try
        menufile.Write(menuXML)
        menufile.Close()
        Exit Sub

err1:
        LogToEventViewer("Error (" & Err.Description & ") in MenuHandler.createTreeMenu in " & Erl(), EventLogEntryType.Error)
    End Sub

    ''
    ' reads all files contained in rootPath and its subfolders (recursively)
    '             and adds them to currentBar and submenus (recursively)
    ' @param rootPath
    ' @param currentBar
    Public Sub readAllFiles(rootPath As String, ByRef currentBar As XElement)
        Dim newBar As XElement = Nothing
        Dim configType As String, entry As String
        Dim i As Long
        Dim DirList() As DirectoryInfo
        Dim fileList() As FileSystemInfo
        Dim specialConfigStoreSeparator As String

        On Error GoTo err1

        ' read all leaf node entries (files) and sort them by name to create action menus
        Dim di As DirectoryInfo = New DirectoryInfo(rootPath)
        fileList = di.GetFileSystemInfos("*.xcl").OrderBy(Function(fi) fi.Name).ToArray()
        If fileList.Length > 0 Then

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
                Dim firstCharLevel As Boolean = CBool(fetchSetting(spclFolder & "FirstLetterLevel", "False"))
                specialConfigStoreSeparator = fetchSetting(spclFolder & "Separator", String.Empty)
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
                            i = i + 2
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
                    newBar.SetAttributeValue("id", theFolder & i)
                    newBar.SetAttributeValue("tag", rootPath & "\" & fileList(i).Name)
                    newBar.SetAttributeValue("label", Left$(fileList(i).Name, Len(fileList(i).Name) - 4))
                    newBar.SetAttributeValue("onAction", "getConfig")
                    currentBar.Add(newBar)
                Next
            End If
        End If

        ' read all dir entries and sort them by name
        DirList = di.GetDirectories().OrderBy(Function(fi) fi.Name).ToArray()
        If DirList.Length = 0 Then Exit Sub

        ' recursively build branched menu structure from dirEntries
        For i = 0 To UBound(DirList)
            hostApp.StatusBar = "Filling DBConfigs Menu: " & rootPath & "\" & DirList(i).Name
            newBar = New XElement(xnspace + "menu")
            newBar.SetAttributeValue("id", DirList(i).Name)
            newBar.SetAttributeValue("label", DirList(i).Name)
            currentBar.Add(newBar)
            readAllFiles(rootPath & "\" & DirList(i).Name, newBar)
        Next
        Exit Sub
err1:
        LogToEventViewer("Error (" & Err.Description & ") in MenuHandler.readAllFiles in " & Erl(), EventLogEntryType.Error)
    End Sub

    ' parses Substrings contained in nameParts (recursively)
    '             and adds them to currentBar and submenus (recursively)
    ' @param nameParts..tokenized string
    ' @param currentBar
    ' @param fullPathName
    ' @param newRootName
    ' @param specialFolderMaxDepth
    Public Sub buildFileSepMenuCtrl(nameParts As String, ByRef currentBar As XElement,
                                         fullPathName As String, newRootName As String, specialFolderMaxDepth As Integer)
        Dim newBar As XElement
        Static currentDepth As Integer

        On Error GoTo buildFileSepMenuCtrl_Err
        ' end node: add callable entry
        If InStr(1, nameParts, " ") = 0 Or currentDepth > specialFolderMaxDepth - 1 Then
            If specialConfigFoldersTempColl.Contains(newRootName & nameParts) Then
                newBar = specialConfigFoldersTempColl(newRootName & nameParts)
            Else
                Dim entryName As String = Mid$(fullPathName, InStrRev(fullPathName, "\") + 1)
                newBar = New XElement(xnspace + "button")
                newBar.SetAttributeValue("id", newRootName & entryName)
                newBar.SetAttributeValue("label", Left$(entryName, Len(entryName) - 4))
                newBar.SetAttributeValue("tag", fullPathName)
                newBar.SetAttributeValue("onAction", "getConfig")
            End If
            currentBar.Add(newBar)
        Else  ' branch node: add new menu, recursively descend
            Dim newName As String = Left$(nameParts, InStr(1, nameParts, " ") - 1)
            If specialConfigFoldersTempColl.Contains(newRootName & newName) Then
                newBar = specialConfigFoldersTempColl(newRootName & newName)
            Else
                newBar = New XElement(xnspace + "menu")
                newBar.SetAttributeValue("id", newRootName & newName)
                newBar.SetAttributeValue("label", newName)
                specialConfigFoldersTempColl.Add(newBar, newRootName & newName)
                currentBar.Add(newBar)
            End If
            currentDepth = currentDepth + 1
            buildFileSepMenuCtrl(Mid$(nameParts, InStr(1, nameParts, " ") + 1), newBar, fullPathName, newRootName & newName, specialFolderMaxDepth)
            currentDepth = currentDepth - 1
        End If
        Exit Sub

buildFileSepMenuCtrl_Err:
        LogToEventViewer("Error (" & Err.Description & ") in MenuHandler.buildFileSepMenuCtrl in " & Erl(), EventLogEntryType.Error)
    End Sub

    ''
    ' return CamelCase string in space separated parts (tokenize String following case switch or when specialConfigStoreSeparator occurs, if given)
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

End Module
