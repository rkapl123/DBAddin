Imports ExcelDna.Integration
Imports ExcelDna.Registration
Imports System.IO ' needed for logfile
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

''' <summary>Global variables and functions for DB Addin</summary>
Public Module DBAddin
    ''' <summary>currently selected environment for DB Functions</summary>
    Public selectedEnvironment As Integer
    ''' <summary>reference object for the Addins ribbon</summary>
    Public theRibbon As ExcelDna.Integration.CustomUI.IRibbonUI
    ''' <summary>Application object used for referencing objects</summary>
    Public hostApp As Object
    ''' <summary>environment definitions</summary>
    Public environdefs As String() = {}
    ''' <summary>DBMapper definition collections (for labels (key of nested dictionary) and target ranges (value of nested dictionary))</summary>
    Public DBMapperDefColl As Dictionary(Of String, Dictionary(Of String, Range))
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
        CnnTimeout = CInt(fetchSetting("CnnTimeout", "15"))
        CmdTimeout = CInt(fetchSetting("CmdTimeout", "60"))
        ConfigStoreFolder = fetchSetting("ConfigStoreFolder", String.Empty)
        specialConfigStoreFolders = Split(fetchSetting("specialConfigStoreFolders", String.Empty), ":")
        DefaultDBDateFormatting = CInt(fetchSetting("DefaultDBDateFormatting", "0"))
        ' load environments
        Dim i As Integer = 1
        ReDim Preserve environdefs(-1)
        Dim ConfigName As String
        Do
            ConfigName = fetchSetting("ConfigName" + i.ToString(), vbNullString)
            If Len(ConfigName) > 0 Then
                ReDim Preserve environdefs(environdefs.Length)
                environdefs(environdefs.Length - 1) = ConfigName + " - " + i.ToString()
            End If
            ' set selectedEnvironment
            If fetchSetting("ConstConnString" + i.ToString(), vbNullString) = ConstConnString Then
                selectedEnvironment = i - 1
            End If
            i += 1
        Loop Until Len(ConfigName) = 0
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
    Public Sub LogError(LogMessage As String, Optional ByRef exitMe As Boolean = False, Optional includeMsg As Boolean = True)
        Dim retval As Integer

        LogToEventViewer(LogMessage, EventLogEntryType.Error)
        'If Not automatedMapper Is Nothing Then automatedMapper.returnedErrorMessages = automatedMapper.returnedErrorMessages & LogMessage & vbCrLf
        If includeMsg Then retval = MsgBox(LogMessage, vbCritical + IIf(exitMe, vbOKCancel, vbOKOnly), "DBAddin: Internal Error !! ")
        If retval = vbCancel Then
            exitMe = True
        Else
            exitMe = False
        End If
    End Sub

    ''' <summary>Logs warning messages</summary>
    ''' <param name="LogMessage"></param>
    ''' <param name="exitMe"></param>
    ''' <param name="includeMsg"></param>
    Public Sub LogWarn(LogMessage As String, Optional ByRef exitMe As Boolean = False, Optional includeMsg As Boolean = True)
        Dim retval As Integer

        LogToEventViewer(LogMessage, EventLogEntryType.Warning)
        'If Not automatedMapper Is Nothing Then automatedMapper.returnedErrorMessages = automatedMapper.returnedErrorMessages & LogMessage & vbCrLf
        If includeMsg Then retval = MsgBox(LogMessage, vbCritical + IIf(exitMe, vbOKCancel, vbOKOnly), "DBAddin Error")
        If retval = vbCancel Then
            exitMe = True
        Else
            exitMe = False
        End If
    End Sub

    ''' <summary>Logs informational messages</summary>
    ''' <param name="LogMessage"></param>
    Public Sub LogInfo(LogMessage As String)
        If DebugAddin Then LogToEventViewer(LogMessage, EventLogEntryType.Information)
    End Sub

    <ExcelCommand(Name:="refreshData", ShortCut:="^R")>
    Public Sub refreshData()
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

    <ExcelCommand(Name:="jumpButton", ShortCut:="^J")>
    Public Sub jumpButton()
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
End Module

''' <summary>Connection class handling basic Events from Excel (Open, Close)</summary>
Public Class AddInEvents
    Implements IExcelAddIn

    WithEvents Application As Excel.Application

    ''' <summary>connect to Excel when opening Addin</summary>
    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        ' necessary for ExplicitRegistration of param arrays (https://groups.google.com/forum/#!topic/exceldna/kf76nqAqDUo)
        ExcelRegistration.GetExcelFunctions().ProcessParamsRegistrations().RegisterFunctions()
        ExcelRegistration.GetExcelCommands().RegisterCommands()
        Application = ExcelDnaUtil.Application
        theHostApp = ExcelDnaUtil.Application
        ' Register Ctrl+Shift+R to call refreshDB
        'XlCall.Excel(XlCall.xlcOnKey, "^R", "refreshData")
        Dim logfilename As String = "C:\\DBAddinlogs\\" + DateTime.Today.ToString("yyyyMMdd") + ".log"
        If Not Directory.Exists("C:\\DBAddinlogs") Then MkDir("C:\\DBAddinlogs")
        Try
            logfile = New StreamWriter(logfilename, True, System.Text.Encoding.GetEncoding(1252))
            logfile.AutoFlush = True
        Catch ex As Exception
            MsgBox("Exception occured when trying to create logfile " + logfilename + ": " + ex.Message)
        End Try
        LogToEventViewer("starting DBAddin", EventLogEntryType.Information)
        theMenuHandler = New MenuHandler
        theDBFuncEventHandler = New DBFuncEventHandler
        initSettings()
    End Sub

    ''' <summary>AutoClose cleans up after finishing addin</summary>
    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        On Error Resume Next
        theMenuHandler = Nothing
        theHostApp = Nothing
        theDBFuncEventHandler = Nothing
        'XlCall.Excel(XlCall.xlcOnKey, "^R")
    End Sub
    Private Sub Workbook_Save(Wb As Excel.Workbook, ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeSave
    End Sub
    Private Sub Workbook_Open(Wb As Excel.Workbook) Handles Application.WorkbookOpen
        ' ribbon invalidation is being treated in WorkbookActivate...
    End Sub

    ''' <summary>Workbook_Activate: gets defined named ranges for DBMapper invocation in the current workbook and updates Ribbon with it</summary>
    Private Sub Workbook_Activate(Wb As Excel.Workbook) Handles Application.WorkbookActivate
        ' load DBMapper definitions
        DBMapperDefColl = New Dictionary(Of String, Dictionary(Of String, Range))
        Dim i As Integer = 0
        For Each namedrange As Name In hostApp.ActiveWorkbook.Names
            Dim cleanname As String = Replace(namedrange.Name, namedrange.Parent.Name & "!", "")
            If Left(cleanname, 8) = "DBMapper" Then
                If InStr(namedrange.RefersTo, "#REF!") > 0 Then LogError("DBMapper definitions range " + namedrange.Parent.name + "!" + namedrange.Name + " contains #REF!", vbOKOnly + vbCritical, "DBAddin: DBMapper definitions range error")
                Dim nodeName As String = Replace(Replace(namedrange.Name, "DBMapper", ""), namedrange.Parent.Name & "!", "")
                If nodeName = "" Then nodeName = "UnnamedDBMapper"

                Dim defColl As Dictionary(Of String, Range)
                If Not DBMapperDefColl.ContainsKey("ID" + i.ToString()) Then
                    ' add to new sheet "menu"
                    defColl = New Dictionary(Of String, Range)
                    defColl.Add(nodeName, namedrange.RefersToRange)
                    DBMapperDefColl.Add("ID" + i.ToString(), defColl)
                    i += 1
                Else
                    ' add definition to existing sheet "menu"
                    defColl = DBMapperDefColl("ID" + i.ToString())
                    defColl.Add(nodeName, namedrange.RefersToRange)
                End If
            End If
            If i >= 15 Then LogError("Not more than 15 sheets with DBMapper definitions possible, ignoring definitions in sheet " + namedrange.Parent.Name)
        Next
        DBAddin.theRibbon.Invalidate()
    End Sub
End Class
