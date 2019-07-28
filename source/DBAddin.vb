Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

''' <summary>Global variables and functions for DB Addin</summary>
Public Module DBAddin
    ' general Global objects/variables
    ''' <summary>ribbon menu handler</summary>
    Public theMenuHandler As MenuHandler
    ''' <summary>for interrupting long running operations with Ctl-Break</summary>
    Public Interrupted As Boolean
    ''' <summary>currently selected environment for DB Functions</summary>
    Public selectedEnvironment As Integer
    ''' <summary>reference object for the Addins ribbon</summary>
    Public theRibbon As ExcelDna.Integration.CustomUI.IRibbonUI
    ''' <summary>Excel Application object used for referencing objects</summary>
    Public hostApp As Application
    ''' <summary>environment definitions</summary>
    Public environdefs As String() = {}
    ''' <summary>DBMapper definition collections (for labels (key of nested dictionary) and target ranges (value of nested dictionary))</summary>
    Public DBMapperDefColl As Dictionary(Of String, Dictionary(Of String, Range))
    ''' <summary>the selected event level in the About box</summary>
    Public EventLevelsSelected As String
    ''' <summary>the log listener</summary>
    Public theLogListener As TraceListener

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

    ''' <summary>encapsulates setting fetching (currently registry)</summary>
    ''' <param name="Key"></param>
    ''' <param name="defaultValue"></param>
    ''' <returns>the setting value</returns>
    Public Function fetchSetting(Key As String, defaultValue As String) As String
        fetchSetting = GetSetting("DBAddin", "Settings", Key, defaultValue)
    End Function

    ''' <summary>encapsulates setting storing (currently registry)</summary>
    ''' <param name="Key"></param>
    ''' <param name="Value"></param>
    Public Sub storeSetting(Key As String, Value As String)
        SaveSetting("DBAddin", "Settings", Key, Value)
    End Sub

    ''' <summary>initializes global configuration variables from registry</summary>
    Public Sub initSettings()
        Try
            DebugAddin = CBool(fetchSetting("DebugAddin", "False"))
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
        Catch ex As Exception
            LogError("Error in initialization of Settings (DBAddin.initSettings):" + ex.Message)
        End Try
    End Sub

    ''' <summary>Logs Message of eEventType to System.Diagnostics.Trace</summary>
    ''' <param name="Message"></param>
    ''' <param name="eEventType"></param>
    ''' <param name="caller"></param>
    Public Sub WriteToLog(Message As String, eEventType As EventLogEntryType, Optional caller As String = "")
        If caller = "" Then
            Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
            caller = theMethod.ReflectedType.FullName & "." & theMethod.Name
        End If
        Select Case eEventType
            Case EventLogEntryType.Information : Trace.TraceInformation("{0}: {1}", caller, Message)
            Case EventLogEntryType.Warning : Trace.TraceWarning("{0}: {1}", caller, Message)
            Case EventLogEntryType.Error : Trace.TraceError("{0}: {1}", caller, Message)
        End Select
    End Sub

    ''' <summary>Logs error messages</summary>
    ''' <param name="LogMessage"></param>
    ''' <param name="exitMe"></param>
    ''' <param name="includeMsg"></param>
    Public Sub LogError(LogMessage As String, Optional ByRef exitMe As Boolean = False, Optional includeMsg As Boolean = True)
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Dim caller As String = theMethod.ReflectedType.FullName & "." & theMethod.Name
        WriteToLog(LogMessage, EventLogEntryType.Error, caller)
        If includeMsg Then
            Dim retval As Integer = MsgBox(LogMessage, vbCritical + IIf(exitMe, vbOKCancel, vbOKOnly), "DBAddin Error")
            If retval = vbCancel Then exitMe = True
        End If
    End Sub

    ''' <summary>Logs warning messages</summary>
    ''' <param name="LogMessage"></param>
    ''' <param name="exitMe"></param>
    ''' <param name="includeMsg"></param>
    Public Sub LogWarn(LogMessage As String, Optional ByRef exitMe As Boolean = False, Optional includeMsg As Boolean = True)
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Dim caller As String = theMethod.ReflectedType.FullName & "." & theMethod.Name
        WriteToLog(LogMessage, EventLogEntryType.Warning, caller)
        If includeMsg Then
            Dim retval As Integer = MsgBox(LogMessage, vbExclamation + IIf(exitMe, vbOKCancel, vbOKOnly), "DBAddin Warning")
            If retval = vbCancel Then exitMe = True
        End If
    End Sub

    ''' <summary>Logs informational messages</summary>
    ''' <param name="LogMessage"></param>
    Public Sub LogInfo(LogMessage As String)
        If DebugAddin Then
            Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
            Dim caller As String = theMethod.ReflectedType.FullName & "." & theMethod.Name
            WriteToLog(LogMessage, EventLogEntryType.Information, caller)
        End If
    End Sub

    <ExcelCommand(Name:="refreshData", ShortCut:="^R")>
    Public Sub refreshData()
        initSettings()

        ' enable events in case there were some problems in procedure with EnableEvents = false
        Try
            hostApp.EnableEvents = True
        Catch ex As Exception
            LogError("Can't refresh data while lookup dropdown is open !!")
            Exit Sub
        End Try

        ' also reset the database connection in case of errors...
        theDBFuncEventHandler.cnn.Close()
        theDBFuncEventHandler.cnn = Nothing
        dontTryConnection = False
        Try
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
                Try
                    Dim ws As Excel.Worksheet
                    Dim qrytbl As Excel.QueryTable
                    Dim pivottbl As Excel.PivotTable

                    For Each ws In hostApp.ActiveWorkbook.Worksheets
                        For Each qrytbl In ws.QueryTables
                            qrytbl.Refresh()
                        Next
                        For Each pivottbl In ws.PivotTables
                            pivottbl.PivotCache.Refresh()
                        Next
                    Next
                Catch ex As Exception
                End Try
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
                        Try : hostApp.Range(jumpName).Parent.Select : Catch ex As Exception : End Try
                    End If
                    hostApp.Range(jumpName).Dirty()
                    ' we're being called on a target area
                ElseIf Left$(jumpName, 9) = "DBFtarget" Then
                    jumpName = Replace(jumpName, "DBFtarget", "DBFsource", 1, , vbTextCompare)

                    If Not hostApp.Range(jumpName).Parent Is hostApp.ActiveSheet Then
                        hostApp.ScreenUpdating = False
                        theDBFuncEventHandler.origWS = hostApp.ActiveSheet
                        Try : hostApp.Range(jumpName).Parent.Select : Catch ex As Exception : End Try
                        hostApp.Range(jumpName).Parent.Select
                    End If
                    hostApp.Range(jumpName).Dirty()
                    ' we're being called on a source (invoking function) cell
                ElseIf Left$(jumpName, 9) = "DBFsource" Then
                    Try : hostApp.Range(jumpName).Dirty() : Catch ex As Exception : End Try
                Else
                    refreshDBFunctions(hostApp.ActiveWorkbook)
                End If
            End If
        Catch ex As Exception
            WriteToLog("Error (" & Err.Description & ") in MenuHandler.refreshData in " & Erl(), EventLogEntryType.Warning)
        End Try
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
        hostApp.Range(jumpName).Parent.Select()
        hostApp.Range(jumpName).Select()
        If Err.Number <> 0 Then LogWarn("Can't jump to target/source, corresponding workbook open? " & Err.Description, 1)
        Err.Clear()
    End Sub
End Module
