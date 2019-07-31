Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

''' <summary>Global variables and functions for DB Addin</summary>
Public Module DBAddin
    ' general Global objects/variables
    ''' <summary>ribbon menu handler</summary>
    Public theMenuHandler As MenuHandler
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
    ''' <summary>global connection timeout (can't be set in DB functions)</summary>
    Public CnnTimeout As Integer
    ''' <summary>global command timeout (can't be set in DB functions)</summary>
    Public CmdTimeout As Integer
    ''' <summary>default formatting style used in DBDate</summary>
    Public DefaultDBDateFormatting As Integer

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

    ''' <summary>refresh DB Functions (and - if called from outside any db function area - all other external data ranges)</summary>
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

        ' also reset the database connection in case of errors (might be nothing or not open...)
        Try : conn.Close() : Catch ex As Exception : End Try
        conn = Nothing
        dontTryConnection = False
        Try
            ' reset query cache, so we really get new data !
            queryCache = New Collection
            StatusCollection = New Collection
            Dim underlyingName As Excel.Name
            underlyingName = getDBRangeName(hostApp.ActiveCell)
            hostApp.ScreenUpdating = True
            ' now for DBListfetch/DBRowfetch resetting, first outside of all db function areas...
            If underlyingName Is Nothing Then
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
            Else ' then inside a db function area (target or source = function cell)
                Dim jumpName As String
                jumpName = underlyingName.Name
                ' we're being called on a target functions area (additionally given in DBListFetch)
                If Left$(jumpName, 10) = "DBFtargetF" Then
                    jumpName = Replace(jumpName, "DBFtargetF", "DBFsource", 1, , vbTextCompare)
                    If Not hostApp.Range(jumpName).Parent Is hostApp.ActiveSheet Then
                        hostApp.ScreenUpdating = False
                        origWS = hostApp.ActiveSheet
                        Try : hostApp.Range(jumpName).Parent.Select : Catch ex As Exception : End Try
                    End If
                    hostApp.Range(jumpName).Dirty()
                    ' we're being called on a target area
                ElseIf Left$(jumpName, 9) = "DBFtarget" Then
                    jumpName = Replace(jumpName, "DBFtarget", "DBFsource", 1, , vbTextCompare)
                    ' return to source functions sheet to work around Dirty method problem (cell's sheet needs to be selected for Dirty to work on that cell)
                    If Not hostApp.Range(jumpName).Parent Is hostApp.ActiveSheet Then
                        hostApp.ScreenUpdating = False
                        origWS = hostApp.ActiveSheet
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

    ''' <summary>jumps between DB Function and target area</summary>
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
        Try
            hostApp.Range(jumpName).Parent.Select()
            hostApp.Range(jumpName).Select()
        Catch ex As Exception
            LogWarn("Can't jump to target/source, corresponding workbook open? " & ex.Message)
        End Try
    End Sub
End Module
