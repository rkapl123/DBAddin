Imports ExcelDna.Integration
Imports ExcelDna.Registration
Imports System.IO ' needed for logfile
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

''' <summary>Global variables and functions for DB Addin</summary>
Public Module DBAddin

    ''' <summary>reference object for the Addins ribbon</summary>
    Public theRibbon As ExcelDna.Integration.CustomUI.IRibbonUI
    ''' <summary>Application object used for referencing objects</summary>
    Public hostApp As Object
    ''' <summary>environment definitions</summary>
    Public environdefs As String() = {}
    ''' <summary>DBMapper definition maps (for labels)</summary>
    Public DBMapperDefMap As Dictionary(Of String, String)
    ''' <summary>DBMapper definition collections (for target ranges)</summary>
    Public DBMapperDefColl As Dictionary(Of String, Dictionary(Of String, Range))

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

    ''' <summary>Workbook_Activate: gets defined named ranges for DBMapper invocation in the current workbook and updates Ribbon with it</summary>
    Private Sub Workbook_Activate(Wb As Excel.Workbook) Handles Application.WorkbookActivate
        ' load DBMapper definitions
        DBMapperDefColl = New Dictionary(Of String, Dictionary(Of String, Range))
        DBMapperDefMap = New Dictionary(Of String, String)
        Dim i As Integer = 0
        For Each namedrange As Name In hostApp.ActiveWorkbook.Names
            Dim cleanname As String = Replace(namedrange.Name, namedrange.Parent.Name & "!", "")
            If Left(cleanname, 8) = "DBMapper" Then
                If InStr(namedrange.RefersTo, "#REF!") > 0 Then MsgBox("DBMapper definitions range " + namedrange.Parent.name + "!" + namedrange.Name + " contains #REF!", vbOKOnly + vbCritical, "DBAddin: DBMapper definitions range error")
                ' final name of entry is without DBMapper and !
                Dim finalname As String = Replace(Replace(namedrange.Name, "DBMapper", ""), "!", "")
                Dim nodeName As String = Replace(Replace(namedrange.Name, "DBMapper", ""), namedrange.Parent.Name & "!", "")
                If nodeName = "" Then nodeName = "MainDBMapper"
                If Not InStr(namedrange.Name, "!") > 0 Then
                    finalname = hostApp.ActiveWorkbook.Name + finalname
                End If

                Dim defColl As Dictionary(Of String, Range)
                If Not DBMapperDefColl.ContainsKey(namedrange.Parent.Name) Then
                    ' add to new sheet "menu"
                    defColl = New Dictionary(Of String, Range)
                    defColl.Add(nodeName, namedrange.RefersToRange)
                    DBMapperDefColl.Add(namedrange.Parent.Name, defColl)
                    DBMapperDefMap.Add("ID" + i.ToString(), namedrange.Parent.Name)
                    i += 1
                Else
                    ' add definition to existing sheet "menu"
                    defColl = DBMapperDefColl(namedrange.Parent.Name)
                    defColl.Add(nodeName, namedrange.RefersToRange)
                End If
            End If
        Next
        DBAddin.theRibbon.Invalidate()
    End Sub
End Class
