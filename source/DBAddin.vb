Imports Microsoft.Office.Interop.Excel
Imports ExcelDna.Integration
Imports System.Runtime.InteropServices
Imports ExcelDna.Integration.CustomUI
Imports ExcelDna.Registration
Imports System.IO ' needed for logfile

''' <summary>All global Variables for DBFuncBuilder and Functions and some global accessible functions</summary>
Public Module Globals
    ''' <summary>For Debugging purpose</summary>
    Public DEBUGME As Boolean
    ''' <summary>the eventlog for the addin session</summary>
    Public myEventlog As EventLog
    ''' <summary>in case of Automation, use this to communicate collected logged Warnings and Errors back</summary>
    Public automatedMapper As Mapper
    ''' <summary>logfile for messages</summary>
    Public logfile As StreamWriter

    ''' <summary>encapsulates setting fetching (currently registry)</summary>
    ''' <param name="Key"></param>
    ''' <param name="defaultValue"></param>
    ''' <param name="DBSheetSetting"></param>
    ''' <returns>the setting value</returns>
    Public Function fetchSetting(Key As String, defaultValue As Object, Optional DBSheetSetting As Boolean = False) As Object
        fetchSetting = GetSetting("DBAddin", IIf(DBSheetSetting, "DBSheetSettings", "Settings"), Key, defaultValue)
    End Function

    ''' <summary>encapsulates setting storing (currently registry)</summary>
    ''' <param name="Key"></param>
    ''' <param name="Value"></param>
    ''' <param name="DBSheetSetting"></param>
    Public Sub storeSetting(Key As String, Value As Object, Optional DBSheetSetting As Boolean = False)
        SaveSetting("DBAddin", IIf(DBSheetSetting, "DBSheetSettings", "Settings"), Key, Value)
    End Sub

    ''' <summary>initializes global configuration variables from registry</summary>
    Public Sub initSettings()
        ConstConnString = fetchSetting("ConstConnString", vbNullString)
        DBidentifierCCS = fetchSetting("DBidentifierCCS", "Database=")
        DBidentifierODBC = fetchSetting("DBidentifierODBC", "Database=")
        CnnTimeout = CInt(fetchSetting("CnnTimeout", "15"))
        CmdTimeout = CInt(fetchSetting("CmdTimeout", "60"))
        ConfigStoreFolder = fetchSetting("ConfigStoreFolder", vbNullString)
        specialConfigStoreFolders = Split(fetchSetting("specialConfigStoreFolders", vbNullString), ":")
        sortConfigStoreFolders = fetchSetting("sortConfigStoreFolders", True)
        DefaultDBDateFormatting = CInt(fetchSetting("DefaultDBDateFormatting", "0"))
    End Sub

    ''' <summary>Logs sErrMsg of eEventType to Logfile</summary>
    ''' <param name="sErrMsg"></param>
    ''' <param name="eEventType"></param>
    ''' <returns></returns>
    Public Function LogToEventViewer(sErrMsg As String, eEventType As EventLogEntryType) As Boolean
        Try
            logfile.WriteLine(Now().ToString() & vbTab & IIf(eEventType = EventLogEntryType.Error, "ERROR:", IIf(eEventType = EventLogEntryType.Information, "INFO:", "WARNING:")) & vbTab & sErrMsg)
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
        If Not automatedMapper Is Nothing Then automatedMapper.returnedErrorMessages = automatedMapper.returnedErrorMessages & LogMessage & vbCrLf
        If includeMsg And automatedMapper Is Nothing Then retval = MsgBox(LogMessage, vbCritical + IIf(exitMe, vbOKCancel, vbOKOnly), "DBAddin: Internal Error !! ")
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
        If Not automatedMapper Is Nothing Then automatedMapper.returnedErrorMessages = automatedMapper.returnedErrorMessages & LogMessage & vbCrLf
        If includeMsg And automatedMapper Is Nothing Then retval = MsgBox(LogMessage, vbCritical + IIf(exitMe, vbOKCancel, vbOKOnly), "DBAddin Error")
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
        If DEBUGME Then LogToEventViewer(LogMessage, EventLogEntryType.Information)
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    Public dontCalcWhileClearing As Boolean

    ' general Global objects/variables
    ''' <summary>Application object used for referencing objects</summary>
    Public theHostApp As Object
    ''' <summary>context menu handler</summary>
    Public theMenuHandler As MenuHandler
    ''' <summary>for interrupting long running operations with Ctl-Break</summary>
    Public Interrupted As Boolean
    ''' <summary>if we're running in the IDE, show errmsgs in immediate window and stop:resume on uncaught errors..</summary>
    Public VBDEBUG As Boolean
    ''' <summary>the environment (for Mapper special cases "Test", "Development" or vbNullString (prod))</summary>
    Public env As String
    ''' <summary>reference object for the Addins ribbon</summary>
    Public theRibbon As ExcelDna.Integration.CustomUI.IRibbonUI

    ' Global settings
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

    ' Global objects/variables for MenuHandler
    Public Const gsBUILDDB_TAG = "DBaddinBn1"
    Public Const gsREFRESH_TAG = "DBaddinBn2"
    Public Const gsITEMLOADCONFIG_TAG = "DBaddinBn3"
    Public Const gsITEMSAVECONFIG_TAG = "DBaddinBn4"
    Public Const gsJUMP_TAG = "DBaddinBn6"
    Public Const gsABOUT_TAG = "DBaddinAbout"
    Public Const gsInsertB_TAG = "DBaddinInsertB"
    Public Const gsDeleteB_TAG = "DBaddinDeleteB"
    Public Const gsForeignB_TAG = "DBaddinForeignB"
    Public Const gsDBConfigB_TAG = "DBaddinDBConfigB"
    Public Const gsDBConfigRefreshB_TAG = "DBaddinDBConfigRefreshB"
    Public Const gsITEMLOADPREPARED_TAG = "DBaddinLOADPREPARED"
    Public Const gsCONSTCONN_TAG = "DBaddinDBConstConnB"
    Public Const gsCONSTCONNACTION_TAG = "DBaddinDBConstConnActionB"

    Public Const gsDBSheetDefinitionB_TAG = "DBaddinDBSheetDefinitionB"
    Public Const gsDBSheetParametersB_TAG = "DBaddinDBSheetParametersB"
    Public Const gsDBSheetAssignB_TAG = "DBaddinDBSheetAssignB"
    Public Const gsDBSheetUnlockB_TAG = "DBaddinDBSheetUnlockB"
    Public Const gsDBSheetAutoRefreshB_TAG = "DBaddinDBSheetAutoRefreshB"

    Public Const gsDBSheetMain_TAG = "DBSheetMain"

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

''' <summary>Events from Excel (Workbook_Save ...)</summary>
Public Class AddInEvents
    Implements IExcelAddIn

    WithEvents Application As Application

    ''' <summary>connect to Excel when opening Addin</summary>
    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        ExcelRegistration.GetExcelFunctions().ProcessParamsRegistrations().RegisterFunctions()
        theHostApp = ExcelDnaUtil.Application
        Try
            MkDir("C:\temp")
            logfile = New StreamWriter("C:\\temp\\DBAddin.log", False, System.Text.Encoding.GetEncoding(1252))
        Catch ex As Exception
            MsgBox("Exception occured when trying to create logfile C:\temp\DBAddin.log: " + ex.Message)
        End Try
        logfile.WriteLine("starting DBAddin")
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

    Private Sub Workbook_Save(Wb As Workbook, ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeSave
    End Sub

    Private Sub Workbook_Open(Wb As Workbook) Handles Application.WorkbookOpen
        ' is being treated in Workbook_Activate...
    End Sub

    Private Sub Workbook_Activate(Wb As Workbook) Handles Application.WorkbookActivate
        Globals.theRibbon.Invalidate()
    End Sub

End Class

''' <summary>Events from Ribbon</summary>
<ComVisible(True)>
Public Class MyRibbon
    Inherits ExcelRibbon

    Public Sub ribbonLoaded(myribbon As ExcelDna.Integration.CustomUI.IRibbonUI)
        Globals.theRibbon = myribbon
    End Sub

End Class