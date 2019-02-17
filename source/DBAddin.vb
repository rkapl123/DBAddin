Imports Microsoft.Office.Interop.Excel
Imports ExcelDna.Integration
Imports System.Runtime.InteropServices
Imports ExcelDna.Integration.CustomUI

' Global types
''
' type of DBFunc created with DBFuncBuilder insertForm
' @param DBList = 0
' @param DBRow = 1
' @param DBCell = 2
' @param DBCtrl = 3
Public Enum tDBFunc
    DBList = 0
    DBRow = 1
    DBCell = 2
    DBCtrl = 3
End Enum
''
' type for failure in reading DBsheet data
' @param dbDataError = 1
' @param dbOK = 2
' @param dbInternalError = 3
Public Enum dbsheetErrReason
    dbDataError = 1
    dbOK = 2
    dbInternalError = 3
End Enum

''
' All global Variables for DBFuncBuilder and Functions and some global accessible functions
Module Globals
    Public DEBUGME As Boolean
    Public myEventlog As EventLog
    Public AutomationMode As Boolean
    ' in case of Automation, use this to communicate collected logged Warnings and Errors back
    Public automatedMapper As Mapper

    ''
    ' encapsulates setting fetching (currently registry)
    ' @param Key
    ' @param defaultValue
    ' @param DBSheetSetting
    Public Function fetchSetting(Key As String, defaultValue As Object, Optional DBSheetSetting As Boolean = False) As Object
        fetchSetting = GetSetting("DBAddin", IIf(DBSheetSetting, "DBSheetSettings", "Settings"), Key, defaultValue)
    End Function

    ''
    ' encapsulates setting storing (currently registry)
    ' @param Key
    ' @param Value
    ' @param DBSheetSetting
    Public Sub storeSetting(Key As String, Value As Object, Optional DBSheetSetting As Boolean = False)
        SaveSetting("DBAddin", IIf(DBSheetSetting, "DBSheetSettings", "Settings"), Key, Value)
    End Sub

    ''
    '  initializes global configuration variables from registry
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

    ''
    ' Logs sErrMsg of eEventType in eCategory to EventLog
    ' @param sErrMsg As String
    ' @param eEventType
    ' @param eCategory
    Public Function LogToEventViewer(sErrMsg As String, eEventType As EventLogEntryType, eCategory As Short) As Boolean
        Try
            myEventlog.WriteEntry(sErrMsg, eEventType, 0, eCategory)
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    Public Sub LogError(LogMessage As String, Optional includeMsg As Boolean = True, Optional ByRef exitMe As Boolean = False, Optional category As Long = 2)
        Dim retval As Integer

        LogToEventViewer(LogMessage, EventLogEntryType.Error, category)
        If Not automatedMapper Is Nothing Then automatedMapper.returnedErrorMessages = automatedMapper.returnedErrorMessages & LogMessage & vbCrLf
        If includeMsg And Not AutomationMode Then retval = MsgBox(LogMessage, vbCritical + IIf(exitMe, vbOKCancel, vbOKOnly), "DBAddin: Internal Error !! ")
        If retval = vbCancel Then
            exitMe = True
        Else
            exitMe = False
        End If
    End Sub

    Public Sub LogWarn(LogMessage As String, Optional ByRef exitMe As Boolean = False, Optional category As Long = 2, Optional includeMsg As Boolean = True)
        Dim retval As Integer

        LogToEventViewer(LogMessage, EventLogEntryType.Warning, category)
        If Not automatedMapper Is Nothing Then automatedMapper.returnedErrorMessages = automatedMapper.returnedErrorMessages & LogMessage & vbCrLf
        If includeMsg And Not AutomationMode Then retval = MsgBox(LogMessage, vbCritical + IIf(exitMe, vbOKCancel, vbOKOnly), "DBAddin Error")
        If retval = vbCancel Then
            exitMe = True
        Else
            exitMe = False
        End If
    End Sub

    Public Sub LogInfo(LogMessage As String, Optional category As Long = 2)
        If DEBUGME Then LogToEventViewer(LogMessage, EventLogEntryType.Information, category)
    End Sub

    Public dontCalcWhileClearing As Boolean

    ' Global objects/variables for all
    Public theHostApp As Object
    Public theMenuHandler As MenuHandler

    ''
    ' for interrupting long running operations with Ctl-Break
    Public Interrupted As Boolean
    ''
    ' if we're running in the IDE, show errmsgs in immediate window and stop:resume on uncaught errors..
    Public VBDEBUG As Boolean

    ' Global configuration variables
    ''
    ' Default ConnectionString, if no connection string is given by user....
    Public ConstConnString As String
    ''
    ' the tag used to identify the Database name within the ConstConnString
    Public DBidentifierCCS As String
    ''
    ' the tag used to identify the Database name within the connection string returned by MSQuery
    Public DBidentifierODBC As String
    ''
    ' the folder used to store predefined DB item definitions
    Public ConfigStoreFolder As String
    Public specialConfigStoreFolders() As String
    Public sortConfigStoreFolders As Boolean
    Public DBSheetDefinitionsFolder As String

    Public DBConnFileName As String

    ''
    ' global connection timeout (can't be set in DB functions)
    Public CnnTimeout As Integer

    ' global command timeout (can't be set in DB functions)
    Public CmdTimeout As Integer

    ' Global flags
    ''
    ' prevent multiple connection retries for each function in case of error
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

    ''
    ' store target filter in case of empty data lists
    Public targetFilterCont As Collection
    ''
    ' global event class, mainly for calc event procedure
    Public theDBFuncEventHandler As DBFuncEventHandler

    ''
    ' global collection of information transport containers between function and calc event procedure
    Public allCalcContainers As Collection
    ''
    ' global collection of information transport containers between function and calc event procedure
    Public allStatusContainers As Collection

    ''
    ' this is prepended before columns that may not be null
    Public specialNonNullableChar As String
    ''
    ' special placeholder for being replaced in lookups by the foreign table of that row (T2, T3...)
    Public tblPlaceHolder As String
    Public closeMsg1 As String
    Public closeMsg2 As String
    Public DefaultDBDateFormatting As Integer

    ''
    ' the environment (for Mapper special cases "Test", "Development" or vbNullString (prod))
    Public env As String
    ''
    ' main db connection for dbsheets
    Public dbcnn As ADODB.Connection
    ''
    ' Helps to prevent unwanted events triggering actions
    Public noDBSheetEvent As Boolean
    ''
    ' global password collection (are set once for each connection/database combination)
    Public passwords As Collection

    Public theRibbon As ExcelDna.Integration.CustomUI.IRibbonUI
End Module

' Events from Excel (Workbook_Save ...)
Public Class AddIn
    Implements IExcelAddIn

    WithEvents Application As Application

    Public Sub DisconnectDBAddin()
        On Error Resume Next
        theMenuHandler = Nothing
        theHostApp = Nothing
        theDBFuncEventHandler = Nothing
    End Sub

    ' connect to Excel when opening Addin
    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        theHostApp = ExcelDnaUtil.Application
        myEventlog = New EventLog("Application")
        initSettings()
        theMenuHandler = New MenuHandler
        theDBFuncEventHandler = New DBFuncEventHandler
    End Sub

    'has to be implemented
    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
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

' Events from Ribbon
<ComVisible(True)>
Public Class MyRibbon
    Inherits ExcelRibbon

    Public Sub ribbonLoaded(myribbon As ExcelDna.Integration.CustomUI.IRibbonUI)
        Globals.theRibbon = myribbon
    End Sub

End Class