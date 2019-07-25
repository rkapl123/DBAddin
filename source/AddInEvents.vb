Imports ExcelDna.Registration
Imports System.IO ' needed for logfile
Imports Microsoft.Office.Interop
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop.Excel

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
        hostApp = ExcelDnaUtil.Application

        Dim logfilename As String = "C:\\DBAddinlogs\\" + DateTime.Today.ToString("yyyyMMdd") + ".log"
        If Not Directory.Exists("C:\\DBAddinlogs") Then MkDir("C:\\DBAddinlogs")
        Try
            logfile = New StreamWriter(logfilename, True, System.Text.Encoding.GetEncoding(1252))
            logfile.AutoFlush = True
        Catch ex As Exception
            MsgBox("Exception occured when trying to create logfile " + logfilename + ": " + ex.Message)
        End Try
        WriteToLog("starting DBAddin", EventLogEntryType.Information)
        theMenuHandler = New MenuHandler
        theDBFuncEventHandler = New DBFuncEventHandler
        initSettings()
    End Sub

    ''' <summary>AutoClose cleans up after finishing addin</summary>
    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        On Error Resume Next
        theMenuHandler = Nothing
        hostApp = Nothing
        theDBFuncEventHandler = Nothing
    End Sub

    ''' <summary>Workbook_Save: saves defined DBMaps (depending on configuration)</summary>
    Private Sub Workbook_Save(Wb As Excel.Workbook, ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeSave
        saveDBMaps(Wb)
    End Sub

    Private Sub Workbook_Open(Wb As Excel.Workbook) Handles Application.WorkbookOpen
        ' ribbon invalidation (refreshing) is being treated in WorkbookActivate...
    End Sub

    ''' <summary>Workbook_Activate: gets defined named ranges for DBMapper invocation in the current workbook and updates Ribbon with it</summary>
    Private Sub Workbook_Activate(Wb As Excel.Workbook) Handles Application.WorkbookActivate
        getDBMapperDefinitions()
    End Sub
End Class
