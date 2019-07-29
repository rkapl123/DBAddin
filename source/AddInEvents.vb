Imports ExcelDna.Registration
Imports Microsoft.Office.Interop
Imports ExcelDna.Integration

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
        Trace.Listeners.Add(New ExcelDna.Logging.LogDisplayTraceListener())
        ExcelDna.IntelliSense.IntelliSenseServer.Install()
        theMenuHandler = New MenuHandler
        theDBFuncEventHandler = New DBFuncEventHandler
        WriteToLog("initialize configuration settings", EventLogEntryType.Information)
        initSettings()
        Dim srchdListener As Object
        For Each srchdListener In Trace.Listeners
            If srchdListener.ToString() = "ExcelDna.Logging.LogDisplayTraceListener" Then
                DBAddin.theLogListener = srchdListener
                Exit For
            End If
        Next
    End Sub

    ''' <summary>AutoClose cleans up after finishing addin</summary>
    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        Try
            theMenuHandler = Nothing
            hostApp = Nothing
            theDBFuncEventHandler = Nothing
            ExcelDna.IntelliSense.IntelliSenseServer.Uninstall()
        Catch ex As Exception
            WriteToLog("DBAddin unloading error: " + ex.Message, EventLogEntryType.Warning)
        End Try
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
