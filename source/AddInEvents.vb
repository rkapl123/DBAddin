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
    End Sub

    Private Sub Workbook_Save(Wb As Excel.Workbook, ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeSave
        ' save all DBmaps on saving except Readonly is recommended on this workbook
        Dim DBmapSheet As String
        If Not Wb.ReadOnlyRecommended Then
            For Each DBmapSheet In DBMapperDefColl.Keys
                For Each dbmapdefkey In DBMapperDefColl(DBmapSheet).Keys
                    saveRangeToDB(DBMapperDefColl(DBmapSheet).Item(dbmapdefkey), dbmapdefkey)
                Next
            Next
        End If
    End Sub

    Private Sub Workbook_Open(Wb As Excel.Workbook) Handles Application.WorkbookOpen
        ' ribbon invalidation (refreshing) is being treated in WorkbookActivate...
    End Sub

    ''' <summary>Workbook_Activate: gets defined named ranges for DBMapper invocation in the current workbook and updates Ribbon with it</summary>
    Private Sub Workbook_Activate(Wb As Excel.Workbook) Handles Application.WorkbookActivate
        ' load DBMapper definitions
        DBAddin.DBMapperDefColl = New Dictionary(Of String, Dictionary(Of String, Range))
        For Each namedrange As Name In hostApp.ActiveWorkbook.Names
            Dim cleanname As String = Replace(namedrange.Name, namedrange.Parent.Name & "!", "")
            If Left(cleanname, 8) = "DBMapper" Then
                If InStr(namedrange.RefersTo, "#REF!") > 0 Then LogError("DBMapper definitions range " + namedrange.Parent.name + "!" + namedrange.Name + " contains #REF!")
                Dim nodeName As String = Replace(Replace(namedrange.Name, "DBMapper", ""), namedrange.Parent.Name & "!", "")
                If nodeName = "" Then nodeName = "UnnamedDBMapper"

                Dim i As Integer = namedrange.RefersToRange.Parent.Index
                Dim defColl As Dictionary(Of String, Range)
                If Not DBMapperDefColl.ContainsKey("ID" + i.ToString()) Then
                    ' add to new sheet "menu"
                    defColl = New Dictionary(Of String, Range)
                    defColl.Add(nodeName, namedrange.RefersToRange)
                    DBMapperDefColl.Add("ID" + i.ToString(), defColl)
                Else
                    ' add definition to existing sheet "menu"
                    defColl = DBMapperDefColl("ID" + i.ToString())
                    defColl.Add(nodeName, namedrange.RefersToRange)
                End If
            End If
            If DBMapperDefColl.Count >= 15 Then LogError("Not more than 15 sheets with DBMapper definitions possible, ignoring definitions in sheet " + namedrange.Parent.Name)
        Next
        DBAddin.theRibbon.Invalidate()
    End Sub
End Class
