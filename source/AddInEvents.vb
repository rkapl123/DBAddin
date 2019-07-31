Imports ExcelDna.Registration
Imports Microsoft.Office.Interop
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop.Excel
Imports System.Timers

''' <summary>AddIn Connection class, also handling Events from Excel (Open, Close, Activate)</summary>
Public Class AddInEvents
    Implements IExcelAddIn

    ''' <summary>the app object needed for excel event handling (most of this class is dedicated to that)</summary>
    WithEvents Application As Excel.Application
    ''' <summary>necessary to asynchronously start refresh of db functions after save event</summary>
    Private aTimer As System.Timers.Timer

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
            ExcelDna.IntelliSense.IntelliSenseServer.Uninstall()
        Catch ex As Exception
            WriteToLog("DBAddin unloading error: " + ex.Message, EventLogEntryType.Warning)
        End Try
    End Sub

    ''' <summary>Workbook_Save: saves defined DBMaps (depending on configuration), also used to remove contents of DBListfunction results (data safety/space consumption)
    ''' choosing functions for removal of target data is done with custom docproperties</summary>
    Private Sub Workbook_Save(Wb As Excel.Workbook, ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeSave
        Dim doRefreshDBFuncsAfterSave As Boolean = True
        Dim docproperty
        Dim DBFCContentColl As Collection, DBFCAllColl As Collection
        Dim theFunc
        Dim ws As Worksheet, lastWs As Worksheet = Nothing
        Dim searchCell As Range
        Dim firstAddress As String

        Try
            saveDBMaps(Wb)
            DBFCContentColl = New Collection
            DBFCAllColl = New Collection
            For Each docproperty In Wb.CustomDocumentProperties
                If TypeName(docproperty.Value) = "Boolean" Then
                    If Left$(docproperty.Name, 5) = "DBFCC" And docproperty.Value Then DBFCContentColl.Add(True, Mid$(docproperty.Name, 6))
                    If Left$(docproperty.Name, 5) = "DBFCA" And docproperty.Value Then DBFCAllColl.Add(True, Mid$(docproperty.Name, 6))
                    If docproperty.Name = "DBFskip" Then doRefreshDBFuncsAfterSave = Not docproperty.Value
                End If
            Next
            dontCalcWhileClearing = True
            For Each ws In Wb.Worksheets
                For Each theFunc In {"DBListFetch(", "DBRowFetch("}
                    searchCell = ws.Cells.Find(What:=theFunc, After:=ws.Range("A1"), LookIn:=XlFindLookIn.xlFormulas, LookAt:=XlLookAt.xlPart, SearchOrder:=XlSearchOrder.xlByRows, SearchDirection:=XlSearchDirection.xlNext, MatchCase:=False)
                    If Not (searchCell Is Nothing) Then
                        firstAddress = searchCell.Address
                        Do
                            ' get DB function target names from source names
                            Dim targetName As String = getDBRangeName(searchCell).Name
                            targetName = Replace(targetName, "DBFsource", "DBFtarget", 1, , vbTextCompare)
                            ' check which DB functions should be content cleared (CC) or all cleared (CA)
                            Dim DBFCC As Boolean = False : Dim DBFCA As Boolean = False
                            DBFCC = DBFCContentColl.Contains("*")
                            DBFCC = DBFCContentColl.Contains(searchCell.Parent.Name & "!" & Replace(searchCell.Address, "$", String.Empty)) Or DBFCC
                            DBFCA = DBFCAllColl.Contains("*")
                            DBFCA = DBFCAllColl.Contains(searchCell.Parent.Name & "!" & Replace(searchCell.Address, "$", String.Empty)) Or DBFCA
                            Dim theTargetRange As Range = hostApp.Range(targetName)
                            If DBFCC Then
                                theTargetRange.Parent.Range(theTargetRange.Parent.Cells(theTargetRange.Row, theTargetRange.Column), theTargetRange.Parent.Cells(theTargetRange.Row + theTargetRange.Rows.Count - 1, theTargetRange.Column + theTargetRange.Columns.Count - 1)).ClearContents
                                LogInfo("App_WorkbookSave/Contents of selected DB Functions targets cleared")
                            End If
                            If DBFCA Then
                                theTargetRange.Parent.Range(theTargetRange.Parent.Cells(theTargetRange.Row + 2, theTargetRange.Column), theTargetRange.Parent.Cells(theTargetRange.Row + theTargetRange.Rows.Count - 1, theTargetRange.Column + theTargetRange.Columns.Count - 1)).Clear
                                theTargetRange.Parent.Range(theTargetRange.Parent.Cells(theTargetRange.Row, theTargetRange.Column), theTargetRange.Parent.Cells(theTargetRange.Row + 2, theTargetRange.Column + theTargetRange.Columns.Count - 1)).ClearContents
                                LogInfo("App_WorkbookSave/All cleared from selected DB Functions targets")
                            End If
                            searchCell = ws.Cells.FindNext(searchCell)
                        Loop While Not searchCell Is Nothing And searchCell.Address <> firstAddress
                    End If
                Next
                lastWs = ws
            Next
            ' reset the cell find dialog....
            searchCell = Nothing
            searchCell = lastWs.Cells.Find(What:="", After:=lastWs.Range("A1"), LookIn:=XlFindLookIn.xlFormulas, LookAt:=XlLookAt.xlPart, SearchOrder:=XlSearchOrder.xlByRows, SearchDirection:=XlSearchDirection.xlNext, MatchCase:=False)
            lastWs = Nothing
            ' refresh after save event
            If doRefreshDBFuncsAfterSave And (DBFCContentColl.Count > 0 Or DBFCAllColl.Count > 0) Then
                aTimer = New Timers.Timer(100)
                AddHandler aTimer.Elapsed, New ElapsedEventHandler(AddressOf refreshDBFuncLater)
                aTimer.Enabled = True
            End If
        Catch ex As Exception
            WriteToLog("Error: " & Wb.Name & ex.Message, EventLogEntryType.Warning)
        End Try
        dontCalcWhileClearing = False
    End Sub

    Private Sub Workbook_Open(Wb As Excel.Workbook) Handles Application.WorkbookOpen
        If Not Wb.IsAddin Then
            Dim refreshDBFuncs As Boolean
            ' when opening, force recalculation of DB functions in workbook.
            ' this is required as there is no recalculation if no dependencies have changed (usually when opening workbooks)
            ' however the most important dependency for DB functions is the database data....
            Try
                refreshDBFuncs = Not Wb.CustomDocumentProperties("DBFskip")
            Catch ex As Exception
                refreshDBFuncs = True
            End Try
            If refreshDBFuncs Then refreshDBFunctions(Wb)
            repairLegacyFunctions()
        End If
    End Sub

    ''' <summary>Workbook_Activate: gets defined named ranges for DBMapper invocation in the current workbook and updates Ribbon with it</summary>
    Private Sub Workbook_Activate(Wb As Excel.Workbook) Handles Application.WorkbookActivate
        getDBMapperDefinitions()
    End Sub

    ''' <summary>"OnTime" event function to "escape" workbook_save: event procedure to refetch DB functions results after saving</summary>
    ''' <param name="sender">the sending object (ourselves)</param>
    ''' <param name="e">Data for the Timer.Elapsed event</param>
    Shared Sub refreshDBFuncLater(ByVal sender As Object, ByVal e As ElapsedEventArgs)
        Dim previouslySaved As Boolean

        If Not hostApp.ActiveWorkbook Is Nothing Then
            previouslySaved = hostApp.ActiveWorkbook.Saved
            refreshDBFunctions(hostApp.ActiveWorkbook, True)
            hostApp.ActiveWorkbook.Saved = previouslySaved
        End If
    End Sub
End Class
