Imports ExcelDna.Registration
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core
Imports System.Timers
Imports System.Diagnostics
Imports System.Runtime.InteropServices


''' <summary>AddIn Connection class, also handling Events from Excel (Open, Close, Activate)</summary>
<ComVisible(True)>
Public Class AddInEvents
    Implements IExcelAddIn

    ''' <summary>the app object needed for excel event handling (most of this class is dedicated to that)</summary>
    WithEvents Application As Excel.Application
    ''' <summary></summary>
    WithEvents ContextButton As CommandBarButton
    WithEvents cb As Microsoft.Vbe.Interop.Forms.CommandButton
    Private cbname As String

    ''' <summary>necessary to asynchronously start refresh of db functions after save event</summary>
    Private aTimer As System.Timers.Timer

    ''' <summary>connect to Excel when opening Addin</summary>
    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        ' necessary for ExplicitRegistration of param arrays (https://groups.google.com/forum/#!topic/exceldna/kf76nqAqDUo)
        ExcelRegistration.GetExcelFunctions().ProcessParamsRegistrations().RegisterFunctions()
        ExcelRegistration.GetExcelCommands().RegisterCommands()
        Application = ExcelDnaUtil.Application
        hostApp = ExcelDnaUtil.Application
        Try
            If hostApp.AddIns("DBAddin.Functions").Installed Then
                MsgBox("Attention: legacy DBAddin (DBAddin.Functions) still active, this might lead to unexpected results!")
            End If
        Catch ex As Exception : End Try
        Trace.Listeners.Add(New ExcelDna.Logging.LogDisplayTraceListener())
        ExcelDna.IntelliSense.IntelliSenseServer.Install()
        theMenuHandler = New MenuHandler
        LogInfo("initialize configuration settings")
        initSettings()
        Dim srchdListener As Object
        For Each srchdListener In Trace.Listeners
            If srchdListener.ToString() = "ExcelDna.Logging.LogDisplayTraceListener" Then
                Globals.theLogListener = srchdListener
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
            LogError("DBAddin unloading error: " + ex.Message)
        End Try
    End Sub

    ''' <summary>Workbook_Save: saves defined DBMaps (depending on configuration), also used to remove contents of DBListfunction results (data safety/space consumption)
    ''' choosing functions for removal of target data is done with custom docproperties</summary>
    Private Sub Application_WorkbookSave(Wb As Excel.Workbook, ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeSave
        Dim doRefreshDBFuncsAfterSave As Boolean = True
        Dim docproperty
        Dim DBFCContentColl As Collection, DBFCAllColl As Collection
        Dim theFunc
        Dim ws As Worksheet, lastWs As Worksheet = Nothing
        Dim searchCell As Range
        Dim firstAddress As String

        ' save all DBmaps/DBActions/DBSequences on saving except Readonly is recommended on this workbook
        Dim DBmapSheet As String
        If Not Wb.ReadOnlyRecommended Then
            For Each DBmapSheet In DBModifDefColl.Keys
                For Each dbmapdefkey In DBModifDefColl(DBmapSheet).Keys
                    If Left(dbmapdefkey, 8) = "DBSeqnce" Then
                        ' DB sequence actions (the sequence to be done) are stored directly in DBMapperDefColl, so different invocation here
                        doDBSeqnce(dbmapdefkey, DBModifDefColl(DBmapSheet).Item(dbmapdefkey), WbIsSaving:=True)
                    Else
                        Dim rngName As String = getDBModifNameFromRange(DBModifDefColl(DBmapSheet).Item(dbmapdefkey))
                        If Left(rngName, 8) = "DBMapper" Then
                            doDBMapper(DBModifDefColl(DBmapSheet).Item(dbmapdefkey), WbIsSaving:=True)
                        ElseIf Left(rngName, 8) = "DBAction" Then
                            doDBAction(DBModifDefColl(DBmapSheet).Item(dbmapdefkey), WbIsSaving:=True)
                        End If
                    End If
                Next
            Next
        End If
        DBFCContentColl = New Collection
        DBFCAllColl = New Collection
        Try
            For Each docproperty In Wb.CustomDocumentProperties
                If TypeName(docproperty.Value) = "Boolean" Then
                    If Left$(docproperty.Name, 5) = "DBFCC" And docproperty.Value Then DBFCContentColl.Add(True, Mid$(docproperty.Name, 6))
                    If Left$(docproperty.Name, 5) = "DBFCA" And docproperty.Value Then DBFCAllColl.Add(True, Mid$(docproperty.Name, 6))
                    If docproperty.Name = "DBFskip" Then doRefreshDBFuncsAfterSave = Not docproperty.Value
                End If
            Next
        Catch ex As Exception
            LogError("Error getting docproperties: " & Wb.Name & ex.Message)
        End Try
        dontCalcWhileClearing = True
        Try
            For Each ws In Wb.Worksheets
                If IsNothing(ws) Then
                    LogWarn("no worksheet in saving workbook...")
                    Exit For
                End If
                For Each theFunc In {"DBListFetch(", "DBRowFetch("}
                    searchCell = ws.Cells.Find(What:=theFunc, After:=ws.Range("A1"), LookIn:=XlFindLookIn.xlFormulas, LookAt:=XlLookAt.xlPart, SearchOrder:=XlSearchOrder.xlByRows, SearchDirection:=XlSearchDirection.xlNext, MatchCase:=False)
                    If Not (searchCell Is Nothing) Then
                        firstAddress = searchCell.Address
                        Do
                            ' get DB function target names from source names
                            Dim targetName As String = getDBunderlyingNameFromRange(searchCell).Name
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
            LogError("Error clearing DBfunction targets: " & Wb.Name & ex.Message)
        End Try
        dontCalcWhileClearing = False
    End Sub

    ''' <summary>open workbook: reset query cache, refresh DB functions and repair legacy functions if existing</summary>
    ''' <param name="Wb"></param>
    Private Sub Application_WorkbookOpen(Wb As Excel.Workbook) Handles Application.WorkbookOpen
        If Not Wb.IsAddin Then
            ' reset query cache !
            queryCache = New Collection
            StatusCollection = New Collection
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

    ''' <summary>WorkbookActivate: gets defined named ranges for DBMapper invocation in the current workbook after activation and updates Ribbon with it</summary>
    Private Sub Application_WorkbookActivate(Wb As Excel.Workbook) Handles Application.WorkbookActivate
        getDBModifDefinitions()
        ' unfortunately, Excel doesn't fire SheetActivate when opening workbooks, so do that here...
        assignHandler(Wb.ActiveSheet)
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
    Private Sub cb_Click() Handles cb.Click
        If Left(cbname, 8) = "DBMapper" Then
            doDBMapper(hostApp.ActiveWorkbook.Names.Item(cbname).RefersToRange)
        ElseIf Left(cbname, 8) = "DBAction" Then
            doDBAction(hostApp.ActiveWorkbook.Names.Item(cbname).RefersToRange)
        ElseIf Left(cbname, 8) = "DBSeqnce" Then
            Dim dbseqname As String = IIf(cbname = "DBSeqnce", "UnnamedDBSeqnce", Replace(cbname, "DBSeqnce", ""))
            doDBSeqnce(cbname, DBModifDefColl("ID0").Item(dbseqname))
        End If
    End Sub

    Sub assignHandler(Sh As Object)
        Dim foundDBModif As Boolean = False
        For Each shp As Excel.Shape In Sh.Shapes
            ' Associate clickhandler with all click events of the CommandButtons.
            Dim ctrlName As String = Sh.OLEObjects(shp.Name).Object.Name
            If Left(ctrlName, 8) = "DBMapper" Or Left(ctrlName, 8) = "DBAction" Or Left(ctrlName, 8) = "DBSeqnce" Then
                If foundDBModif Then
                    MsgBox("only one DBModifier Button allowed on a Worksheet, currently using " & cbname & " !")
                    Exit For
                End If
                cb = Sh.OLEObjects(shp.Name).Object
                cbname = ctrlName
                foundDBModif = True
            End If
        Next
    End Sub
    Private Sub Application_SheetActivate(Sh As Object) Handles Application.SheetActivate
        assignHandler(Sh)
    End Sub

    ''' <summary>SheetDeactivate: gets defined named ranges for DBMapper invocation after sheet was deleted/added (changes index of sheets-> IDs!) and updates Ribbon with it</summary>
    Private Sub Application_SheetDeactivate(Sh As Object) Handles Application.SheetDeactivate
        getDBModifDefinitions()
    End Sub

    Private Sub Application_SheetBeforeRightClick(Sh As Object, Target As Range, ByRef Cancel As Boolean) Handles Application.SheetBeforeRightClick
        Try : ContextButton.Delete() : Catch ex As Exception : End Try
        Dim dbModifName As String = getDBModifNameFromRange(hostApp.ActiveCell)
        If dbModifName <> "" Then
            Try
                ContextButton = hostApp.CommandBars("Cell").Controls.Add(Type:=MsoControlType.msoControlButton, Before:=1, Temporary:=True)
                ContextButton.Caption = "do " & Left(dbModifName, 8) & " (Ctrl-Shift-Alt to edit)"
                ContextButton.Tag = Left(dbModifName, 8)
                ContextButton.FaceId = 582
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub ContextButton_Click(Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles ContextButton.Click
        If My.Computer.Keyboard.CtrlKeyDown And My.Computer.Keyboard.ShiftKeyDown And My.Computer.Keyboard.AltKeyDown Then
            createDBModif(Ctrl.Tag, targetRange:=hostApp.ActiveCell)
        Else
            If Ctrl.Tag = "DBAction" Then
                doDBAction(hostApp.ActiveCell)
            ElseIf Ctrl.Tag = "DBMapper" Then
                doDBMapper(hostApp.ActiveCell)
            End If
        End If
    End Sub

End Class
