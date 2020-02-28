Imports ExcelDna.Integration
Imports ExcelDna.Registration
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core
Imports Microsoft.Vbe.Interop
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Collections.Generic


''' <summary>AddIn Connection class, also handling Events from Excel (Open, Close, Activate)</summary>
<ComVisible(True)>
Public Class AddInEvents
    Implements IExcelAddIn

    ''' <summary>the app object needed for excel event handling (most of this class is dedicated to that)</summary>
    WithEvents Application As Excel.Application
    ''' <summary>CommandButton that can be inserted on a worksheet (name property being the same as the respective target range (for DBMapper/DBAction) or DBSeqnce Name)</summary>
    Shared WithEvents cb1 As Forms.CommandButton
    ''' <summary>CommandButton that can be inserted on a worksheet (name property being the same as the respective target range (for DBMapper/DBAction) or DBSeqnce Name)</summary>
    Shared WithEvents cb2 As Forms.CommandButton
    ''' <summary>CommandButton that can be inserted on a worksheet (name property being the same as the respective target range (for DBMapper/DBAction) or DBSeqnce Name)</summary>
    Shared WithEvents cb3 As Forms.CommandButton
    ''' <summary>CommandButton that can be inserted on a worksheet (name property being the same as the respective target range (for DBMapper/DBAction) or DBSeqnce Name)</summary>
    Shared WithEvents cb4 As Forms.CommandButton
    ''' <summary>CommandButton that can be inserted on a worksheet (name property being the same as the respective target range (for DBMapper/DBAction) or DBSeqnce Name)</summary>
    Shared WithEvents cb5 As Forms.CommandButton
    ''' <summary>necessary to asynchronously start refresh of db functions after save event</summary>
    Private aTimer As System.Timers.Timer

    ''' <summary>connect to Excel when opening Addin</summary>
    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        ' necessary for ExplicitRegistration of param arrays (https://groups.google.com/forum/#!topic/exceldna/kf76nqAqDUo)
        ExcelRegistration.GetExcelFunctions().ProcessParamsRegistrations().RegisterFunctions()
        ExcelRegistration.GetExcelCommands().RegisterCommands()
        Application = ExcelDnaUtil.Application
        Try
            If ExcelDnaUtil.Application.AddIns("DBAddin.Functions").Installed Then
                MsgBox("Attention: legacy DBAddin (DBAddin.Functions) still active, this might lead to unexpected results!")
            End If
        Catch ex As Exception : End Try
        Trace.Listeners.Add(New ExcelDna.Logging.LogDisplayTraceListener())
        ExcelDna.IntelliSense.IntelliSenseServer.Install()
        theMenuHandler = New MenuHandler
        LogInfo("initialize configuration settings")
        queryCache = New Dictionary(Of String, String)
        StatusCollection = New Dictionary(Of String, ContainedStatusMsg)
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
            ExcelDna.IntelliSense.IntelliSenseServer.Uninstall()
        Catch ex As Exception
            LogError("DBAddin unloading error: " + ex.Message)
        End Try
    End Sub

    ''' <summary>Workbook_Save: saves defined DBMaps (depending on configuration), also used to remove contents of DBListfunction results (data safety/space consumption)
    ''' choosing functions for removal of target data is done with custom docproperties</summary>
    Private Sub Application_WorkbookSave(Wb As Excel.Workbook, ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeSave
        ' ask if modifications should be done if no overriding flag is defined...
        Dim doDBMOnSave As Boolean = False
        Try
            If Wb.CustomDocumentProperties("doDBMOnSave").Value Then doDBMOnSave = True
        Catch ex As Exception : End Try
        ' if overriding flag not given, ask for saving if this is necessary for any DBmodifier...
        If Not doDBMOnSave Then
            For Each DBmodifType As String In Globals.DBModifDefColl.Keys
                For Each dbmapdefkey As String In Globals.DBModifDefColl(DBmodifType).Keys
                    If Globals.DBModifDefColl(DBmodifType).Item(dbmapdefkey).DBModifSaveNeeded() Then
                        Dim answer As MsgBoxResult = MsgBox("do the DB Modifications defined in Workbook ?", vbYesNo, "DB Modifications on Save")
                        If answer = vbYes Then doDBMOnSave = True
                        GoTo done
                    End If
                Next
            Next
        End If
done:
        ' save all DBmaps/DBActions/DBSequences on saving except Readonly is recommended on this workbook
        If Not Wb.ReadOnlyRecommended And doDBMOnSave Then
            For Each DBmodifType As String In Globals.DBModifDefColl.Keys
                For Each dbmapdefkey As String In Globals.DBModifDefColl(DBmodifType).Keys
                    With Globals.DBModifDefColl(DBmodifType).Item(dbmapdefkey)
                        If .DBModifSaveNeeded Then .doDBModif(WbIsSaving:=True)
                    End With
                Next
            Next
        End If

        ' clear DB Functions content and refresh afterwards..
        Dim DBFCContentColl As Collection = New Collection
        Dim DBFCAllColl As Collection = New Collection
        Dim doRefreshDBFuncsAfterSave As Boolean = True
        ' first insert docproperty information into collections for easier handling
        Try
            Dim docproperty As DocumentProperty
            For Each docproperty In Wb.CustomDocumentProperties
                If docproperty.Type = MsoDocProperties.msoPropertyTypeBoolean Then
                    If Left$(docproperty.Name, 5) = "DBFCC" And docproperty.Value Then DBFCContentColl.Add(True, Mid(docproperty.Name, 6))
                    If Left$(docproperty.Name, 5) = "DBFCA" And docproperty.Value Then DBFCAllColl.Add(True, Mid(docproperty.Name, 6))
                    If docproperty.Name = "DBFskip" Then doRefreshDBFuncsAfterSave = Not docproperty.Value
                End If
            Next
        Catch ex As Exception
            LogError("Error getting docproperties: " & Wb.Name & ex.Message)
        End Try

        ' now clear content/all
        Dim searchCell As Excel.Range
        dontCalcWhileClearing = True
        Try
            Dim ws As Excel.Worksheet, lastWs As Excel.Worksheet = Nothing
            For Each ws In Wb.Worksheets
                If IsNothing(ws) Then
                    LogWarn("no worksheet in saving workbook...")
                    Exit For
                End If
                For Each theFunc As String In {"DBListFetch(", "DBRowFetch("}
                    searchCell = ws.Cells.Find(What:=theFunc, After:=ws.Range("A1"), LookIn:=Excel.XlFindLookIn.xlFormulas, LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlNext, MatchCase:=False)
                    If Not (searchCell Is Nothing) Then
                        Dim firstAddress As String = searchCell.Address
                        Do
                            ' get DB function target names from source names
                            Dim targetName As String = getDBunderlyingNameFromRange(searchCell)
                            ' in case of commented cells and purged db underlying names, getDBunderlyingNameFromRange doesn't return a name ...
                            If InStr(UCase(targetName), "DBFSOURCE") > 0 Then
                                targetName = Replace(targetName, "DBFsource", "DBFtarget", 1, , vbTextCompare)
                                ' check which DB functions should be content cleared (CC) or all cleared (CA)
                                Dim DBFCC As Boolean = False : Dim DBFCA As Boolean = False
                                DBFCC = DBFCContentColl.Contains("*")
                                DBFCC = DBFCContentColl.Contains(searchCell.Parent.Name & "!" & Replace(searchCell.Address, "$", String.Empty)) Or DBFCC
                                DBFCA = DBFCAllColl.Contains("*")
                                DBFCA = DBFCAllColl.Contains(searchCell.Parent.Name & "!" & Replace(searchCell.Address, "$", String.Empty)) Or DBFCA
                                Dim theTargetRange As Excel.Range
                                Try : theTargetRange = ExcelDnaUtil.Application.Range(targetName)
                                Catch ex As Exception
                                    MsgBox("Error in finding target range of DB Function " & theFunc & "in " & firstAddress & "), refreshing all DB functions should solve this.", MsgBoxStyle.Critical)
                                    searchCell = Nothing
                                    searchCell = lastWs.Cells.Find(What:="", After:=lastWs.Range("A1"), LookIn:=Excel.XlFindLookIn.xlFormulas, LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlNext, MatchCase:=False)
                                    dontCalcWhileClearing = False
                                    Exit Sub
                                End Try
                                If DBFCC Then
                                    theTargetRange.Parent.Range(theTargetRange.Parent.Cells(theTargetRange.Row, theTargetRange.Column), theTargetRange.Parent.Cells(theTargetRange.Row + theTargetRange.Rows.Count - 1, theTargetRange.Column + theTargetRange.Columns.Count - 1)).ClearContents
                                    LogInfo("Contents of selected DB Functions targets cleared")
                                End If
                                If DBFCA Then
                                    theTargetRange.Parent.Range(theTargetRange.Parent.Cells(theTargetRange.Row + 2, theTargetRange.Column), theTargetRange.Parent.Cells(theTargetRange.Row + theTargetRange.Rows.Count - 1, theTargetRange.Column + theTargetRange.Columns.Count - 1)).Clear
                                    theTargetRange.Parent.Range(theTargetRange.Parent.Cells(theTargetRange.Row, theTargetRange.Column), theTargetRange.Parent.Cells(theTargetRange.Row + 2, theTargetRange.Column + theTargetRange.Columns.Count - 1)).ClearContents
                                    LogInfo("All cleared from selected DB Functions targets")
                                End If
                            End If
                            searchCell = ws.Cells.FindNext(searchCell)
                        Loop While Not searchCell Is Nothing And searchCell.Address <> firstAddress
                    End If
                Next
                lastWs = ws
            Next
            ' always reset the cell find dialog....
            searchCell = Nothing
            searchCell = lastWs.Cells.Find(What:="", After:=lastWs.Range("A1"), LookIn:=Excel.XlFindLookIn.xlFormulas, LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlNext, MatchCase:=False)
            lastWs = Nothing

            ' refresh content area of dbfunctions after save event, requires execution out of context of Application_WorkbookSave
            If doRefreshDBFuncsAfterSave And (DBFCContentColl.Count > 0 Or DBFCAllColl.Count > 0) Then
                ExcelAsyncUtil.QueueAsMacro(Sub()
                                                Globals.refreshDBFuncLater()
                                            End Sub)
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
            Dim refreshDBFuncs As Boolean
            ' in case of reopening workbooks with dbfunctions, look for old query caches and status collections (returned error messages) and reset them to get new data
            resetCachesForWorkbook(Wb.Name)

            ' when opening, force recalculation of DB functions in workbook.
            ' this is required as there is no recalculation if no dependencies have changed (usually when opening workbooks)
            ' however the most important dependency for DB functions is the database data....
            Try
                refreshDBFuncs = Not Wb.CustomDocumentProperties("DBFskip").Value
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

    ''' <summary>specific click handlers for the five definable commandbuttons</summary>
    Private Shared Sub cb1_Click() Handles cb1.Click
        cbClick(cb1.Name)
    End Sub
    Private Shared Sub cb2_Click() Handles cb2.Click
        cbClick(cb2.Name)
    End Sub
    Private Shared Sub cb3_Click() Handles cb3.Click
        cbClick(cb3.Name)
    End Sub
    Private Shared Sub cb4_Click() Handles cb4.Click
        cbClick(cb4.Name)
    End Sub
    Private Shared Sub cb5_Click() Handles cb5.Click
        cbClick(cb5.Name)
    End Sub

    ''' <summary>common click handler for all commandbuttons</summary>
    ''' <param name="cbName">name of command button, defines whether a DBModification is invoked (starts with DBMapper/DBAction/DBSeqnce)</param>
    Private Shared Sub cbClick(cbName As String)
        Dim DBModifType As String = Left(cbName, 8)
        If DBModifType <> "DBSeqnce" Then
            Dim targetRange As Excel.Range
            Try
                targetRange = ExcelDnaUtil.Application.ActiveWorkbook.Names.Item(cbName).RefersToRange
            Catch ex As Exception
                MsgBox("No underlying " & Left(cbName, 8) & " Range named " & cbName & " found, exiting without DBModification.")
                LogWarn("targetRange assignment failed: " & ex.Message)
                Exit Sub
            End Try
            Dim DBModifName As String = getDBModifNameFromRange(targetRange)
        End If
        If My.Computer.Keyboard.CtrlKeyDown And My.Computer.Keyboard.ShiftKeyDown Then
            createDBModif(DBModifType, targetDefName:=cbName)
        Else
            Globals.DBModifDefColl(DBModifType).Item(cbName).doDBModif()
        End If
    End Sub

    ''' <summary>assign click handlers to commandbuttons in passed sheet Sh, maximum 5 buttons are supported</summary>
    ''' <param name="Sh"></param>
    Public Shared Function assignHandler(Sh As Object) As Boolean
        cb1 = Nothing : cb2 = Nothing : cb3 = Nothing : cb4 = Nothing : cb5 = Nothing
        assignHandler = True
        For Each shp As Excel.Shape In Sh.Shapes
            ' Associate clickhandler with all click events of the CommandButtons.
            Dim ctrlName As String
            Try : ctrlName = Sh.OLEObjects(shp.Name).Object.Name : Catch ex As Exception : ctrlName = "" : End Try
            If Left(ctrlName, 8) = "DBMapper" Or Left(ctrlName, 8) = "DBAction" Or Left(ctrlName, 8) = "DBSeqnce" Then
                If IsNothing(cb1) Then
                    cb1 = Sh.OLEObjects(shp.Name).Object
                ElseIf IsNothing(cb2) Then
                    cb2 = Sh.OLEObjects(shp.Name).Object
                ElseIf IsNothing(cb3) Then
                    cb3 = Sh.OLEObjects(shp.Name).Object
                ElseIf IsNothing(cb4) Then
                    cb4 = Sh.OLEObjects(shp.Name).Object
                ElseIf IsNothing(cb5) Then
                    cb5 = Sh.OLEObjects(shp.Name).Object
                Else
                    MsgBox("only max. of five DBModifier Buttons allowed on a Worksheet, currently using " & cb1.Name & "," & cb2.Name & "," & cb3.Name & "," & cb4.Name & " and " & cb5.Name & " !")
                    assignHandler = False
                    Exit For
                End If
            End If
        Next
    End Function

    ''' <summary>assign commandbuttons new and refresh DBAddins DBModification Menu with each change of sheets</summary>
    ''' <param name="Sh"></param>
    Private Sub Application_SheetActivate(Sh As Object) Handles Application.SheetActivate
        getDBModifDefinitions()
        ' only when needed assign button handler for this sheet ...
        If Globals.DBModifDefColl.Count > 0 Then assignHandler(Sh)
    End Sub

    Private Sub Application_WorkbookBeforeClose(Wb As Workbook, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeClose
        If Globals.DBModifDefColl.Count > 0 Then
            Globals.DBModifDefColl.Clear()
            Globals.theRibbon.Invalidate()
        End If
    End Sub

    Private Sub Application_SheetChange(Sh As Object, Target As Range) Handles Application.SheetChange
        If Globals.DBModifDefColl.ContainsKey("DBMapper") And Not DBModifs.preventChangeWhileFetching Then
            Dim targetName As String = getDBModifNameFromRange(Target)
            If Left(targetName, 8) = "DBMapper" Then DirectCast(Globals.DBModifDefColl("DBMapper").Item(targetName), DBMapper).doCUDMarks(Target)
        End If
    End Sub
End Class
