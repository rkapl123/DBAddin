Imports ExcelDna.Integration
Imports ExcelDna.Registration
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel ' for event procedures...
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

    ''' <summary>connect to Excel when opening Addin</summary>
    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        ' necessary for ExplicitRegistration of param arrays (https://groups.google.com/forum/#!topic/exceldna/kf76nqAqDUo)
        ExcelRegistration.GetExcelFunctions().ProcessParamsRegistrations().RegisterFunctions()
        ExcelRegistration.GetExcelCommands().RegisterCommands()
        Application = ExcelDnaUtil.Application
        Try
            If ExcelDnaUtil.Application.AddIns("DBAddin.Functions").Installed Then
                Globals.UserMsg("Attention: legacy DBAddin (DBAddin.Functions) still active, this might lead to unexpected results!")
            End If
        Catch ex As Exception : End Try
        ' for finding out what happened...
        Trace.Listeners.Add(New ExcelDna.Logging.LogDisplayTraceListener())
        ' IntelliSense needed for DB- and supporting functions
        ExcelDna.IntelliSense.IntelliSenseServer.Install()
        ' Ribbon and context menu setup
        Globals.theMenuHandler = New MenuHandler
        Globals.LogInfo("initialize configuration settings")
        Functions.queryCache = New Dictionary(Of String, String)
        Functions.StatusCollection = New Dictionary(Of String, ContainedStatusMsg)
        Globals.DBModifDefColl = New Dictionary(Of String, Dictionary(Of String, DBModif))
        ' get the ExcelDna LogDisplayTraceListener for filtering log messages by level in about box
        For Each srchdListener As Object In Trace.Listeners
            If srchdListener.ToString() = "ExcelDna.Logging.LogDisplayTraceListener" Then
                Globals.theLogListener = srchdListener
                Exit For
            End If
        Next
        ' initialize settings and get the default environment
        Globals.initSettings()
        ' Configs are 1 based, selectedEnvironment(index of environment dropdown) is 0 based. negative values not allowed!
        Dim selEnv As Integer = CInt(fetchSetting("DefaultEnvironment", "1")) - 1
        If selEnv > environdefs.Length - 1 OrElse selEnv < 0 Then
            Globals.UserMsg("Default Environment " + (selEnv + 1).ToString() + " not existing, setting to first environment !")
            selEnv = 0
        End If
        Globals.selectedEnvironment = selEnv
        ' after getting the default environment (should exist), set the Const Connection String again to avoid problems in generating DB ListObjects
        ConstConnString = fetchSetting("ConstConnString" + Globals.env(), "")
        ConfigStoreFolder = fetchSetting("ConfigStoreFolder" + Globals.env(), "")
        ' last, check if any updates are available on github...
        ExcelAsyncUtil.QueueAsMacro(Sub()
                                        Globals.checkForUpdate(False)
                                    End Sub)
    End Sub

    ''' <summary>AutoClose cleans up after finishing addin</summary>
    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        Try
            Globals.theMenuHandler = Nothing
            ExcelDna.IntelliSense.IntelliSenseServer.Uninstall()
        Catch ex As Exception
            Globals.UserMsg("DBAddin unloading error: " + ex.Message, "AutoClose")
        End Try
    End Sub

    ''' <summary>saves defined DBMaps (depending on configuration), also used to remove contents of DBListfunction results (data safety/space consumption)
    ''' choosing functions for removal of target data is done with custom docproperties</summary>
    ''' <param name="Wb"></param>
    ''' <param name="SaveAsUI"></param>
    ''' <param name="Cancel"></param>
    Private Sub Application_WorkbookSave(Wb As Excel.Workbook, ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeSave
        ' ask if modifications should be done if no overriding flag is defined...
        Dim doDBMOnSave As Boolean = Globals.getCustPropertyBool("doDBMOnSave", Wb)

        Dim askForEveryModifier As Boolean = False
        ' if overriding flag not given and Readonly is NOT recommended on this workbook and workbook IS NOT Readonly, ...
        If Not (Wb.ReadOnlyRecommended And Wb.ReadOnly) And Not doDBMOnSave Then
            Dim TotalDBModifCount As Integer = 0
            For Each DBmodifType As String In Globals.DBModifDefColl.Keys
                For Each dbmapdefkey As String In Globals.DBModifDefColl(DBmodifType).Keys
                    If Globals.DBModifDefColl(DBmodifType).Item(dbmapdefkey).execOnSave Then TotalDBModifCount += 1
                Next
            Next
            ' ...ask for saving, if this is necessary for any DBmodifier...
            If TotalDBModifCount > 1 Then
                ' multiple DBmodifiers, ask how to proceed
                Dim answer As MsgBoxResult = QuestionMsg(theMessage:="do all DB Modifications defined in Workbook (Yes=All with exec on Save, No=Ask everytime. Cancel=Don't do any DB Modifications) ?", questionType:=MsgBoxStyle.YesNoCancel, questionTitle:="Do DB Modifiers on Save")
                If answer = MsgBoxResult.Yes Then doDBMOnSave = True
                If answer = MsgBoxResult.Cancel Then doDBMOnSave = False
                If answer = MsgBoxResult.No Then
                    doDBMOnSave = True
                    askForEveryModifier = True
                End If
            ElseIf TotalDBModifCount = 1 Then
                ' only one DBModifier needs saving, ask only once for saving...
                For Each DBmodifType As String In Globals.DBModifDefColl.Keys
                    For Each dbmapdefkey As String In Globals.DBModifDefColl(DBmodifType).Keys
                        If Globals.DBModifDefColl(DBmodifType).Item(dbmapdefkey).execOnSave Then
                            If nonInteractive Then
                                ' always save in noninteractive (headless/automation) mode
                                doDBMOnSave = True
                            Else
                                doDBMOnSave = IIf(Globals.DBModifDefColl(DBmodifType).Item(dbmapdefkey).confirmExecution(WbIsSaving:=True) = MsgBoxResult.Yes, True, False)
                            End If
                        End If
                    Next
                Next
            End If
        End If

        ' save all DBmappers/DBActions/DBSequences on saving if above resulted in YES!
        If doDBMOnSave Then
            ' first do DBModifiers defined on active sheet or any DB Sequence:
            For Each DBmodifType As String In Globals.DBModifDefColl.Keys
                For Each dbmapdefkey As String In Globals.DBModifDefColl(DBmodifType).Keys
                    With Globals.DBModifDefColl(DBmodifType).Item(dbmapdefkey)
                        If (DBmodifType = "DBSeqnce" OrElse .getTargetRange().Parent Is ExcelDnaUtil.Application.ActiveSheet) And .DBModifSaveNeeded Then
                            ' ask for saving, if decided so...
                            If askForEveryModifier Then
                                Dim answer As MsgBoxResult = .confirmExecution(WbIsSaving:=True)
                                If answer = MsgBoxResult.Yes Then .doDBModif(WbIsSaving:=True)
                                If answer = MsgBoxResult.Cancel Then GoTo done
                            Else
                                .doDBModif(WbIsSaving:=True)
                            End If
                        End If
                    End With
                Next
            Next
            ' then all the rest (no defined order!)
            For Each DBmodifType As String In Globals.DBModifDefColl.Keys
                For Each dbmapdefkey As String In Globals.DBModifDefColl(DBmodifType).Keys
                    With Globals.DBModifDefColl(DBmodifType).Item(dbmapdefkey)
                        If Not (DBmodifType = "DBSeqnce" OrElse .getTargetRange().Parent Is ExcelDnaUtil.Application.ActiveSheet) And .DBModifSaveNeeded Then
                            If askForEveryModifier Then
                                Dim answer As MsgBoxResult = .confirmExecution(WbIsSaving:=True)
                                If answer = MsgBoxResult.Yes Then .doDBModif(WbIsSaving:=True)
                                If answer = MsgBoxResult.Cancel Then GoTo done
                            Else
                                .doDBModif(WbIsSaving:=True)
                            End If
                        End If
                    End With
                Next
            Next
        End If
done:
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
            Globals.UserMsg("Error getting docproperties: " + Wb.Name + ex.Message)
        End Try

        ' now clear content/all
        Dim searchCell As Excel.Range
        dontCalcWhileClearing = True
        Try
            Dim ws As Excel.Worksheet = Nothing
            For Each ws In Wb.Worksheets
                If ws Is Nothing Then
                    Globals.LogWarn("no worksheet in saving workbook...")
                    Exit For
                End If
                For Each theFunc As String In {"DBListFetch(", "DBRowFetch("}
                    searchCell = ws.Cells.Find(What:=theFunc, After:=ws.Range("A1"), LookIn:=Excel.XlFindLookIn.xlFormulas, LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlNext, MatchCase:=False)
                    If Not (searchCell Is Nothing) Then
                        Dim firstAddress As String = searchCell.Address
                        Do
                            ' get DB function target names from source names
                            Dim targetName As String = getUnderlyingDBNameFromRange(searchCell)
                            ' in case of commented cells and purged db underlying names, getDBunderlyingNameFromRange doesn't return a name ...
                            If InStr(UCase(targetName), "DBFSOURCE") > 0 Then
                                targetName = Replace(targetName, "DBFsource", "DBFtarget", 1, , vbTextCompare)
                                ' check which DB functions should be content cleared (CC) or all cleared (CA)
                                Dim DBFCC As Boolean = False : Dim DBFCA As Boolean = False
                                DBFCC = DBFCContentColl.Contains("*")
                                DBFCC = DBFCContentColl.Contains(searchCell.Parent.Name + "!" + Replace(searchCell.Address, "$", "")) Or DBFCC
                                DBFCA = DBFCAllColl.Contains("*")
                                DBFCA = DBFCAllColl.Contains(searchCell.Parent.Name + "!" + Replace(searchCell.Address, "$", "")) Or DBFCA
                                Dim theTargetRange As Excel.Range
                                Try : theTargetRange = ExcelDnaUtil.Application.Range(targetName)
                                Catch ex As Exception
                                    Globals.LogWarn("Error in finding target range of DB Function " + theFunc + "in " + firstAddress + "), refreshing all DB functions should solve this.")
                                    searchCell = Nothing
                                    searchCell = ws.Cells.Find(What:="", After:=ws.Range("A1"), LookIn:=Excel.XlFindLookIn.xlFormulas, LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlNext, MatchCase:=False)
                                    dontCalcWhileClearing = False
                                    Exit Sub
                                End Try
                                If DBFCC Then
                                    theTargetRange.Parent.Range(theTargetRange.Parent.Cells(theTargetRange.Row, theTargetRange.Column), theTargetRange.Parent.Cells(theTargetRange.Row + theTargetRange.Rows.Count - 1, theTargetRange.Column + theTargetRange.Columns.Count - 1)).ClearContents
                                    Globals.LogInfo("Contents of selected DB Functions targets cleared")
                                End If
                                If DBFCA Then
                                    theTargetRange.Parent.Range(theTargetRange.Parent.Cells(theTargetRange.Row + 2, theTargetRange.Column), theTargetRange.Parent.Cells(theTargetRange.Row + theTargetRange.Rows.Count - 1, theTargetRange.Column + theTargetRange.Columns.Count - 1)).Clear
                                    theTargetRange.Parent.Range(theTargetRange.Parent.Cells(theTargetRange.Row, theTargetRange.Column), theTargetRange.Parent.Cells(theTargetRange.Row + 2, theTargetRange.Column + theTargetRange.Columns.Count - 1)).ClearContents
                                    Globals.LogInfo("All cleared from selected DB Functions targets")
                                End If
                            End If
                            searchCell = ws.Cells.FindNext(searchCell)
                        Loop While searchCell IsNot Nothing AndAlso searchCell.Address <> firstAddress
                    End If
                Next
            Next
            ' always reset the cell find dialog....
            If ws IsNot Nothing Then
                searchCell = ws.Cells.Find(What:="", After:=ws.Range("A1"), LookIn:=Excel.XlFindLookIn.xlFormulas, LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlNext, MatchCase:=False)
            End If

            ' refresh content area of dbfunctions after save event, requires execution out of context of Application_WorkbookSave
            If doRefreshDBFuncsAfterSave And (DBFCContentColl.Count > 0 Or DBFCAllColl.Count > 0) Then
                ExcelAsyncUtil.QueueAsMacro(Sub()
                                                Globals.refreshDBFuncLater()
                                            End Sub)
            End If
        Catch ex As Exception
            Globals.UserMsg("Error clearing DBfunction targets in Workbook " + Wb.Name + ": " + ex.Message)
        End Try
        dontCalcWhileClearing = False
    End Sub


    ''' <summary>reset query cache, refresh DB functions and repair legacy functions if existing</summary>
    ''' <param name="Wb"></param>
    Private Sub Application_WorkbookOpen(Wb As Excel.Workbook) Handles Application.WorkbookOpen
        If Not Wb.IsAddin Then
            ' in case of reopening workbooks with dbfunctions, look for old query caches and status collections (returned error messages) and reset them to get new data
            Globals.resetCachesForWorkbook(Wb.Name)

            ' when opening, force recalculation of DB functions in workbook.
            ' this is required as there is no recalculation if no dependencies have changed (usually when opening workbooks)
            ' however the most important dependency for DB functions is the database data....
            If Not Globals.getCustPropertyBool("DBFskip", Wb) Then Globals.refreshDBFunctions(Wb)
            Globals.repairLegacyFunctions()
        End If
    End Sub

    ''' <summary>gets defined named ranges for DBMapper invocation in the current workbook after activation and updates Ribbon with it</summary>
    ''' <param name="Wb"></param>
    Private Sub Application_WorkbookActivate(Wb As Excel.Workbook) Handles Application.WorkbookActivate
        ' avoid when being activated by DBFuncsAction
        If Not DBModifs.preventChangeWhileFetching Then
            ' in case AutoOpen hasn't been triggered (e.g. when started via IE)...
            If Globals.DBModifDefColl Is Nothing Then
                Globals.DBModifDefColl = New Dictionary(Of String, Dictionary(Of String, DBModif))
            End If
            DBModifs.getDBModifDefinitions()
            ' unfortunately, Excel doesn't fire SheetActivate when opening workbooks, so do that here...
            assignHandler(Wb.ActiveSheet)
        End If
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
        ' reset noninteractive messages (used for VBA invocations) and hadError for interactive invocations
        nonInteractiveErrMsgs = "" : hadError = False
        Dim DBModifType As String = Left(cbName, 8)
        If DBModifType <> "DBSeqnce" Then
            Dim targetRange As Excel.Range
            Try
                targetRange = ExcelDnaUtil.Application.ActiveWorkbook.Names.Item(cbName).RefersToRange
            Catch ex As Exception
                ' if target name relates to an invalid (offset) formula, referstorange fails  ...
                If InStr(ExcelDnaUtil.Application.ActiveWorkbook.Names.Item(cbName).RefersTo, "OFFSET(") > 0 Then
                    Globals.UserMsg("Offset formula that '" + cbName + "' refers to, did not return a valid range." + vbCrLf + "Please check the offset formula to return a valid range !", "DBModifier Definitions Error")
                    ExcelDnaUtil.Application.Dialogs(Excel.XlBuiltInDialog.xlDialogNameManager).Show()
                Else
                    Globals.UserMsg("No underlying " + Left(cbName, 8) + " Range named " + cbName + " found, exiting without DBModification.")
                    Globals.LogWarn("targetRange assignment failed: " + ex.Message)
                End If
                Exit Sub
            End Try
            Dim DBModifName As String = getDBModifNameFromRange(targetRange)
        End If
        If My.Computer.Keyboard.CtrlKeyDown And My.Computer.Keyboard.ShiftKeyDown Then
            DBModifs.createDBModif(DBModifType, targetDefName:=cbName)
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
                If cb1 Is Nothing Then
                    cb1 = Sh.OLEObjects(shp.Name).Object
                ElseIf cb2 Is Nothing Then
                    cb2 = Sh.OLEObjects(shp.Name).Object
                ElseIf cb3 Is Nothing Then
                    cb3 = Sh.OLEObjects(shp.Name).Object
                ElseIf cb4 Is Nothing Then
                    cb4 = Sh.OLEObjects(shp.Name).Object
                ElseIf cb5 Is Nothing Then
                    cb5 = Sh.OLEObjects(shp.Name).Object
                Else
                    Globals.UserMsg("only max. of five DBModifier Buttons allowed on a Worksheet, currently using " + cb1.Name + "," + cb2.Name + "," + cb3.Name + "," + cb4.Name + " and " + cb5.Name + " !")
                    assignHandler = False
                    Exit For
                End If
            End If
        Next
    End Function

    ''' <summary>assign commandbuttons new and refresh DBAddins DBModification Menu with each change of sheets</summary>
    ''' <param name="Sh"></param>
    Private Sub Application_SheetActivate(Sh As Object) Handles Application.SheetActivate
        ' avoid when being activated by DBFuncsAction 
        If Not DBModifs.preventChangeWhileFetching Then
            'DBModifs.getDBModifDefinitions()
            ' only when needed assign button handler for this sheet ...
            If Globals.DBModifDefColl.Count > 0 Then assignHandler(Sh)
        End If
    End Sub

    ''' <summary>Clean up after closing workbook</summary>
    ''' <param name="Wb"></param>
    ''' <param name="Cancel"></param>
    Private Sub Application_WorkbookBeforeClose(Wb As Workbook, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeClose
        If Globals.DBModifDefColl.Count > 0 Then
            Globals.DBModifDefColl.Clear()
            Globals.theRibbon.Invalidate()
        End If
    End Sub

    ''' <summary>Event Procedure needed for CUD DBMappers to capture changes/insertions and set U/D Flag</summary>
    ''' <param name="Sh"></param>
    ''' <param name="Target"></param>
    Private Sub Application_SheetChange(Sh As Object, Target As Range) Handles Application.SheetChange
        ' avoid entering into check or doCUDMarks if not table, whole row/column modified, no DBMapper and prevention while fetching (on refresh) being set
        If Not IsNothing(Target.ListObject) AndAlso Not Target.Columns.Count = Sh.Columns.Count AndAlso Not Target.Rows.Count = Sh.Rows.Count AndAlso Globals.DBModifDefColl.ContainsKey("DBMapper") AndAlso Not DBModifs.preventChangeWhileFetching Then
            Dim targetName As String = DBModifs.getDBModifNameFromRange(Target)
            If Left(targetName, 8) = "DBMapper" Then
                DirectCast(Globals.DBModifDefColl("DBMapper").Item(targetName), DBMapper).insertCUDMarks(Target)
            End If
        End If
    End Sub


    Private WithEvents mInsertButton As Microsoft.Office.Core.CommandBarButton
    Private WithEvents mDeleteButton As Microsoft.Office.Core.CommandBarButton

    ''' <summary>Additionally to statically defined context menu in Ribbon this is needed to handle the dynamically displayed CUD DBMapper context menu entries (insert/delete)</summary>
    ''' <param name="Sh"></param>
    ''' <param name="Target"></param>
    ''' <param name="Cancel"></param>
    Private Sub Application_SheetBeforeRightClick(Sh As Object, Target As Range, ByRef Cancel As Boolean) Handles Application.SheetBeforeRightClick
        ' check if we are in a DBMapper, if not then leave...
        If Globals.DBModifDefColl.ContainsKey("DBMapper") Then
            Dim targetName As String = getDBModifNameFromRange(Target)
            If Left(targetName, 8) <> "DBMapper" Then Exit Sub
        Else
            Exit Sub
        End If
        Dim appsCommandBars As String() = {"List Range Popup", "Row"}
        'first delete buttons
        For Each builtin As String In appsCommandBars
            Dim srchInsertButton = ExcelDnaUtil.Application.CommandBars(builtin).FindControl(Tag:="insTag")
            Dim srchDeleteButton = ExcelDnaUtil.Application.CommandBars(builtin).FindControl(Tag:="delTag")
            If srchInsertButton IsNot Nothing Then srchInsertButton.Delete()
            If srchDeleteButton IsNot Nothing Then srchDeleteButton.Delete()
        Next
        ' add context menus
        ' for whole sheet don't display DBSheet context menus !!
        If Not Target.Rows.Count = Target.EntireColumn.Rows.Count Then
            For Each builtin As String In appsCommandBars
                With ExcelDnaUtil.Application.CommandBars(builtin).Controls.Add(Type:=1, Before:=1, Temporary:=True)
                    .caption = "delete Row (Ctl-Sh-D)"
                    .FaceID = 214
                    .Tag = "delTag"
                End With
                With ExcelDnaUtil.Application.CommandBars(builtin).Controls.Add(Type:=1, Before:=1, Temporary:=True)
                    .caption = "insert Row (Ctl-Sh-I)"
                    .FaceID = 213
                    .Tag = "insTag"
                End With
            Next
        End If
        mInsertButton = ExcelDnaUtil.Application.CommandBars.FindControl(Tag:="insTag")
        mDeleteButton = ExcelDnaUtil.Application.CommandBars.FindControl(Tag:="delTag")
    End Sub

    ''' <summary>dynamic context menu item delete: delete row in CUD Style DBMappers</summary>
    ''' <param name="Ctrl"></param>
    ''' <param name="CancelDefault"></param>
    Private Sub mDeleteButton_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles mDeleteButton.Click
        DBModifs.deleteRow()
    End Sub

    ''' <summary>dynamic context menu item insert: insert row in CUD Style DBMappers</summary>
    ''' <param name="Ctrl"></param>
    ''' <param name="CancelDefault"></param>
    Private Sub mInsertButton_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles mInsertButton.Click
        DBModifs.insertRow()
    End Sub
End Class