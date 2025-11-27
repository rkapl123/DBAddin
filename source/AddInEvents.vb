Imports ExcelDna.Integration
Imports ExcelDna.Registration
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel ' for event procedures...
Imports Microsoft.Office.Core
Imports Microsoft.Vbe.Interop ' also need to add reference to Microsoft.Vbe.Interop.Forms, otherwise commandbuttons cb1 to cb0 won't work
Imports System.Runtime.InteropServices
Imports System.Collections.Generic
Imports System.Xml.Linq


''' <summary>AddIn Connection class, also handling Events from Excel (Open, Close, Activate)</summary>
<ComVisible(True)>
Public Class AddInEvents
    Implements IExcelAddIn

    ''' <summary>collection of query refresh handlers for query objects inside list objects</summary>
    Public colQueries As Collection
    ''' <summary>collection of command button handlers for assigned DB modifiers</summary>
    Public Shared colCommandButtons As Collection

    ''' <summary>the application object needed for excel event handling (most of this class is dedicated to that)</summary>
    WithEvents Application As Excel.Application
    ''' <summary>context menu insert button</summary>
    Private WithEvents mInsertButton As Microsoft.Office.Core.CommandBarButton
    ''' <summary>context menu delete button</summary>
    Private WithEvents mDeleteButton As Microsoft.Office.Core.CommandBarButton

    ''' <summary>connect to Excel when opening Addin</summary>
    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        ' necessary for ExplicitRegistration of param arrays (https://groups.google.com/forum/#!topic/exceldna/kf76nqAqDUo)
        ExcelRegistration.GetExcelFunctions().ProcessParamsRegistrations().RegisterFunctions()
        ExcelRegistration.GetExcelCommands().RegisterCommands()
        Application = ExcelDnaUtil.Application
        Try
            If ExcelDnaUtil.Application.AddIns("DBAddin.Functions").Installed Then
                UserMsg("Attention: legacy DBAddin (DBAddin.Functions) still active, this might lead to unexpected results!")
            End If
        Catch ex As Exception : End Try
        ' for finding out what happened attach internal trace to ExcelDNA LogDisplay
        theLogDisplaySource = New Diagnostics.TraceSource("ExcelDna.Integration")
        ' and also define a LogSource for DBAddin itself for writing text log messages
        theLogFileSource = New Diagnostics.TraceSource("DBAddin")

        ' IntelliSense needed for DB- and supporting functions
        ExcelDna.IntelliSense.IntelliSenseServer.Install()

        ' caches and collection initialization 
        LogInfo("initialize configuration settings")
        Functions.queryCache = New Dictionary(Of String, String)
        Functions.StatusCollection = New Dictionary(Of String, ContainedStatusMsg)
        DBModifDefColl = New Dictionary(Of String, Dictionary(Of String, DBModif))
        previousCUDLengthColl = New Dictionary(Of String, Integer)

        ' initialize settings and get the default environment
        initSettings()
        If environdefs.Length = 0 Then
            UserMsg("Couldn't load any Environment, please check DB-Addin configurations !")
            Exit Sub
        End If
        ' Configs are 1 based, selectedEnvironment(index of environment dropdown) is 0 based. negative values not allowed!
        Dim selEnv As Integer = fetchSettingInt("DefaultEnvironment", "1") - 1
        If selEnv > environdefs.Length - 1 OrElse selEnv < 0 Then
            UserMsg("Default Environment " + (selEnv + 1).ToString() + " not existing, setting to first environment !")
            selEnv = 0
        End If
        selectedEnvironment = selEnv
        ' after getting the default environment (should exist), set the Const Connection String again to avoid problems in generating DB ListObjects
        ConstConnString = fetchSetting("ConstConnString" + env(), "")
        ConfigStoreFolder = fetchSetting("ConfigStoreFolder" + env(), "")
        checkHiddenExcelInstance()
    End Sub

    ''' <summary>AutoClose cleans up after finishing addin</summary>
    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        ' reset last assigned shortcuts
        Try : ExcelDnaUtil.Application.OnKey(refreshDataKey) : Catch ex As Exception : End Try
        Try : ExcelDnaUtil.Application.OnKey(jumpButtonKey) : Catch ex As Exception : End Try
        Try : ExcelDnaUtil.Application.OnKey(deleteRowKey) : Catch ex As Exception : End Try
        Try : ExcelDnaUtil.Application.OnKey(insertRowKey) : Catch ex As Exception : End Try
        ExcelDna.IntelliSense.IntelliSenseServer.Uninstall()
    End Sub

    ''' <summary>saves defined DBMaps (depending on configuration), also used to remove contents of DBListfunction results (data safety/space consumption)
    ''' choosing functions for removal of target data is done with custom doc-properties</summary>
    ''' <param name="Wb"></param>
    ''' <param name="SaveAsUI"></param>
    ''' <param name="Cancel"></param>
    Private Sub Application_WorkbookSave(Wb As Excel.Workbook, ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeSave
        Try
            ' ask if modifications should be done if no overriding flag is defined...
            Dim doDBMOnSave As Boolean = getCustPropertyBool("doDBMOnSave", Wb)
            Dim askForEveryModifier As Boolean = False
            ' if overriding flag not given and Readonly is NOT recommended on this workbook and workbook IS NOT Readonly, ...
            If Not (Wb.ReadOnlyRecommended And Wb.ReadOnly) And Not doDBMOnSave Then
                ' prevent action/editing in case of errors while DB Modifiers were added
                If DBModifDefColl.ContainsKey("Error") Then
                    Dim keylist As New List(Of String)(DBModifDefColl("Error").Keys())
                    UserMsg("DBModifier definitions not accessible due to error: " + vbCrLf + keylist(0))
                    Exit Sub
                End If

                Dim TotalDBModifCount As Integer = 0
                For Each DBmodifType As String In DBModifDefColl.Keys
                    For Each dbmapdefkey As String In DBModifDefColl(DBmodifType).Keys
                        If DBModifDefColl(DBmodifType).Item(dbmapdefkey).execOnSave Then TotalDBModifCount += 1
                    Next
                Next
                ' ...ask for saving, if this is necessary for any DBmodifier...
                If TotalDBModifCount > 1 Then
                    ' multiple DBmodifiers, ask how to proceed
                    Dim answer As MsgBoxResult = QuestionMsg(theMessage:="do all DB Modifications defined in Workbook (Yes=All with exec on Save, No=Ask every time. Cancel=Don't do any DB Modifications) ?", questionType:=MsgBoxStyle.YesNoCancel, questionTitle:="Do DB Modifiers on Save")
                    If answer = MsgBoxResult.Yes Then doDBMOnSave = True
                    If answer = MsgBoxResult.Cancel Then doDBMOnSave = False
                    If answer = MsgBoxResult.No Then
                        doDBMOnSave = True
                        askForEveryModifier = True
                    End If
                ElseIf TotalDBModifCount = 1 Then
                    ' only one DBModifier needs saving, ask only once for saving...
                    For Each DBmodifType As String In DBModifDefColl.Keys
                        For Each dbmapdefkey As String In DBModifDefColl(DBmodifType).Keys
                            If DBModifDefColl(DBmodifType).Item(dbmapdefkey).execOnSave Then
                                If nonInteractive Then
                                    ' always save in non interactive (headless/automation) mode
                                    doDBMOnSave = True
                                Else
                                    doDBMOnSave = IIf(DBModifDefColl(DBmodifType).Item(dbmapdefkey).confirmExecution(WbIsSaving:=True) = MsgBoxResult.Yes, True, False)
                                End If
                            End If
                        Next
                    Next
                End If
            End If

            ' save all DBmappers/DBActions/DBSequences on saving if above resulted in YES!
            If doDBMOnSave Then
                ' first do DBModifiers defined on active sheet or any DB Sequence:
                For Each DBmodifType As String In DBModifDefColl.Keys
                    For Each dbmapdefkey As String In DBModifDefColl(DBmodifType).Keys
                        With DBModifDefColl(DBmodifType).Item(dbmapdefkey)
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
                For Each DBmodifType As String In DBModifDefColl.Keys
                    For Each dbmapdefkey As String In DBModifDefColl(DBmodifType).Keys
                        With DBModifDefColl(DBmodifType).Item(dbmapdefkey)
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
        Catch ex As Exception
            UserMsg("Error doing DBMapper on save: " + Wb.Name + ex.Message)
        End Try
done:
        ' clear DB Functions content and refresh afterwards..
        Dim DBFCContentColl As New Collection
        Dim DBFCAllColl As New Collection
        Dim doRefreshDBFuncsAfterSave As Boolean = True
        ' first insert doc property information into collections for easier handling
        Try
            Dim docproperty As DocumentProperty
            For Each docproperty In Wb.CustomDocumentProperties
                If docproperty.Type = MsoDocProperties.msoPropertyTypeBoolean Then
                    If Left$(docproperty.Name, 5) = "DBFCC" And docproperty.Value Then DBFCContentColl.Add(True, Mid(docproperty.Name, 6))
                    If Left$(docproperty.Name, 5) = "DBFCA" And docproperty.Value Then DBFCAllColl.Add(True, Mid(docproperty.Name, 6))
                    If docproperty.Name = "DBFskip" AndAlso docproperty.Value Then doRefreshDBFuncsAfterSave = False
                End If
            Next
        Catch ex As Exception
            UserMsg("Error getting doc properties: " + Wb.Name + ex.Message)
        End Try
        ' skip searching for functions (takes time in large sheets!) if not necessary
        If Not doRefreshDBFuncsAfterSave Or (DBFCContentColl.Count = 0 And DBFCAllColl.Count = 0) Then Exit Sub

        ' now clear content/all
        dontCalcWhileClearing = True
        Try
            ' walk through all DB functions (having hidden names DBFsource*) cells there to find DB Functions and change their formula, adding " " to trigger recalculation
            For Each DBname As Excel.Name In ExcelDnaUtil.Application.ActiveWorkbook.Names
                Dim DBFuncCell As Excel.Range = Nothing
                If DBname.Name Like "*DBFsource*" Then
                    ' some names might have lost their reference to the cell, so catch this here...
                    Try : DBFuncCell = DBname.RefersToRange : Catch ex As Exception : End Try
                End If
                If Not IsNothing(DBFuncCell) Then
                    Try
                        ' get DB function target names from source names
                        Dim targetName As String = getUnderlyingDBNameFromRange(DBFuncCell)
                        ' in case of commented cells and purged db underlying names, getDBunderlyingNameFromRange doesn't return a name ...
                        If InStr(UCase(targetName), "DBFSOURCE") > 0 Then
                            targetName = Replace(targetName, "DBFsource", "DBFtarget", 1, , vbTextCompare)
                            ' check which DB functions should be content cleared (CC) or all cleared (CA)
                            Dim DBFCC As Boolean = False : Dim DBFCA As Boolean = False
                            DBFCC = DBFCContentColl.Contains("*")
                            DBFCC = DBFCContentColl.Contains(DBFuncCell.Parent.Name + "!" + Replace(DBFuncCell.Address, "$", "")) Or DBFCC
                            DBFCA = DBFCAllColl.Contains("*")
                            DBFCA = DBFCAllColl.Contains(DBFuncCell.Parent.Name + "!" + Replace(DBFuncCell.Address, "$", "")) Or DBFCA
                            Dim theTargetRange As Excel.Range
                            Try : theTargetRange = ExcelDnaUtil.Application.Range(targetName)
                            Catch ex As Exception
                                LogWarn("Error in finding target range of DB Function in " + DBFuncCell.Address + "), refreshing all DB functions should solve this.")
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
                    Catch ex As Exception
                        LogWarn("Exception when clearing target range of db function in Cell (" + DBFuncCell.Address + "): " + ex.Message)
                    End Try
                End If
            Next
            ExcelDnaUtil.Application.Statusbar = False

            ' refresh content area of db functions after save event, requires execution out of context of Application_WorkbookSave
            If DBFCContentColl.Count > 0 Or DBFCAllColl.Count > 0 Then
                ExcelAsyncUtil.QueueAsMacro(Sub()
                                                refreshDBFuncLater()
                                            End Sub)
            End If
        Catch ex As Exception
            UserMsg("Error clearing DBfunction targets in Workbook " + Wb.Name + ": " + ex.Message)
        End Try
        dontCalcWhileClearing = False
    End Sub

    ''' <summary>reset query cache, refresh DB functions and repair legacy functions if existing</summary>
    ''' <param name="Wb"></param>
    Private Sub Application_WorkbookOpen(Wb As Excel.Workbook) Handles Application.WorkbookOpen
        If Not Wb.IsAddin Then
            LogInfo("repair legacy functions for workbook on opening: " + Wb.Name)
            repairLegacyFunctions(Wb)
            Functions.preventRefreshFlag = False
            ' when opening, force recalculation of DB functions in workbook.
            ' this is required as there is no recalculation as no dependencies have changed (usually when opening workbooks) and all db functions have been realized as non-volatile.
            ' still, the most important dependency for DB functions is the database data, nevertheless, there is an exception:
            ' if a dbfunction depends on other volatile functions (e.g. today()), then it is already being calculated before the WorkbookOpen event,
            ' in this case the queryCache prevents another fetching from the database during refreshDBFunctions
            If Not getCustPropertyBool("DBFskip", Wb) Then
                LogInfo("refreshing DB functions for workbook: " + Wb.Name)
                refreshDBFunctions(Wb, False, True)
            End If
        End If
    End Sub

    ''' <summary>gets defined named ranges for DBMapper invocation in the current workbook after activation and updates Ribbon with it</summary>
    ''' <param name="Wb"></param>
    Private Sub Application_WorkbookActivate(Wb As Excel.Workbook) Handles Application.WorkbookActivate
        ' avoid when being activated by DBFuncsAction
        If Not DBModifHelper.preventChangeWhileFetching And Not Wb.IsAddin Then
            ' in case AutoOpen hasn't been triggered (e.g. when Excel was started via Internet Explorer)...
            If DBModifDefColl Is Nothing Then
                DBModifDefColl = New Dictionary(Of String, Dictionary(Of String, DBModif))
            End If
            LogInfo("getting DBModif Definitions on Workbook Activate")
            getDBModifDefinitions(Wb) ' this also invalidates the ribbon to reflect any DB Modifier changes there
            ' unfortunately, Excel doesn't fire SheetActivate when opening workbooks, so do that here...
            LogInfo("assign command button click and listobject query handlers on Workbook Activate")
            InitializeQueryHandlers(Wb)
            InitializeCBHandlers(Wb)
            LogInfo("finished actions on Workbook Activate")
        End If
    End Sub

    ''' <summary>Flag for sharing closing state between Application_WorkbookBeforeClose and Application_WorkbookDeactivate (true for actually closed workbooks)</summary>
    Private WbIsClosing As Boolean = False

    ''' <summary>Clean up after closing workbook, only set flag here, the actual cleanup is only done if workbook is really closed (in WB_Deactivate event)</summary>
    ''' <param name="Wb"></param>
    ''' <param name="Cancel"></param>
    Private Sub Application_WorkbookBeforeClose(Wb As Workbook, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeClose
        LogInfo("Workbook closing: " + Wb.Name)
        WbIsClosing = True
    End Sub

    ''' <summary>Actually clean up after closing workbook</summary>
    ''' <param name="Wb"></param>
    Private Sub Application_WorkbookDeactivate(Wb As Workbook) Handles Application.WorkbookDeactivate
        Try
            LogInfo("Workbook Deactivating: " + Wb.Name)
            If WbIsClosing AndAlso preventRefreshFlagColl.ContainsKey(Wb.Name) Then preventRefreshFlagColl.Remove(Wb.Name)
            If WbIsClosing AndAlso Not IsNothing(DBModifDefColl) AndAlso DBModifDefColl.Count > 0 Then
                DBModifDefColl.Clear()
                theRibbon.Invalidate()
                ' reset query caches and status collections (returned error messages) to get new data later on reopening the workbook
                resetCachesForWorkbook(Wb.Name)
            End If
            cleanupHandlers()
        Catch ex As Exception : End Try
        WbIsClosing = False
    End Sub

    ''' <summary>Event Procedure needed for CUD DBMappers to capture changes/insertions and set U/D Flag</summary>
    ''' <param name="Sh"></param>
    ''' <param name="Target"></param>
    Private Sub Application_SheetChange(Sh As Object, Target As Range) Handles Application.SheetChange
        ' avoid entering into insert/update check resp. doCUDMarks if not list-object (data table), whole column modified, no DBMapper present and prevention while fetching (on refresh) being set
        If Not IsNothing(Target.ListObject) AndAlso Not Target.Rows.Count = Sh.Rows.Count AndAlso DBModifDefColl.ContainsKey("DBMapper") AndAlso Not DBModifHelper.preventChangeWhileFetching Then
            Dim targetName As String = getDBModifNameFromRange(Target)
            If Left(targetName, 8) = "DBMapper" Then
                DirectCast(DBModifDefColl("DBMapper").Item(targetName), DBMapper).insertCUDMarks(Target)
            End If
        End If
    End Sub

    ''' <summary>use Application_AfterCalculate to overcome the problem of auto-fitting formula ranges AFTER calculation (final column width is not available in dblistfetchAction procedure)</summary>
    Private Sub Application_AfterCalculate() Handles Application.AfterCalculate
        Try
            Dim actWbNames As Excel.Names = ExcelDnaUtil.Application.ActiveWorkbook.Names
            For Each nm As Excel.Name In actWbNames
                LogInfo("Application_AfterCalculate: " + nm.Name)
                ' look only for formula target ranges
                If Left$(nm.Name, 10) = "DBFtargetF" Then
                    ' get to the source cell to look up info in the StatusCollection
                    Dim sourceName As String = Replace(nm.Name, "DBFtargetF", "DBFsource", 1, , vbTextCompare)
                    Dim sourceCell As Excel.Range = Nothing
                    Try : sourceCell = ExcelDnaUtil.Application.ActiveWorkbook.Names(sourceName).RefersToRange : Catch ex As Exception : End Try
                    ' build the lookup key from the sourceCell
                    Dim callID As String = "[" + sourceCell.Parent.Parent.Name + "]" + sourceCell.Parent.Name + "!" + sourceCell.Address
                    If StatusCollection.ContainsKey(callID) Then
                        ' if formulaRange is available (only set by dblistfetchAction if AutoFit is set to true), then auto fit columns and rows
                        If StatusCollection(callID).formulaRange IsNot Nothing Then
                            StatusCollection(callID).formulaRange.Columns.EntireColumn.AutoFit()
                            StatusCollection(callID).formulaRange.Rows.EntireRow.AutoFit()
                            StatusCollection(callID).formulaRange = Nothing
                        End If
                    End If
                End If
            Next
        Catch ex As Exception : End Try
    End Sub

    ''' <summary>Additionally to statically defined context menu in Ribbon this is needed to handle the dynamically displayed CUD DBMapper context menu entries (insert/delete)</summary>
    ''' <param name="Sh"></param>
    ''' <param name="Target"></param>
    ''' <param name="Cancel"></param>
    Private Sub Application_SheetBeforeRightClick(Sh As Object, Target As Range, ByRef Cancel As Boolean) Handles Application.SheetBeforeRightClick
        ' check if we are in a CUD DBMapper, if not then leave...
        If Not IsNothing(Target.ListObject) AndAlso Not IsNothing(DBModifDefColl) AndAlso DBModifDefColl.ContainsKey("DBMapper") Then
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
        deleteRow()
    End Sub

    ''' <summary>dynamic context menu item insert: insert row in CUD Style DBMappers</summary>
    ''' <param name="Ctrl"></param>
    ''' <param name="CancelDefault"></param>
    Private Sub mInsertButton_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles mInsertButton.Click
        insertRow()
    End Sub

    ''' <summary>used for releasing com objects</summary>
    Protected Overrides Sub Finalize()
        LogInfo("Addin finalizing: Base finalize")
        MyBase.Finalize()
        LogInfo("Addin finalizing: releasing com objects of control buttons")
        cleanupHandlers()
    End Sub

    ''' <summary>assign click handlers to command buttons in passed workbook Wb</summary>
    ''' <param name="wb">Workbook where command buttons are located</param>
    Public Sub InitializeCBHandlers(wb As Object)
        Dim cbCH As CommandbuttonClickHandler
        colCommandButtons = New Collection

        If Not IsNothing(DBModifDefColl) AndAlso DBModifDefColl.Count > 0 Then
            Try
                For Each ws As Worksheet In wb.Worksheets
                    For Each shp As Excel.Shape In ws.Shapes
                        ' only for OLE Control buttons...
                        If shp.Type = MsoShapeType.msoOLEControlObject Then
                            ' Associate click-event handler of a CommandButton if its name matches the DB modifiers name.
                            Dim ctrlName As String
                            Try : ctrlName = ws.OLEObjects(shp.Name).Object.Name : Catch ex As Exception : ctrlName = "" : End Try
                            If Left(ctrlName, 8) = "DBMapper" Or Left(ctrlName, 8) = "DBAction" Or Left(ctrlName, 8) = "DBSeqnce" Then
                                cbCH = New CommandbuttonClickHandler With {.cb = ws.OLEObjects(shp.Name).Object}
                                colCommandButtons.Add(cbCH)
                            End If
                        End If
                    Next
                Next
            Catch ex As Exception
                LogWarn("InitializeCBHandlers exception occurred: " + ex.Message)
            End Try
        End If
    End Sub

    ''' <summary>initialize query handlers inside list objects for workbook wb, needed for catching refresh event to avoid sheetchange event execution</summary>
    ''' <param name="wb"></param>
    Private Sub InitializeQueryHandlers(wb As Workbook)
        Dim qryRH As QueryRefreshHandler
        colQueries = New Collection

        For Each ws As Worksheet In wb.Worksheets
            For Each lo As ListObject In ws.ListObjects
                qryRH = New QueryRefreshHandler With {.qry = lo.QueryTable}
                colQueries.Add(qryRH)
            Next
        Next
    End Sub

    ''' <summary>COM cleanup for query table objects and command buttons inside their handlers</summary>
    Private Sub cleanupHandlers()
        For Each qryRH As QueryRefreshHandler In colQueries
            Marshal.ReleaseComObject(qryRH.qry)
        Next
        colQueries.Clear()
        For Each cbCH As CommandbuttonClickHandler In colCommandButtons
            Marshal.ReleaseComObject(cbCH.cb)
        Next
        colCommandButtons.Clear()
    End Sub

End Class

''' <summary>Event handler class to catch refresh events on query tables inside list objects</summary>
Class QueryRefreshHandler
    ''' <summary>QueryTable that needs to be observed to avoid change registration</summary>
    Public WithEvents qry As QueryTable

    ''' <summary>open preventChangeWhileFetching bracket before refresh (avoid any sheet change event triggered activity for DBSheets)</summary>
    ''' <param name="Cancel"></param>
    Private Sub qry_BeforeRefresh(ByRef Cancel As Boolean) Handles qry.BeforeRefresh
        DBModifHelper.preventChangeWhileFetching = True
    End Sub

    ''' <summary>close preventChangeWhileFetching bracket afterwards (enable sheet change event triggered activity again)</summary>
    Private Sub qry_AfterRefresh(Success As Boolean) Handles qry.AfterRefresh
        DBModifHelper.preventChangeWhileFetching = False
    End Sub
End Class

''' <summary>Event handler class for click events on control buttons that are associated to DB modifiers</summary>
Class CommandbuttonClickHandler
    ''' <summary>CommandButton that can be inserted on a worksheet (name property being the same as the respective target range (for DBMapper/DBAction) or DBSeqnce Name)</summary>
    Public WithEvents cb As Forms.CommandButton

    ''' <summary>click handler for all command buttons that are associated to DB modifiers</summary>
    Private Sub cb_Click() Handles cb.Click
        ' name of command button, defines whether a DBModification is invoked (starts with DBMapper/DBAction/DBSeqnce)
        Dim cbName As String = cb.Name
        ' reset non interactive messages (used for VBA invocations) and hadError for interactive invocations
        nonInteractiveErrMsgs = "" : DBModifHelper.hadError = False
        ' prevent action/editing in case of errors while DB Modifiers were added
        If DBModifDefColl.ContainsKey("Error") Then
            Dim keylist As New List(Of String)(DBModifDefColl("Error").Keys())
            UserMsg("DBModifier definitions not accessible due to error: " + vbCrLf + keylist(0))
            Exit Sub
        End If
        Dim DBModifType As String = Left(cbName, 8)
        If DBModifType <> "DBSeqnce" Then
            Dim targetRange As Excel.Range
            Try
                targetRange = ExcelDnaUtil.Application.ActiveWorkbook.Names.Item(cbName).RefersToRange
            Catch ex As Exception
                ' if target name relates to an invalid (offset) formula, referstorange fails  ...
                If InStr(ExcelDnaUtil.Application.ActiveWorkbook.Names.Item(cbName).RefersTo, "OFFSET(") > 0 Then
                    UserMsg("Offset formula that '" + cbName + "' refers to, did not return a valid range." + vbCrLf + "Please check the offset formula to return a valid range !", "DBModifier Definitions Error")
                    ExcelDnaUtil.Application.Dialogs(Excel.XlBuiltInDialog.xlDialogNameManager).Show()
                Else
                    UserMsg("No underlying " + Left(cbName, 8) + " Range named " + cbName + " found, exiting without DBModification.")
                    LogWarn("targetRange assignment failed: " + ex.Message)
                End If
                Exit Sub
            End Try
            Dim DBModifName As String = getDBModifNameFromRange(targetRange)
        End If
        If My.Computer.Keyboard.CtrlKeyDown And My.Computer.Keyboard.ShiftKeyDown Then
            createDBModif(DBModifType, targetDefName:=cbName)
        Else
            DBModifDefColl(DBModifType).Item(cbName).doDBModif()
        End If
    End Sub

End Class