Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports System.Collections.Generic

''' <summary>Tools repairLegacyFunctions, checkpurgeNames and fixOrphanedDBFunctions</summary>
Public Module SupportingTools

    ''' <summary>"repairs" legacy functions from old VB6-COM Addin by removing "DBAddin.Functions." before function name</summary>
    ''' <param name="showResponse">in case this is called interactively, provide a response in case of no legacy functions there</param>
    Public Sub repairLegacyFunctions(actWB As Excel.Workbook, Optional showResponse As Boolean = False)
        Dim foundLegacyFunc As Boolean = False
        Dim xlcalcmode As Long = ExcelDnaUtil.Application.Calculation
        Dim WbNames As Excel.Names

        ' skip repair on auto open if explicitly set
        If Not fetchSettingBool("repairLegacyFunctionsAutoOpen", "True") AndAlso Not showResponse Then Exit Sub

        Try : WbNames = actWB.Names
        Catch ex As Exception
            LogWarn("Exception when trying to get Workbook names for repairing legacy functions: " + ex.Message + ", this might be due to errors in the VBA Macros (missing references)")
            Exit Sub
        End Try
        If actWB Is Nothing Then
            ' only log warning, no user message !
            LogWarn("no active workbook available !")
            Exit Sub
        End If
        DBModifHelper.preventChangeWhileFetching = True ' WorksheetFunction.CountIf triggers Change event with target in argument 1, so make sure this doesn't trigger anything inside DBAddin)
        Try
            ' count nonempty cells in workbook for time estimate...
            Dim cellcount As Long = 0
            For Each ws In actWB.Worksheets
                cellcount += ExcelDnaUtil.Application.WorksheetFunction.CountIf(ws.Range("1:" + ws.Rows.Count.ToString()), "<>")
            Next
            ' if interactive, enforce replace...
            If showResponse Then foundLegacyFunc = True
            Dim timeEstInSec As Double = cellcount / 3500000
            For Each DBname As Excel.Name In WbNames
                If DBname.Name Like "*DBFsource*" Then
                    ' some names might have lost their reference to the cell, so catch this here...
                    Try : foundLegacyFunc = DBname.RefersToRange.Formula.ToString().Contains("DBAddin.Functions") : Catch ex As Exception : End Try
                End If
                If foundLegacyFunc Then Exit For
            Next
            Dim retval As MsgBoxResult
            If foundLegacyFunc Then
                retval = QuestionMsg(fetchSetting("legacyFunctionMsg", IIf(showResponse, "Fix legacy DBAddin functions", "Found legacy DBAddin functions") + " in active workbook, should they be replaced with current addin functions (Save workbook afterwards to persist)? Estimated time for replace: ") + timeEstInSec.ToString("0.0") + " sec.", MsgBoxStyle.OkCancel, "Legacy DBAddin functions")
            ElseIf showResponse Then
                retval = QuestionMsg("No DBListfetch/DBRowfetch/DBSetQuery found in active workbook (via hidden names), still try to fix legacy DBAddin functions (Save workbook afterwards to persist)? Estimated time for replace: " + timeEstInSec.ToString("0.0") + " sec.", MsgBoxStyle.OkCancel, "Legacy DBAddin functions")
            End If
            If retval = MsgBoxResult.Ok Then
                Dim replaceSheets As String = ""
                ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual ' avoid recalculations during replace action
                ExcelDnaUtil.Application.DisplayAlerts = False ' avoid warnings for sheet where "DBAddin.Functions." is not found
                ' remove "DBAddin.Functions." in each sheet...
                For Each ws In actWB.Worksheets
                    ExcelDnaUtil.Application.StatusBar = "Replacing legacy DB functions in active workbook, sheet '" + ws.Name + "'."
                    If ws.Cells.Replace(What:="DBAddin.Functions.", Replacement:="", LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False) Then
                        replaceSheets += ws.Name + ","
                    End If
                Next
                ExcelDnaUtil.Application.Calculation = xlcalcmode
                ' reset the cell find dialog....
                ExcelDnaUtil.Application.ActiveSheet.Cells.Find(What:="", After:=ExcelDnaUtil.Application.ActiveSheet.Range("A1"), LookIn:=Excel.XlFindLookIn.xlFormulas, LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlNext, MatchCase:=False)
                ExcelDnaUtil.Application.DisplayAlerts = True
                ExcelDnaUtil.Application.StatusBar = False
                If showResponse And replaceSheets.Length > 0 Then
                    UserMsg("Replaced legacy functions in active workbook from sheets: " + Left(replaceSheets, replaceSheets.Length - 1), "Legacy DBAddin functions", MsgBoxStyle.Exclamation)
                End If
            End If
        Catch ex As Exception
            UserMsg("Exception occurred: " + ex.Message, "Legacy DBAddin functions")
        End Try
        DBModifHelper.preventChangeWhileFetching = False
    End Sub

    ''' <summary>maintenance procedure to check/purge names used for db-functions from workbook, or unhide DB names</summary>
    Public Sub checkpurgeNames()
        Dim actWb As Excel.Workbook = Nothing
        Try : actWb = ExcelDnaUtil.Application.ActiveWorkbook : Catch ex As Exception
            UserMsg("Exception when trying to get the active workbook for purging names: " + ex.Message + ", this might be either due to errors in the VBA-IDE (missing references) or due to opening this workbook from an MS-Office hyperlink, starting up Excel (timing issue).")
            Exit Sub
        End Try
        If IsNothing(actWb) Then Exit Sub
        Dim actWbNames As Excel.Names = Nothing
        Try : actWbNames = actWb.Names : Catch ex As Exception
            UserMsg("Exception when trying to get the active workbook names for purging names: " + ex.Message + ", this might be either due to errors in the VBA-IDE (missing references) or due to opening this workbook from an MS-Office hyperlink, starting up Excel (timing issue).")
            Exit Sub
        End Try
        Dim NamesWithErrors As New List(Of Excel.Name)
        If IsNothing(actWbNames) Then Exit Sub
        ' with Ctrl unhide all DB names and show Name Manager...
        If My.Computer.Keyboard.CtrlKeyDown And Not My.Computer.Keyboard.ShiftKeyDown Then
            Dim retval As MsgBoxResult = QuestionMsg("Unhiding all hidden DB function names, should ALL names (also non DB function names) also be revealed (refreshing will only hide DB function names again)?", MsgBoxStyle.YesNoCancel, "Unhide names")
            If retval = vbCancel Then Exit Sub
            For Each DBname As Excel.Name In actWbNames
                If DBname.Name Like "*DBFtarget*" Or DBname.Name Like "*DBFsource*" Or retval = vbYes Then DBname.Visible = True
            Next
            Try
                ExcelDnaUtil.Application.Dialogs(Excel.XlBuiltInDialog.xlDialogNameManager).Show()
            Catch ex As Exception
                UserMsg("The name manager dialog can't be displayed, maybe you are in the formula/cell editor?", "Name manager dialog display")
            End Try
            ' with Shift remove DBFunc names
        ElseIf My.Computer.Keyboard.ShiftKeyDown And Not My.Computer.Keyboard.CtrlKeyDown Then
            Dim resultingPurges As String = ""
            Dim retval As MsgBoxResult = QuestionMsg("Purging DBFunc names, should associated ExternalData definitions (from Queries) also be purged ? (DB-Functions can be refreshed in cells manually again)", MsgBoxStyle.YesNoCancel, "Purge DBFunc names")
            If retval = vbCancel Then Exit Sub
            Dim calcMode = ExcelDnaUtil.Application.Calculation
            ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual
            Try
                If retval = vbYes Then
                    Dim curWs As Excel.Worksheet = ExcelDnaUtil.Application.ActiveSheet
                    For Each ws As Excel.Worksheet In actWb.Worksheets
                        ws.Activate()
                        For Each DBname As Excel.Name In ws.Names
                            ' external data names
                            Dim underlyingDBName As String = ""
                            Try : underlyingDBName = getUnderlyingDBNameFromRange(ExcelDnaUtil.Application.Range(DBname.Name)) : Catch ex As Exception : End Try
                            Dim possibleQryTblName As String = "" : Dim possibleQryTbl As Excel.QueryTable = Nothing
                            Try : possibleQryTbl = ExcelDnaUtil.Application.Range(DBname.Name).QueryTable : possibleQryTblName = possibleQryTbl.Name : Catch ex As Exception : End Try
                            If DBname.Name = IIf(InStr(ws.Name, " "), "'", "") + ws.Name + IIf(InStr(ws.Name, " "), "'", "") + "!" + possibleQryTblName And Left(underlyingDBName, 9) = "DBFtarget" Then
                                resultingPurges += DBname.Name + ", "
                                possibleQryTbl.Delete()
                            End If
                        Next
                    Next
                    curWs.Activate()
                End If
                For Each DBname As Excel.Name In actWbNames
                    ' only DBFunc names...
                    Dim underlyingDBName As String = ""
                    Try : underlyingDBName = getUnderlyingDBNameFromRange(ExcelDnaUtil.Application.Range(DBname.Name)) : Catch ex As Exception : End Try
                    If Left(underlyingDBName, 9) = "DBFtarget" Or Left(underlyingDBName, 9) = "DBFsource" Then
                        resultingPurges += DBname.Name + ", "
                        DBname.Delete()
                    End If
                Next
                If resultingPurges = "" Then
                    UserMsg("nothing purged...", "purge Names", MsgBoxStyle.Information)
                Else
                    UserMsg("removed " + resultingPurges, "purge Names", MsgBoxStyle.Information)
                End If
            Catch ex As Exception
                UserMsg("Exception: " + ex.Message, "purge Names")
            End Try
            ExcelDnaUtil.Application.Calculation = calcMode
        ElseIf My.Computer.Keyboard.ShiftKeyDown And My.Computer.Keyboard.CtrlKeyDown Then
            Try
                ExcelDnaUtil.Application.Dialogs(Excel.XlBuiltInDialog.xlDialogNameManager).Show()
            Catch ex As Exception
                UserMsg("The name manager dialog can't be displayed, maybe you are in the formula/cell editor?", "Name manager dialog display")
            End Try
        Else
            Dim NamesList As Excel.Names = actWbNames
            Dim collectedErrors As String = ""
            For Each DBname As Excel.Name In NamesList
                Dim checkExists As Excel.Name = Nothing
                If DBname.Name Like "*DBFtarget*" Then
                    Dim replaceName = "DBFtarget"
                    If DBname.Name Like "*DBFtargetF*" Then replaceName = "DBFtargetF"
                    Try : checkExists = NamesList.Item(Replace(DBname.Name, replaceName, "DBFsource")) : Catch ex As Exception : End Try
                    If IsNothing(checkExists) Then
                        NamesWithErrors.Add(DBname)
                        collectedErrors += DBname.Name + "' doesn't have a corresponding DBFsource name" + vbCrLf
                    End If
                    If InStr(DBname.RefersTo, "#REF!") > 0 Then
                        NamesWithErrors.Add(DBname)
                        collectedErrors += DBname.Name + "' contains #REF!" + vbCrLf
                    End If
                    Dim checkRange As Excel.Range
                    ' might fail if target name relates to an invalid (offset) formula ...
                    Try
                        checkRange = DBname.RefersToRange
                    Catch ex As Exception
                        If InStr(DBname.RefersTo, "OFFSET(") > 0 Then
                            collectedErrors += "Offset formula that '" + DBname.Name + "' refers to, did not return a valid range" + vbCrLf
                        ElseIf InStr(DBname.RefersTo, "#REF!") > 0 Then
                            ' RefersToRange throws exception, but do nothing as already collected above ...
                        Else
                            collectedErrors += DBname.Name + "' checkRange = DBname.RefersToRange resulted in unexpected Exception " + ex.Message + vbCrLf
                        End If
                    End Try
                    If DBname.Visible Then
                        collectedErrors += DBname.Name + "' is visible" + vbCrLf
                    End If
                End If
                If DBname.Name Like "*DBFsource*" Then
                    Try : checkExists = NamesList.Item(Replace(DBname.Name, "DBFsource", "DBFtarget")) : Catch ex As Exception : End Try
                    If IsNothing(checkExists) Then
                        NamesWithErrors.Add(DBname)
                        collectedErrors += DBname.Name + "' doesn't have a corresponding DBFtarget name" + vbCrLf
                    End If
                    If InStr(DBname.RefersTo, "#REF!") > 0 Then
                        NamesWithErrors.Add(DBname)
                        collectedErrors += DBname.Name + "' contains #REF!" + vbCrLf
                    End If
                    If DBname.RefersTo = "" Then
                        NamesWithErrors.Add(DBname)
                        collectedErrors += DBname.Name + "' is empty" + vbCrLf
                    End If
                    If DBname.Visible Then
                        collectedErrors += DBname.Name + "' is visible" + vbCrLf
                    End If
                End If
            Next
            If collectedErrors = "" Then
                UserMsg("No DBfunction name problems detected.", "DBfunction check Error", MsgBoxStyle.Information)
            Else
                If QuestionMsg(collectedErrors + vbCrLf + "Should names containing #REF! errors, DBFsource names being empty or all names not having a corresponding source/target name be removed?",, "DBfunction check Error") = MsgBoxResult.Ok Then
                    For Each DBname As Excel.Name In NamesWithErrors
                        Try : DBname.Delete() : Catch ex As Exception : End Try
                    Next
                End If
            End If
            ' also provide possibility to fix orphaned DBFunctions
            fixOrphanedDBFunctions(ExcelDnaUtil.Application.ActiveWorkbook)
            ' last check any possible DB Modifier Definitions for validity
            getDBModifDefinitions(actWb, True)
        End If
    End Sub

    ''' <summary>fix orphaned DB Functions by replacing function names with themselves, triggering recalculation</summary>
    ''' <param name="actWB"></param>
    Sub fixOrphanedDBFunctions(actWB As Excel.Workbook)
        Dim xlcalcmode As Long = ExcelDnaUtil.Application.Calculation

        DBModifHelper.preventChangeWhileFetching = True ' WorksheetFunction.CountIf triggers Change event with target in argument 1, so make sure this doesn't trigger anything inside DBAddin)
        Try
            ' count nonempty cells in workbook for time estimate...
            Dim cellcount As Long = 0
            For Each ws In actWB.Worksheets
                cellcount += ExcelDnaUtil.Application.WorksheetFunction.CountIf(ws.Range("1:" + ws.Rows.Count.ToString()), "<>")
            Next

            Dim timeEstInSec As Double = cellcount / 3500000
            Dim retval As MsgBoxResult
            retval = QuestionMsg("Try to fix possibly orphaned DBAddin functions (Save workbook afterwards to persist)? Estimated time for fix: " + timeEstInSec.ToString("0.0") + " sec.", MsgBoxStyle.OkCancel, "Orphaned DBAddin functions fix")
            If retval = MsgBoxResult.Ok Then
                Dim replaceSheets As String = ""
                ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual ' avoid recalculations during replace action
                ExcelDnaUtil.Application.DisplayAlerts = False ' avoid warnings for sheets where no DBAddin functions were found
                ' replace function names in each sheet, triggering recalculation
                For Each ws In actWB.Worksheets
                    ExcelDnaUtil.Application.StatusBar = "Replacing possibly orphaned DB functions in active workbook, sheet '" + ws.Name + "'."
                    If Not IsNothing(ws.Cells.Find(What:="=DBListfetch(", LookIn:=Excel.XlFindLookIn.xlFormulas, LookAt:=Excel.XlLookAt.xlPart, MatchCase:=False, SearchFormat:=False)) Then
                        ws.Cells.Replace(What:="=DBListfetch(", Replacement:="=DBListFetch(", LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)
                        replaceSheets += "DBListfetch in: " + ws.Name + ","
                    End If
                    If Not IsNothing(ws.Cells.Find(What:="=DBRowfetch(", LookIn:=Excel.XlFindLookIn.xlFormulas, LookAt:=Excel.XlLookAt.xlPart, MatchCase:=False, SearchFormat:=False)) Then
                        ws.Cells.Replace(What:="=DBRowfetch(", Replacement:="=DBRowFetch(", LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)
                        replaceSheets += "DBRowfetch in: " + ws.Name + ","
                    End If
                    If Not IsNothing(ws.Cells.Find(What:="=DBSetquery(", LookIn:=Excel.XlFindLookIn.xlFormulas, LookAt:=Excel.XlLookAt.xlPart, MatchCase:=False, SearchFormat:=False)) Then
                        ws.Cells.Replace(What:="=DBSetquery(", Replacement:="=DBSetQuery(", LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)
                        replaceSheets += "DBSetquery in: " + ws.Name + ","
                    End If
                Next
                ExcelDnaUtil.Application.Calculation = xlcalcmode
                ' reset the cell find dialog....
                ExcelDnaUtil.Application.ActiveSheet.Cells.Find(What:="", After:=ExcelDnaUtil.Application.ActiveSheet.Range("A1"), LookIn:=Excel.XlFindLookIn.xlFormulas, LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlNext, MatchCase:=False)
                ExcelDnaUtil.Application.DisplayAlerts = True
                ExcelDnaUtil.Application.StatusBar = False
                If replaceSheets.Length > 0 Then
                    UserMsg("Fixed possibly orphaned DB functions in active workbook from sheets: " + Left(replaceSheets, replaceSheets.Length - 1), "Orphaned DBAddin functions fix", MsgBoxStyle.Exclamation)
                End If
            End If
        Catch ex As Exception
            UserMsg("Exception occurred: " + ex.Message, "Orphaned DBAddin functions fix")
        End Try
        DBModifHelper.preventChangeWhileFetching = False
    End Sub

End Module