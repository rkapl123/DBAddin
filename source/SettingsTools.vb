Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports System.Collections.Generic
Imports System.Configuration

''' <summary>Global setting variables/functions and repairLegacyFunctions and checkpurgeNames tools</summary>
Public Module SettingsTools
    ''' <summary>currently selected environment for DB Functions, zero based (env -1) !!</summary>
    Public selectedEnvironment As Integer
    ''' <summary>environment definitions</summary>
    Public environdefs As String()
    ''' <summary>Default ConnectionString, if no connection string is given by user....</summary>
    Public ConstConnString As String
    ''' <summary>global connection timeout (can't be set in DB functions)</summary>
    Public CnnTimeout As Integer
    ''' <summary>global command timeout (can't be set in DB functions)</summary>
    Public CmdTimeout As Integer
    ''' <summary>default formatting style used in DBDate</summary>
    Public DefaultDBDateFormatting As Integer
    ''' <summary>The path where the User specific settings (overrides) can be found</summary>
    Private UserSettingsPath As String

    ''' <summary>exception proof fetching of integer settings</summary>
    ''' <param name="Key">settings key to take value from</param>
    ''' <param name="defaultValue">Value that is taken if Key was not found</param>
    ''' <returns>the setting value</returns>
    Public Function fetchSettingInt(Key As String, defaultValue As String) As Integer
        fetchSettingInt = 0
        ' catch invalid boolean expression (e.g. empty string) -> false
        Try : fetchSettingInt = CInt(fetchSetting(Key, defaultValue)) : Catch ex As Exception : End Try
        Return fetchSettingInt
    End Function

    ''' <summary>exception proof fetching of boolean settings</summary>    
    ''' <param name="Key">settings key to take value from</param>
    ''' <param name="defaultValue">Value that is taken if Key was not found</param>
    ''' <returns>the setting value</returns>
    Public Function fetchSettingBool(Key As String, defaultValue As String) As Boolean
        fetchSettingBool = False
        ' catch invalid boolean expression (e.g. empty string) -> false
        Try : fetchSettingBool = CBool(fetchSetting(Key, defaultValue)) : Catch ex As Exception : End Try
        Return fetchSettingBool
    End Function

    ''' <summary>encapsulates setting fetching (currently ConfigurationManager from DBAddin.xll.config), use only for strings. For Integer and Boolean use fetchSettingInt and fetchSettingBool</summary>
    ''' <param name="Key">settings key to take value from</param>
    ''' <param name="defaultValue">Value that is taken if Key was not found</param>
    ''' <returns>the setting value</returns>
    Public Function fetchSetting(Key As String, defaultValue As String) As String
        Dim UserSettings As Collections.Specialized.NameValueCollection = Nothing
        Dim AddinAppSettings As Collections.Specialized.NameValueCollection = Nothing
        Try : UserSettings = ConfigurationManager.GetSection("UserSettings") : Catch ex As Exception : LogWarn("Error reading UserSettings: " + ex.Message) : End Try
        Try : AddinAppSettings = ConfigurationManager.AppSettings : Catch ex As Exception : LogWarn("Error reading AppSettings: " + ex.Message) : End Try
        ' user specific settings are in UserSettings section in separate file
        If IsNothing(UserSettings) OrElse IsNothing(UserSettings(Key)) Then
            If Not IsNothing(AddinAppSettings) Then
                fetchSetting = AddinAppSettings(Key)
            Else
                fetchSetting = Nothing
            End If
        ElseIf Not (IsNothing(UserSettings) OrElse IsNothing(UserSettings(Key))) Then
            fetchSetting = UserSettings(Key)
        Else
            fetchSetting = Nothing
        End If
        ' rough type check based on default value
        If defaultValue <> "" And fetchSetting <> "" Then
            Dim checkDefaultInt As Integer = 0
            Dim checkDefaultBool As Boolean = False
            If Integer.TryParse(defaultValue, checkDefaultInt) AndAlso Not Integer.TryParse(fetchSetting, checkDefaultInt) Then
                UserMsg("couldn't parse the setting " + Key + " as an Integer: " + fetchSetting + ", using default value: " + defaultValue)
                fetchSetting = Nothing
            ElseIf Boolean.TryParse(defaultValue, checkDefaultBool) AndAlso Not Boolean.TryParse(fetchSetting, checkDefaultBool) Then
                UserMsg("couldn't parse the setting " + Key + " as a Boolean: " + fetchSetting + ", using default value: " + defaultValue)
                fetchSetting = Nothing
            End If
        End If
        If fetchSetting Is Nothing Then fetchSetting = defaultValue
    End Function

    ''' <summary>change or add a key/value pair in the user settings</summary>
    ''' <param name="theKey">key to change (or add)</param>
    ''' <param name="theValue">value for key</param>
    Public Sub setUserSetting(theKey As String, theValue As String)
        ' check if key exists
        Dim doc As New Xml.XmlDocument()
        doc.Load(UserSettingsPath)
        Dim keyNode As Xml.XmlNode = doc.SelectSingleNode("/UserSettings/add[@key='" + System.Security.SecurityElement.Escape(theKey) + "']")
        If IsNothing(keyNode) Then
            ' if not, add to settings
            Dim nodeRegion As Xml.XmlElement = doc.CreateElement("add")
            nodeRegion.SetAttribute("key", theKey)
            nodeRegion.SetAttribute("value", theValue)
            doc.SelectSingleNode("//UserSettings").AppendChild(nodeRegion)
        Else
            keyNode.Attributes().GetNamedItem("value").InnerText = theValue
        End If
        doc.Save(UserSettingsPath)
        ConfigurationManager.RefreshSection("UserSettings")
    End Sub

    ''' <summary>environment for settings (+1 of selected Environment which is the index of the dropdown, if baseZero is set then simply the index)</summary>
    ''' <returns></returns>
    Public Function env(Optional baseZero As Boolean = False) As String
        Return (selectedEnvironment + IIf(baseZero, 0, 1)).ToString()
    End Function

    Public refreshDataKey, jumpButtonKey, deleteRowKey, insertRowKey As String
    ''' <summary>initializes global configuration variables</summary>
    Public Sub initSettings()
        Try
            DebugAddin = fetchSettingBool("DebugAddin", "False")
            ConstConnString = fetchSetting("ConstConnString" + env(), "")
            CnnTimeout = fetchSettingInt("CnnTimeout", "15")
            CmdTimeout = fetchSettingInt("CmdTimeout", "60")
            ConfigStoreFolder = fetchSetting("ConfigStoreFolder" + env(), "")
            specialConfigStoreFolders = Split(fetchSetting("specialConfigStoreFolders", ""), ":")
            DefaultDBDateFormatting = fetchSettingInt("DefaultDBDateFormatting", "0")
            ' load environments
            Dim i As Integer = 1
            ReDim Preserve environdefs(-1)
            Dim ConfigName As String
            Do
                ConfigName = fetchSetting("ConfigName" + i.ToString(), vbNullString)
                If Len(ConfigName) > 0 Then
                    ReDim Preserve environdefs(environdefs.Length)
                    environdefs(environdefs.Length - 1) = ConfigName + " - " + i.ToString()
                End If
                i += 1
            Loop Until Len(ConfigName) = 0
        Catch ex As Exception
            UserMsg("Error in initialization of Settings: " + ex.Message)
        End Try
        ' overridable shortcuts, first reset
        Try : ExcelDnaUtil.Application.OnKey(refreshDataKey) : Catch ex As Exception : End Try
        Try : ExcelDnaUtil.Application.OnKey(jumpButtonKey) : Catch ex As Exception : End Try
        Try : ExcelDnaUtil.Application.OnKey(deleteRowKey) : Catch ex As Exception : End Try
        Try : ExcelDnaUtil.Application.OnKey(insertRowKey) : Catch ex As Exception : End Try
        refreshDataKey = fetchSetting("shortCutRefreshData", "^R")
        jumpButtonKey = fetchSetting("shortCutJumpButton", "^J")
        deleteRowKey = fetchSetting("shortCutDeleteRow", "^D")
        insertRowKey = fetchSetting("shortCutInsertRow", "^I")
        ' then set to (new) values:
        Try : ExcelDnaUtil.Application.OnKey(refreshDataKey, "refreshData") : Catch ex As Exception : End Try
        Try : ExcelDnaUtil.Application.OnKey(jumpButtonKey, "jumpButton") : Catch ex As Exception : End Try
        Try : ExcelDnaUtil.Application.OnKey(deleteRowKey, "deleteRow") : Catch ex As Exception : End Try
        Try : ExcelDnaUtil.Application.OnKey(insertRowKey, "insertRow") : Catch ex As Exception : End Try
        ' get module info for path of xll (to get config there):
        For Each tModule As Diagnostics.ProcessModule In Diagnostics.Process.GetCurrentProcess().Modules
            UserSettingsPath = tModule.FileName
            If UserSettingsPath.ToUpper.Contains("DBADDIN") Then
                UserSettingsPath = Replace(UserSettingsPath, ".xll", "User.config")
                Exit For
            End If
        Next
    End Sub

    ''' <summary>resets the caches for given workbook</summary>
    ''' <param name="WBname"></param>
    Public Sub resetCachesForWorkbook(WBname As String)
        ' reset query cache for current workbook, so we really get new data !
        Dim tempColl1 As New Dictionary(Of String, String)(queryCache) ' clone dictionary to be able to remove items...
        For Each resetkey As String In tempColl1.Keys
            If InStr(resetkey, "[" + WBname + "]") > 0 Then queryCache.Remove(resetkey)
        Next
        Dim tempColl2 As New Dictionary(Of String, ContainedStatusMsg)(StatusCollection)
        For Each resetkey As String In tempColl2.Keys
            If InStr(resetkey, "[" + WBname + "]") > 0 Then StatusCollection.Remove(resetkey)
        Next
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''' Supporting Tools ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
            LogWarn("Exception when trying to get Workbook names: " + ex.Message + ", this might be due to errors in the VBA Macros (missing references)")
            Exit Sub
        End Try
        If actWB Is Nothing Then
            ' only log warning, no user message !
            LogWarn("no active workbook available !")
            Exit Sub
        End If
        DBModifs.preventChangeWhileFetching = True ' WorksheetFunction.CountIf triggers Change event with target in argument 1, so make sure this doesn't trigger anything inside DBAddin)
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
        DBModifs.preventChangeWhileFetching = False
    End Sub

    ''' <summary>maintenance procedure to check/purge names used for db-functions from workbook, or unhide DB names</summary>
    Public Sub checkpurgeNames()
        Dim actWbNames As Excel.Names = Nothing
        Try : actWbNames = ExcelDnaUtil.Application.ActiveWorkbook.Names : Catch ex As Exception
            UserMsg("Exception when trying to get the active workbook names for purging names: " + ex.Message + ", this might be either due to errors in the VBA-IDE (missing references) or due to opening this workbook from an MS-Office hyperlink, starting up Excel (timing issue). Switch to another workbook and back to fix.")
            Exit Sub
        End Try
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
                    For Each ws As Excel.Worksheet In ExcelDnaUtil.Application.ActiveWorkbook.Worksheets
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
                        collectedErrors += DBname.Name + "' doesn't have a corresponding DBFsource name" + vbCrLf
                    End If
                    If InStr(DBname.RefersTo, "#REF!") > 0 Then
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
                            ' do nothing, already collected...
                        Else
                            collectedErrors += DBname.Name + "' RefersToRange resulted in Exception " + ex.Message + vbCrLf
                        End If
                    End Try
                    If DBname.Visible Then
                        collectedErrors += DBname.Name + "' is visible" + vbCrLf
                    End If
                End If
                If DBname.Name Like "*DBFsource*" Then
                    Try : checkExists = NamesList.Item(Replace(DBname.Name, "DBFsource", "DBFtarget")) : Catch ex As Exception : End Try
                    If IsNothing(checkExists) Then
                        collectedErrors += DBname.Name + "' doesn't have a corresponding DBFtarget name" + vbCrLf
                    End If
                    If InStr(DBname.RefersTo, "#REF!") > 0 Then
                        collectedErrors += DBname.Name + "' contains #REF!" + vbCrLf
                    End If
                    If DBname.Visible Then
                        collectedErrors += DBname.Name + "' is visible" + vbCrLf
                    End If
                End If
            Next
            If collectedErrors = "" Then
                UserMsg("No Problems detected.", "DBfunction check Error", MsgBoxStyle.Information)
            Else
                UserMsg(collectedErrors, "DBfunction check Error")
            End If
            ' last check any possible DB Modifier Definitions for validity
            getDBModifDefinitions(True)
        End If
    End Sub

End Module