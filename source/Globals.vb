Imports ExcelDna.Integration
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports System.Collections.Generic
Imports System.Configuration
Imports System.Diagnostics

''' <summary>Global variables and functions for DB Addin</summary>
Public Module Globals
    ' general Global objects/variables
    ''' <summary>ribbon menu handler</summary>
    Public theMenuHandler As MenuHandler
    ''' <summary>currently selected environment for DB Functions, zero based (env -1) !!</summary>
    Public selectedEnvironment As Integer
    ''' <summary>reference object for the Addins ribbon</summary>
    Public theRibbon As CustomUI.IRibbonUI
    ''' <summary>environment definitions</summary>
    Public environdefs As String()
    ''' <summary>DBModif definition collections of DBmodif types (key of top level dictionary) with values being collections of DBModifierNames (key of contained dictionaries) and DBModifiers (value of contained dictionaries))</summary>
    Public DBModifDefColl As Dictionary(Of String, Dictionary(Of String, DBModif))

    ''' <summary>for DBMapper invocations by execDBModif, this is set to true, avoiding MsgBox</summary>
    Public nonInteractive As Boolean = False
    ''' <summary>collect non interactive error messages here</summary>
    Public nonInteractiveErrMsgs As String
    ''' <summary>set to true if warning was issued, this flag indicates that the log button should get an exclamation sign</summary>
    Public WarningIssued As Boolean
    ''' <summary>the Textfile log source</summary>
    Public theLogFileSource As TraceSource
    ''' <summary>the LogDisplay (Diagnostic Display) log source</summary>
    Public theLogDisplaySource As TraceSource

    ' Global settings
    ''' <summary>Debug the Addin: write trace messages</summary>
    Public DebugAddin As Boolean
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

    ''' <summary>encapsulates setting fetching (currently ConfigurationManager from DBAddin.xll.config)</summary>
    ''' <param name="Key">registry key to take value from</param>
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
                Globals.UserMsg("couldn't parse the setting " + Key + " as an Integer: " + fetchSetting + ", using default value: " + defaultValue)
                fetchSetting = Nothing
            ElseIf Boolean.TryParse(defaultValue, checkDefaultBool) AndAlso Not Boolean.TryParse(fetchSetting, checkDefaultBool) Then
                Globals.UserMsg("couldn't parse the setting " + Key + " as a Boolean: " + fetchSetting + ", using default value: " + defaultValue)
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
        Return (Globals.selectedEnvironment + IIf(baseZero, 0, 1)).ToString()
    End Function

    ''' <summary>initializes global configuration variables</summary>
    Public Sub initSettings()
        Try
            DebugAddin = CBool(fetchSetting("DebugAddin", "False"))
            ConstConnString = fetchSetting("ConstConnString" + Globals.env(), "")
            CnnTimeout = CInt(fetchSetting("CnnTimeout", "15"))
            CmdTimeout = CInt(fetchSetting("CmdTimeout", "60"))
            ConfigStoreFolder = fetchSetting("ConfigStoreFolder" + Globals.env(), "")
            specialConfigStoreFolders = Split(fetchSetting("specialConfigStoreFolders", ""), ":")
            DefaultDBDateFormatting = CInt(fetchSetting("DefaultDBDateFormatting", "0"))
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
        ' get module info for path of xll (to get config there):
        For Each tModule As Diagnostics.ProcessModule In Diagnostics.Process.GetCurrentProcess().Modules
            UserSettingsPath = tModule.FileName
            If UserSettingsPath.ToUpper.Contains("DBADDIN") Then
                UserSettingsPath = Replace(UserSettingsPath, ".xll", "User.config")
                Exit For
            End If
        Next
    End Sub

    ''' <summary>Logs Message of eEventType to System.Diagnostics.Trace</summary>
    ''' <param name="Message">Message to be logged</param>
    ''' <param name="eEventType">event type: info, warning, error</param>
    ''' <param name="caller">reflection based caller information: module.method</param>
    Private Sub WriteToLog(Message As String, eEventType As TraceEventType, caller As String)
        ' collect errors and warnings for returning messages in executeDBModif
        If eEventType = TraceEventType.Error Or eEventType = TraceEventType.Warning Then nonInteractiveErrMsgs += caller + ":" + Message + vbCrLf

        Dim timestamp As Int32 = DateAndTime.Now().Month * 100000000 + DateAndTime.Now().Day * 1000000 + DateAndTime.Now().Hour * 10000 + DateAndTime.Now().Minute * 100 + DateAndTime.Now().Second
        If nonInteractive Then
            theLogDisplaySource.TraceEvent(TraceEventType.Information, timestamp, "Non-interactive: {0}: {1}", caller, Message)
            theLogFileSource.TraceEvent(TraceEventType.Information, timestamp, "Non-interactive: {0}: {1}", caller, Message)
        Else
            Select Case eEventType
                Case TraceEventType.Information
                    theLogDisplaySource.TraceEvent(TraceEventType.Information, timestamp, "{0}: {1}", caller, Message)
                    theLogFileSource.TraceEvent(TraceEventType.Information, timestamp, "{0}: {1}", caller, Message)
                Case TraceEventType.Warning
                    theLogDisplaySource.TraceEvent(TraceEventType.Warning, timestamp, "{0}: {1}", caller, Message)
                    theLogFileSource.TraceEvent(TraceEventType.Warning, timestamp, "{0}: {1}", caller, Message)
                    WarningIssued = True
                    ' at Addin Start ribbon has not been loaded so avoid call to it here..
                    If theRibbon IsNot Nothing Then theRibbon.InvalidateControl("showLog")
                Case TraceEventType.Error
                    theLogDisplaySource.TraceEvent(TraceEventType.Error, timestamp, "{0}: {1}", caller, Message)
                    theLogFileSource.TraceEvent(TraceEventType.Error, timestamp, "{0}: {1}", caller, Message)
                    WarningIssued = True
                    If theRibbon IsNot Nothing Then theRibbon.InvalidateControl("showLog")
            End Select
        End If
    End Sub

    ''' <summary>Logs error messages</summary>
    ''' <param name="LogMessage">the message to be logged</param>
    Public Sub LogError(LogMessage As String)
        Dim caller As String
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Try : caller = theMethod.ReflectedType.FullName + "." + theMethod.Name : Catch ex As Exception : caller = theMethod.Name : End Try
        WriteToLog(LogMessage, TraceEventType.Error, caller)
    End Sub

    ''' <summary>Logs warning messages</summary>
    ''' <param name="LogMessage">the message to be logged</param>
    Public Sub LogWarn(LogMessage As String)
        Dim caller As String
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Try : caller = theMethod.ReflectedType.FullName + "." + theMethod.Name : Catch ex As Exception : caller = theMethod.Name : End Try
        WriteToLog(LogMessage, TraceEventType.Warning, caller)
    End Sub

    ''' <summary>Logs informational messages</summary>
    ''' <param name="LogMessage">the message to be logged</param>
    Public Sub LogInfo(LogMessage As String)
        If DebugAddin Then
            Dim caller As String
            Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
            Try : caller = theMethod.ReflectedType.FullName + "." + theMethod.Name : Catch ex As Exception : caller = theMethod.Name : End Try
            WriteToLog(LogMessage, TraceEventType.Information, caller)
        End If
    End Sub

    ''' <summary>show message to User (default Error message) and log as warning if Critical Or Exclamation (logged errors would pop up the trace information window)</summary> 
    ''' <param name="LogMessage">the message to be shown/logged</param>
    ''' <param name="errTitle">optionally pass a title for the msgbox instead of default DBAddin Error</param>
    ''' <param name="msgboxIcon">optionally pass a different Msgbox icon (style) instead of default MsgBoxStyle.Critical</param>
    Public Sub UserMsg(LogMessage As String, Optional errTitle As String = "DBAddin Error", Optional msgboxIcon As MsgBoxStyle = MsgBoxStyle.Critical)
        Dim caller As String
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Try : caller = theMethod.ReflectedType.FullName + "." + theMethod.Name : Catch ex As Exception : caller = theMethod.Name : End Try
        WriteToLog(LogMessage, If(msgboxIcon = MsgBoxStyle.Critical Or msgboxIcon = MsgBoxStyle.Exclamation, TraceEventType.Warning, TraceEventType.Information), caller) ' to avoid popup of trace log in nonInteractive mode...
        If Not nonInteractive Then
            MsgBox(LogMessage, msgboxIcon + MsgBoxStyle.OkOnly, errTitle)
            ' avoid activation of ribbon in AutoOpen as this throws an exception (ribbon is not assigned until AutoOpen has finished)
            If theRibbon IsNot Nothing Then theRibbon.ActivateTab("DBaddinTab")
        End If
    End Sub

    ''' <summary>ask User (default OKCancel) and log as warning if Critical Or Exclamation (logged errors would pop up the trace information window)</summary> 
    ''' <param name="theMessage">the question to be shown/logged</param>
    ''' <param name="questionType">optionally pass question box type, default MsgBoxStyle.OKCancel</param>
    ''' <param name="questionTitle">optionally pass a title for the msgbox instead of default DBAddin Question</param>
    ''' <param name="msgboxIcon">optionally pass a different Msgbox icon (style) instead of default MsgBoxStyle.Question</param>
    ''' <returns>choice as MsgBoxResult (Yes, No, OK, Cancel...)</returns>
    Public Function QuestionMsg(theMessage As String, Optional questionType As MsgBoxStyle = MsgBoxStyle.OkCancel, Optional questionTitle As String = "DBAddin Question", Optional msgboxIcon As MsgBoxStyle = MsgBoxStyle.Question) As MsgBoxResult
        Dim caller As String
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Try : caller = theMethod.ReflectedType.FullName + "." + theMethod.Name : Catch ex As Exception : caller = theMethod.Name : End Try
        WriteToLog(theMessage, If(msgboxIcon = MsgBoxStyle.Critical Or msgboxIcon = MsgBoxStyle.Exclamation, TraceEventType.Warning, TraceEventType.Information), caller) ' to avoid popup of trace log
        If nonInteractive Then
            If questionType = MsgBoxStyle.OkCancel Then Return MsgBoxResult.Cancel
            If questionType = MsgBoxStyle.YesNo Then Return MsgBoxResult.No
            If questionType = MsgBoxStyle.YesNoCancel Then Return MsgBoxResult.No
            If questionType = MsgBoxStyle.RetryCancel Then Return MsgBoxResult.Cancel
        End If
        ' tab is not activated BEFORE Msgbox as Excel first has to get into the interaction thread outside this one..
        If theRibbon IsNot Nothing Then theRibbon.ActivateTab("DBaddinTab")
        Return MsgBox(theMessage, msgboxIcon + questionType, questionTitle)
    End Function

    ''' <summary>refresh DB Functions (and - if called from outside any db function area - all other external data ranges)</summary>
    <ExcelCommand(Name:="refreshData", ShortCut:="^R")>
    Public Sub refreshData()
        initSettings()
        ' enable events in case there were some problems in procedure with EnableEvents = false, this fails if a cell dropdown is open.
        Try
            ExcelDnaUtil.Application.EnableEvents = True
        Catch ex As Exception
            UserMsg("Can't refresh data while lookup dropdown is open !!", "DB-Addin Refresh")
            Exit Sub
        End Try
        Dim actWb As Excel.Workbook = Nothing
        Try : actWb = ExcelDnaUtil.Application.ActiveWorkbook : Catch ex As Exception : End Try
        If IsNothing(actWb) Then
            UserMsg("Couldn't get active workbook for refreshing data, this might be due to the active workbook being hidden, Errors in VBA Macros or missing references", "DB-Addin Refresh")
            Exit Sub
        End If
        Try : Dim actWbName As String = actWb.Name : Catch ex As Exception
            UserMsg("Couldn't get active workbook name for refreshing data, this might be due to Errors in VBA Macros or missing references", "DB-Addin Refresh")
            Exit Sub
        End Try
        If ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual Then
            UserMsg("Calculation is set to manual, this prevents DB Functions from being recalculated. Please set calculation to automatic and retry", "DB-Addin Refresh")
            Exit Sub
        End If
        ' also reset the database connection in case of errors (might be nothing or not open...)
        Try : conn.Close() : Catch ex As Exception : End Try
        conn = Nothing
        dontTryConnection = False
        Try
            ' look for old query caches and status collections (returned error messages) in active workbook and reset them to get new data
            resetCachesForWorkbook(actWb.Name)
            Dim underlyingName As String = getUnderlyingDBNameFromRange(ExcelDnaUtil.Application.ActiveCell)
            ' now for DBListfetch/DBRowfetch resetting, either called outside of all db function areas...
            If underlyingName = "" Then
                refreshDBFunctions(actWb)
                ' general refresh: also refresh all embedded queries, pivot tables and list objects..
                Try
                    Dim ws As Excel.Worksheet
                    For Each ws In actWb.Worksheets
                        If ws.ProtectContents And (ws.QueryTables.Count > 0 Or ws.PivotTables.Count > 0) Then
                            UserMsg("Worksheet " + ws.Name + " is content protected, can't refresh QueryTables/PivotTables !", "DB-Addin Refresh")
                            Continue For
                        End If
                        If Not CBool(fetchSetting("AvoidUpdateQueryTables_Refresh", "False")) Then
                            For Each qrytbl As Excel.QueryTable In ws.QueryTables
                                ' no need to avoid double refreshing as query tables are always deleted after inserting data in DBListFetchAction
                                qrytbl.Refresh()
                            Next
                        End If
                        If Not CBool(fetchSetting("AvoidUpdatePivotTables_Refresh", "False")) Then
                            For Each pivottbl As Excel.PivotTable In ws.PivotTables
                                ' avoid double refreshing of dbsetquery list objects
                                If getUnderlyingDBNameFromRange(pivottbl.TableRange1) = "" Then pivottbl.PivotCache.Refresh()
                            Next
                        End If
                        If Not CBool(fetchSetting("AvoidUpdateListObjects_Refresh", "False")) Then
                            For Each listobj As Excel.ListObject In ws.ListObjects
                                ' avoid double refreshing of dbsetquery list objects
                                If getUnderlyingDBNameFromRange(listobj.Range) = "" Then listobj.QueryTable.Refresh()
                            Next
                        End If
                    Next
                    If Not CBool(fetchSetting("AvoidUpdateLinks_Refresh", "False")) Then actWb.UpdateLink(Name:=actWb.LinkSources, Type:=Excel.XlLink.xlExcelLinks)
                Catch ex As Exception
                End Try
            Else ' or called inside a db function area (target or source = function cell)
                If Left$(underlyingName, 10) = "DBFtargetF" Then
                    underlyingName = Replace(underlyingName, "DBFtargetF", "DBFsource", 1, , vbTextCompare)
                    If ExcelDnaUtil.Application.Range(underlyingName).Parent.ProtectContents Then
                        UserMsg("Worksheet " + ExcelDnaUtil.Application.Range(underlyingName).Parent.Name + " is content protected, can't refresh DB Function !", "DB-Addin Refresh")
                        Exit Sub
                    End If
                    ExcelDnaUtil.Application.Range(underlyingName).Formula += " "
                    ' we're being called on a target area
                ElseIf Left$(underlyingName, 9) = "DBFtarget" Then
                    underlyingName = Replace(underlyingName, "DBFtarget", "DBFsource", 1, , vbTextCompare)
                    If ExcelDnaUtil.Application.Range(underlyingName).Parent.ProtectContents Then
                        UserMsg("Worksheet " + ExcelDnaUtil.Application.Range(underlyingName).Parent.Name + " is content protected, can't refresh DB Function !", "DB-Addin Refresh")
                        Exit Sub
                    End If
                    ExcelDnaUtil.Application.Range(underlyingName).Formula += " "
                    ' we're being called on a source (invoking function) cell
                ElseIf Left$(underlyingName, 9) = "DBFsource" Then
                    If ExcelDnaUtil.Application.Range(underlyingName).Parent.ProtectContents Then
                        UserMsg("Worksheet " + ExcelDnaUtil.Application.Range(underlyingName).Parent.Name + " is content protected, can't refresh DB Function !", "DB-Addin Refresh")
                        Exit Sub
                    End If
                    ExcelDnaUtil.Application.Range(underlyingName).Formula += " "
                Else
                    UserMsg("Error in refreshData, underlyingName does not begin with DBFtarget, DBFtargetF or DBFsource: " + underlyingName, "DB-Addin Refresh")
                End If
            End If
        Catch ex As Exception
            UserMsg("Exception: " + ex.Message, "refresh Data", "DB-Addin Refresh")
        End Try
    End Sub

    ''' <summary>jumps between DB Function and target area</summary>
    <ExcelCommand(Name:="jumpButton", ShortCut:="^J")>
    Public Sub jumpButton()
        If checkMultipleDBRangeNames(ExcelDnaUtil.Application.ActiveCell) Then
            UserMsg("Multiple hidden DB Function names in selected cell (making 'jump' ambiguous/impossible), please use purge names tool!", "DB-Addin Jump")
            Exit Sub
        End If
        Dim underlyingName As String = getUnderlyingDBNameFromRange(ExcelDnaUtil.Application.ActiveCell)
        If underlyingName = "" Then Exit Sub
        If Left$(underlyingName, 10) = "DBFtargetF" Then
            underlyingName = Replace(underlyingName, "DBFtargetF", "DBFsource", 1, , vbTextCompare)
        ElseIf Left$(underlyingName, 9) = "DBFtarget" Then
            underlyingName = Replace(underlyingName, "DBFtarget", "DBFsource", 1, , vbTextCompare)
        Else
            underlyingName = Replace(underlyingName, "DBFsource", "DBFtarget", 1, , vbTextCompare)
        End If
        Try
            ExcelDnaUtil.Application.Range(underlyingName).Parent.Select()
            ExcelDnaUtil.Application.Range(underlyingName).Select()
        Catch ex As Exception
            UserMsg("Can't jump to target/source, corresponding workbook open? " + ex.Message, "DB-Addin Jump")
        End Try
    End Sub

    ''' <summary>splits theString into tokens delimited by delimiter, ignoring delimiters inside quotes and brackets</summary>
    ''' <param name="theString">string to be split into tokens, case insensitive !</param>
    ''' <param name="delimiter">delimiter that string is to be split by</param>
    ''' <param name="quote">quote character where delimiters should be ignored inside</param>
    ''' <param name="startStr">part of theString where splitting should start after, case insensitive !</param>
    ''' <param name="openBracket">opening bracket character</param>
    ''' <param name="closeBracket">closing bracket character</param>
    ''' <returns>the list of tokens</returns>
    ''' <remarks>theString is split starting from startStr up to the first balancing closing Bracket (as defined by openBracket and closeBracket)
    ''' startStr, openBracket and closeBracket are case insensitive for comparing with theString.
    ''' the tokens are not blank trimmed !!</remarks>
    Public Function functionSplit(ByVal theString As String, delimiter As String, quote As String, startStr As String, openBracket As String, closeBracket As String) As Object
        Dim tempString As String
        Dim finalResult
        Try
            ' find startStr
            tempString = Mid$(theString, InStr(1, UCase$(theString), UCase$(startStr)) + Len(startStr))
            ' rip out the balancing string now...
            tempString = balancedString(tempString, openBracket, closeBracket, quote)
            If tempString.Length = 0 Then
                UserMsg("couldn't produce balanced string from " + theString)
                functionSplit = Nothing
                Exit Function
            End If
            tempString = replaceDelimsWithSpecialSep(tempString, delimiter, quote, openBracket, closeBracket, vbTab)
            finalResult = Split(tempString, vbTab)
            functionSplit = finalResult
        Catch ex As Exception
            UserMsg("Exception: " + ex.Message, "function split into tokens")
            functionSplit = Nothing
        End Try
    End Function

    ''' <summary>returns the minimal bracket balancing string contained in theString, opening bracket defined in openBracket, closing bracket defined in closeBracket
    ''' disregarding quoted areas inside optionally given quote character/string</summary>
    ''' <param name="theString"></param>
    ''' <param name="openBracket"></param>
    ''' <param name="closeBracket"></param>
    ''' <param name="quote"></param>
    ''' <returns>the balanced string</returns>
    Public Function balancedString(theString As String, openBracket As String, closeBracket As String, Optional quote As String = "") As String
        Dim startBalance As Long, endBalance As Long, i As Long, countOpen As Long, countClose As Long
        balancedString = ""
        Dim quoteMode As Boolean = False
        Try
            startBalance = 0
            For i = 1 To Len(theString)
                If Left$(Mid$(theString, i), Len(quote)) = quote And quote.Length > 0 And Not quoteMode Then
                    quoteMode = True
                Else
                    If Not quoteMode Then
                        If Left$(Mid$(theString, i), Len(openBracket)) = openBracket Then
                            If startBalance = 0 Then startBalance = i
                            countOpen += 1
                        End If
                        If startBalance <> 0 And Left$(Mid$(theString, i), Len(closeBracket)) = closeBracket Then countClose += 1
                    Else
                        If Left$(Mid$(theString, i), Len(quote)) = quote And quote.Length > 0 Then quoteMode = False
                    End If
                End If

                If countOpen = countClose And startBalance <> 0 Then
                    endBalance = i - 1
                    Exit For
                End If
            Next
            If endBalance <> 0 Then
                balancedString = Mid$(theString, startBalance + 1, endBalance - startBalance)
            End If
        Catch ex As Exception
            UserMsg("Exception: " + ex.Message, "get minimal balanced string")
        End Try
    End Function

    ''' <summary>replaces the delimiter (delimiter) inside theString with specialSep, regarding both quoted areas inside quote and bracketed areas (inside openBracket/closeBracket)</summary>
    ''' <param name="theString"></param>
    ''' <param name="delimiter"></param>
    ''' <param name="quote"></param>
    ''' <param name="openBracket"></param>
    ''' <param name="closeBracket"></param>
    ''' <param name="specialSep"></param>
    ''' <returns>replaced string</returns>
    Public Function replaceDelimsWithSpecialSep(theString As String, delimiter As String, quote As String, openBracket As String, closeBracket As String, specialSep As String) As String
        Dim openedBrackets As Long, quoteMode As Boolean
        Dim i As Long
        replaceDelimsWithSpecialSep = ""
        Try
            For i = 1 To Len(theString)
                If Left$(Mid$(theString, i), Len(quote)) = quote And quote.Length > 0 And Not quoteMode Then
                    quoteMode = True
                Else
                    If quoteMode And Left$(Mid$(theString, i), Len(quote)) = quote And quote.Length > 0 Then quoteMode = False
                End If

                If Left$(Mid$(theString, i), Len(openBracket)) = openBracket And openBracket.Length > 0 And Not quoteMode Then
                    openedBrackets += 1
                End If
                If Left$(Mid$(theString, i), Len(closeBracket)) = closeBracket And closeBracket.Length > 0 And Not quoteMode Then
                    openedBrackets -= 1
                End If

                If Not (openedBrackets > 0 Or quoteMode) Then
                    If Left$(Mid$(theString, i), Len(delimiter)) = delimiter Then
                        replaceDelimsWithSpecialSep += specialSep
                    Else
                        replaceDelimsWithSpecialSep += Mid$(theString, i, 1)
                    End If
                Else
                    replaceDelimsWithSpecialSep += Mid$(theString, i, 1)
                End If
            Next
        Catch ex As Exception
            UserMsg("Exception: " + ex.Message, "replace delimiters with special separator")
        End Try
    End Function

    ''' <summary>changes theString to changedString by replacing substring starting AFTER keystr and ending with separator (so "(keystr)...;" will become "(keystr)(changedString);", case insensitive !!</summary>
    ''' <param name="theString"></param>
    ''' <param name="keystr"></param>
    ''' <param name="changedString"></param>
    ''' <param name="separator"></param>
    ''' <returns>the changed string</returns>
    Public Function Change(ByVal theString As String, ByVal keystr As String, ByVal changedString As String, ByVal separator As String) As String
        Dim replaceBeg, replaceEnd As Integer

        replaceBeg = InStr(1, UCase$(theString), UCase$(keystr))
        If replaceBeg = 0 Then
            Change = theString
            Exit Function
        End If
        replaceEnd = InStr(replaceBeg, UCase$(theString), UCase$(separator))
        If replaceEnd = 0 Then replaceEnd = Len(theString) + 1
        Change = Left$(theString, replaceBeg - 1 + Len(keystr)) + changedString + Right$(theString, Len(theString) - replaceEnd + 1)
    End Function

    ''' <summary>fetches substring starting after keystr and ending with separator from theString, case insensitive !! if separator is "" then fetch to end of string</summary>
    ''' <param name="theString">string to be searched</param>
    ''' <param name="keystr">string indicating the start of the substring combination</param>
    ''' <param name="separator">string ending the whole substring, not included in returned string!</param>
    ''' <param name="includeKeyStr">if includeKeyStr is set to true, include keystr in returned string</param>
    ''' <returns>the fetched substring</returns>
    Public Function fetchSubstr(ByVal theString As String, ByVal keystr As String, ByVal separator As String, Optional includeKeyStr As Boolean = False) As String
        Dim fetchBeg As Integer, fetchEnd As Integer

        fetchBeg = InStr(1, UCase$(theString), UCase$(keystr))
        If fetchBeg = 0 Then Return ""
        fetchEnd = InStr(fetchBeg + Len(keystr), UCase$(theString), UCase$(separator))
        If fetchEnd = 0 Or separator.Length = 0 Then fetchEnd = Len(theString) + 1
        fetchSubstr = Mid$(theString, fetchBeg + IIf(includeKeyStr, 0, Len(keystr)), fetchEnd - (fetchBeg + IIf(includeKeyStr, 0, Len(keystr))))
    End Function

    ''' <summary>checks whether worksheet called theName exists</summary>
    ''' <param name="theName"></param>
    ''' <returns>True if sheet exists</returns>
    Public Function existsSheet(ByRef theName As String) As Boolean
        existsSheet = True
        Try
            Dim dummy As String = ExcelDnaUtil.Application.Worksheets(theName).name
        Catch ex As Exception
            existsSheet = False
        End Try
    End Function

    ''' <summary>helper function for check whether name exists in workbook</summary>
    ''' <param name="CheckForName">name to be checked</param>
    ''' <returns>true if name exists</returns>
    Public Function existsName(CheckForName As String) As Boolean
        existsName = False
        On Error GoTo Last
        If Len(ExcelDnaUtil.Application.ActiveWorkbook.Names(CheckForName).Name) <> 0 Then existsName = True
Last:
    End Function

    ''' <summary>gets underlying DBtarget/DBsource Name from theRange</summary>
    ''' <param name="theRange"></param>
    ''' <returns>the retrieved name</returns>
    Public Function getUnderlyingDBNameFromRange(theRange As Excel.Range) As String
        Dim nm As Excel.Name
        Dim rng, testRng As Excel.Range

        getUnderlyingDBNameFromRange = ""
        Try
            For Each nm In theRange.Parent.Parent.Names
                rng = Nothing
                Try : rng = nm.RefersToRange : Catch ex As Exception : End Try
                If rng IsNot Nothing Then
                    testRng = Nothing
                    Try : testRng = ExcelDnaUtil.Application.Intersect(theRange, rng) : Catch ex As Exception : End Try
                    If testRng IsNot Nothing And (InStr(nm.Name, "DBFtarget") > 0 Or InStr(nm.Name, "DBFsource") > 0) Then
                        Dim WbkSepPos As Integer = InStr(nm.Name, "!")
                        If WbkSepPos > 1 Then
                            getUnderlyingDBNameFromRange = Mid(nm.Name, WbkSepPos + 1)
                        Else
                            getUnderlyingDBNameFromRange = nm.Name
                        End If
                        Exit Function
                    End If
                End If
            Next
        Catch ex As Exception
            UserMsg("Exception: " + ex.Message, "get underlying DBFName from Range")
        End Try
    End Function

    ''' <summary>check if multiple (hidden, containing DBtarget or DBsource) DB Function names exist in theRange</summary>
    ''' <param name="theRange"></param>
    ''' <returns>True if multiple names exist</returns>
    Public Function checkMultipleDBRangeNames(theRange As Excel.Range) As Boolean
        Dim nm As Excel.Name
        Dim rng, testRng As Excel.Range
        Dim foundNames As Integer = 0

        checkMultipleDBRangeNames = False
        Try
            For Each nm In theRange.Parent.Parent.Names
                rng = Nothing
                Try : rng = nm.RefersToRange : Catch ex As Exception : End Try
                If rng IsNot Nothing And Not (nm.Name Like "*ExterneDaten*" Or nm.Name Like "*_FilterDatabase") Then
                    testRng = Nothing
                    Try : testRng = ExcelDnaUtil.Application.Intersect(theRange, rng) : Catch ex As Exception : End Try
                    If testRng IsNot Nothing And (InStr(1, nm.Name, "DBFtarget") >= 1 Or InStr(1, nm.Name, "DBFsource") >= 1) Then
                        foundNames += 1
                    End If
                End If
            Next
            If foundNames > 1 Then checkMultipleDBRangeNames = True
        Catch ex As Exception
            UserMsg("Exception: " + ex.Message, "check Multiple DBRange Names")
        End Try
    End Function

    ''' <summary>converts a Mso Menu ID to a Drawing Image</summary>
    ''' <param name="idMso">the Mso Menu ID to be converted</param>
    ''' <returns>a System.Drawing.Image to be used by </returns>
    Public Function convertFromMso(idMso As String) As System.Drawing.Image
        Try
            Dim p As stdole.IPictureDisp = ExcelDnaUtil.Application.CommandBars.GetImageMso(idMso, 16, 16)
            Dim hPal As IntPtr = p.hPal
            convertFromMso = System.Drawing.Image.FromHbitmap(p.Handle, hPal)
        Catch ex As Exception
            ' in case above image fetching doesn't work then no image is displayed (the image parameter is still required for ContextMenuStrip.Items.Add !)
            convertFromMso = Nothing
        End Try
    End Function

    ''' <summary>recalculate fully the DB functions, if we have DBFuncs in the workbook somewhere</summary>
    ''' <param name="Wb">workbook to refresh DB Functions in</param>
    ''' <param name="ignoreCalcMode">when calling refreshDBFunctions time delayed (when saving a workbook and DBFC* is set), need to trigger calculation regardless of calculation mode being manual, otherwise data is not refreshed</param>
    Public Sub refreshDBFunctions(Wb As Excel.Workbook, Optional ignoreCalcMode As Boolean = False)
        Dim WbNames As Excel.Names
        Try : WbNames = Wb.Names
        Catch ex As Exception
            LogWarn("Exception when trying to get Workbook names: " + ex.Message + ", this might be due to errors in the VBA Macros (missing references)")
            Exit Sub
        End Try

        ' hidden workbooks produce an error when searching for cells, this is captured by 
        If TypeName(ExcelDnaUtil.Application.Calculation) = "Error" Then
            LogWarn("ExcelDnaUtil.Application.Calculation = Error, " + Wb.Path + "\" + Wb.Name + " (hidden workbooks produce calculation errors...)")
            Exit Sub
        End If
        DBModifs.preventChangeWhileFetching = True
        Try
            ' walk through all DB functions (having hidden names DBFsource*) cells there to find DB Functions and change their formula, adding " " to trigger recalculation
            For Each DBname As Excel.Name In WbNames
                Dim DBFuncCell As Excel.Range = Nothing
                If DBname.Name Like "*DBFsource*" Then
                    ' some names might have lost their reference to the cell, so catch this here...
                    Try : DBFuncCell = DBname.RefersToRange : Catch ex As Exception : End Try
                End If
                If Not IsNothing(DBFuncCell) Then
                    If DBFuncCell.Parent.ProtectContents Then
                        UserMsg("Worksheet " + DBFuncCell.Parent.Name + " is content protected, can't refresh DB Functions !")
                        Continue For
                    End If
                    Dim callID As String = "" : Dim underlyingName As String = ""
                    If Not (DBFuncCell.Formula.ToString().ToUpper.Contains("DBLISTFETCH") Or DBFuncCell.Formula.ToString().ToUpper.Contains("DBROWFETCH") Or DBFuncCell.Formula.ToString().ToUpper.Contains("DBSETQUERY")) Then
                        LogWarn("Found former DB Function in Cell " + DBFuncCell.Parent.Name + "!" + DBFuncCell.Address + " that doesn't contain a DB Function anymore.")
                    End If
                    Try
                        ' repair DBSheet auto-filling lookup functionality, in case it was lost due to accidental editing of these cells.
                        underlyingName = Replace(DBname.Name, "DBFsource", "DBFtarget", 1, , vbTextCompare)
                        Dim DBTargetListObject As Excel.ListObject = Nothing
                        Try : DBTargetListObject = ExcelDnaUtil.Application.Range(underlyingName).ListObject : Catch ex As Exception : End Try
                        If Not IsNothing(DBTargetListObject) Then
                            ' walk through all columns
                            For Each listcol As Excel.ListColumn In DBTargetListObject.ListColumns
                                Dim colFormula As String = ""
                                ' check for formula and store it
                                colFormula = listcol.DataBodyRange.Cells(1, 1).Formula
                                If Left(colFormula, 1) = "=" Then
                                    DBModifs.preventChangeWhileFetching = True
                                    ' delete whole column
                                    listcol.DataBodyRange.Clear()
                                    ' re-insert the formula, this repairs the auto-filling functionality
                                    listcol.DataBodyRange.Cells(1, 1).Formula = colFormula
                                    DBModifs.preventChangeWhileFetching = False
                                End If
                            Next
                        End If
                        ' get the callID of the underlying name of the target (key of the queryCache and StatusCollection)
                        callID = "[" + DBFuncCell.Parent.Parent.Name + "]" + DBFuncCell.Parent.Name + "!" + DBFuncCell.Address
                        ' remove query cache to force re-fetching
                        queryCache.Remove(callID)
                        ' trigger recalculation by changing formula of DB Function
                        DBFuncCell.Formula += " "
                    Catch ex As Exception
                        LogWarn("Exception when setting Formula or getting callID (" + callID + ") of DB Function in Cell (" + DBFuncCell.Parent.Name + "!" + DBFuncCell.Address + "): " + ex.Message + ", this might be due to errors in the VBA Macros (missing references)")
                    End Try
                End If
            Next
            If ignoreCalcMode And ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual Then
                LogInfo("ignoreCalcMode = True and Application.Calculation = xlCalculationManual, Application.CalculateFull called " + Wb.Path + "\" + Wb.Name)
                ExcelDnaUtil.Application.CalculateFull()
            End If
        Catch ex As Exception
            UserMsg("Exception: " + ex.Message + ", " + Wb.Path + "\" + Wb.Name, "refresh DBFunctions")
        End Try
        DBModifs.preventChangeWhileFetching = False
    End Sub

    ''' <summary>"OnTime" event function to "escape" current (main) thread: event procedure to re-fetch DB functions results after triggering a recalculation inside Application.WorkbookBeforeSave</summary>
    Public Sub refreshDBFuncLater()
        Dim previouslySaved As Boolean
        Dim actWb As Excel.Workbook = Nothing
        Try : actWb = ExcelDnaUtil.Application.ActiveWorkbook : Catch ex As Exception
            LogWarn("Exception when trying to get the active workbook for refreshing DBfunc later: " + ex.Message + ", this might be due to errors in the VBA Macros (missing references)")
        End Try
        Try
            If actWb IsNot Nothing Then
                previouslySaved = actWb.Saved
                LogInfo("clearing DBfunction targets: refreshDBFunctions after clearing")
                refreshDBFunctions(actWb, True)
                actWb.Saved = previouslySaved
            End If
        Catch ex As Exception
            UserMsg("Exception: " + ex.Message, "refresh DBFunc later")
        End Try
    End Sub

    ''' <summary>resets the caches for given workbook</summary>
    ''' <param name="WBname"></param>
    Public Sub resetCachesForWorkbook(WBname As String)
        ' reset query cache for current workbook, so we really get new data !
        Dim tempColl1 As Dictionary(Of String, String) = New Dictionary(Of String, String)(queryCache) ' clone dictionary to be able to remove items...
        For Each resetkey As String In tempColl1.Keys
            If InStr(resetkey, "[" + WBname + "]") > 0 Then queryCache.Remove(resetkey)
        Next
        Dim tempColl2 As Dictionary(Of String, ContainedStatusMsg) = New Dictionary(Of String, ContainedStatusMsg)(StatusCollection)
        For Each resetkey As String In tempColl2.Keys
            If InStr(resetkey, "[" + WBname + "]") > 0 Then StatusCollection.Remove(resetkey)
        Next
    End Sub

    ''' <summary>get a boolean type custom property</summary>
    ''' <param name="name">name of the property</param>
    ''' <param name="Wb">workbook of the property</param>
    ''' <returns>the value of the custom property</returns>
    Public Function getCustPropertyBool(name As String, Wb As Excel.Workbook) As Boolean
        Try
            getCustPropertyBool = Wb.CustomDocumentProperties(name).Value
        Catch ex As Exception
            getCustPropertyBool = False
        End Try
    End Function

    ''' <summary>"repairs" legacy functions from old VB6-COM Addin by removing "DBAddin.Functions." before function name</summary>
    ''' <param name="showResponse">in case this is called interactively, provide a response in case of no legacy functions there</param>
    Public Sub repairLegacyFunctions(actWB As Excel.Workbook, Optional showResponse As Boolean = False)
        Dim foundLegacyFunc As Boolean = False
        Dim xlcalcmode As Long = ExcelDnaUtil.Application.Calculation
        Dim WbNames As Excel.Names
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
                    UserMsg("Replaced legacy functions in active workbook from sheets: " + Left(replaceSheets, replaceSheets.Length - 1), "Legacy DBAddin functions")
                End If
            End If
        Catch ex As Exception
            UserMsg("Exception occurred: " + ex.Message, "Legacy DBAddin functions")
        End Try
        DBModifs.preventChangeWhileFetching = False
    End Sub

    ''' <summary>maintenance procedure to purge names used for dbfunctions from workbook, or unhide DB names</summary>
    Public Sub purgeNames()
        Dim actWbNames As Excel.Names = Nothing
        Try : actWbNames = ExcelDnaUtil.Application.ActiveWorkbook.Names : Catch ex As Exception
            Globals.UserMsg("Exception when trying to get the active workbook's names for purging names: " + ex.Message + ", this might be due to errors in the VBA Macros (missing references)")
        End Try
        If IsNothing(actWbNames) Then Exit Sub
        ' with Ctrl unhide all DB names and show Name Manager...
        If My.Computer.Keyboard.CtrlKeyDown And Not My.Computer.Keyboard.ShiftKeyDown Then
            Dim retval As MsgBoxResult = QuestionMsg("Unhiding all hidden DB function names, continue (refreshing will hide them again)?", MsgBoxStyle.OkCancel, "Unhide names")
            If retval = vbCancel Then Exit Sub
            For Each DBname As Excel.Name In actWbNames
                If DBname.Name Like "*DBFtarget*" Or DBname.Name Like "*DBFsource*" Then DBname.Visible = True
            Next
            Try
                ExcelDnaUtil.Application.Dialogs(Excel.XlBuiltInDialog.xlDialogNameManager).Show()
            Catch ex As Exception
                Globals.UserMsg("The name manager dialog can't be displayed, maybe you are in the formula/cell editor?", "Name manager dialog display")
            End Try
            ' with Shift remove hidden names
        ElseIf My.Computer.Keyboard.ShiftKeyDown And Not My.Computer.Keyboard.CtrlKeyDown Then
            Dim resultingPurges As String = ""
            Dim retval As MsgBoxResult = QuestionMsg("Purging hidden names, should ExternalData names (from Queries) also be purged?", MsgBoxStyle.YesNoCancel, "Purge names")
            If retval = vbCancel Then Exit Sub
            Dim calcMode = ExcelDnaUtil.Application.Calculation
            ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual
            Try
                For Each DBname As Excel.Name In actWbNames
                    If Not DBname.Visible Then ' only hidden names...
                        If (DBname.Name Like "*ExterneDaten*" Or DBname.Name Like "*ExternalData*") And retval = vbYes Then
                            resultingPurges += DBname.Name + ", "
                            DBname.Delete()
                        ElseIf DBname.Name Like "*DBFtarget*" Then
                            resultingPurges += DBname.Name + ", "
                            DBname.Delete()
                        ElseIf DBname.Name Like "*DBFsource*" Then
                            resultingPurges += DBname.Name + ", "
                            DBname.Delete()
                        End If
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
            ExcelDnaUtil.Application.Dialogs(Excel.XlBuiltInDialog.xlDialogNameManager).Show()
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
                Globals.UserMsg("No Problems detected.", "DBfunction check Error", MsgBoxStyle.Information)
            Else
                Globals.UserMsg(collectedErrors, "DBfunction check Error")
            End If
            ' last check any possible DB Modifier Definitions for validity
            DBModifs.getDBModifDefinitions(True)
        End If
    End Sub

End Module