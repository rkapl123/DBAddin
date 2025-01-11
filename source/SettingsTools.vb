Imports ExcelDna.Integration
Imports System.Configuration

''' <summary>Global variables/functions for settings access</summary>
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
    ''' <summary>The path where the User specific settings (overrides of standard/global settings) can be found (hard-coded to path of xll)</summary>
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

    ''' <summary>the configurable shortcut for the refresh data key (default ^R)</summary>
    Public refreshDataKey As String
    ''' <summary>the configurable shortcut for the jump button key (default ^J)</summary>
    Public jumpButtonKey As String
    ''' <summary>the configurable shortcut for the delete row key (default ^D)</summary>
    Public deleteRowKey As String
    ''' <summary>the configurable shortcut for the insert row key (default ^I)</summary>
    Public insertRowKey As String
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

End Module
