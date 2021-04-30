Imports ExcelDna.Integration
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports System.Configuration

''' <summary>handles all Menu related aspects (context menu for building/refreshing, "DBAddin"/"Load Config" tree menu for retrieving stored configuration files, etc.)</summary>
<ComVisible(True)>
Public Class MenuHandler
    Inherits CustomUI.ExcelRibbon

    ''' <summary>callback after Excel loaded the Ribbon, used to initialize data for the Ribbon</summary>
    Public Sub ribbonLoaded(theRibbon As CustomUI.IRibbonUI)
        Globals.theRibbon = theRibbon
    End Sub

    ''' <summary>creates the Ribbon (only at startup). any changes to the ribbon can only be done via dynamic menus</summary>
    ''' <returns></returns>
    Public Overrides Function GetCustomUI(RibbonID As String) As String
        ' Ribbon definition XML
        Dim customUIXml As String = "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='ribbonLoaded'><ribbon><tabs><tab id='DBaddinTab' label='DB Addin'>"
        ' DBAddin Group: environment choice, DBConfics selection tree, purge names tool button and dialogBoxLauncher for AboutBox
        customUIXml +=
        "<group id='DBAddinGroup' label='DBAddin settings'>" +
            "<dropDown id='envDropDown' label='Environment:' sizeString='1234567890123456' getEnabled='GetEnabledSelect' getSelectedItemIndex='GetSelectedEnvironment' getItemCount='GetItemCount' getItemID='GetItemID' getItemLabel='GetItemLabel' getSupertip='getSelectedTooltip' onAction='selectEnvironment'/>" +
            "<buttonGroup id='buttonGroup1'>" +
                "<menu id='configMenu' label='Settings'>" +
                    "<button id='user' label='User settings' onAction='showAddinConfig' imageMso='ControlProperties' screentip='Show/edit user settings for DB Addin' />" +
                    "<button id='central' label='Central settings' onAction='showAddinConfig' imageMso='TablePropertiesDialog' screentip='Show/edit central settings for DB Addin' />" +
                    "<button id='addin' label='DBAddin settings' onAction='showAddinConfig' imageMso='ServerProperties' screentip='Show/edit standard Addin settings for DB Addin' />" +
                "</menu>" +
                "<button id='props' label='Workbook Properties' onAction='showCProps' getImage='getCPropsImage' screentip='Change custom properties relevant for DB Addin:' getSupertip='getToggleCPropsScreentip' />" +
            "</buttonGroup>" +
            "<labelControl id='InfoLabel' getLabel='getInfoLabel'/>" +
            "<dialogBoxLauncher><button id='dialog' label='About DBAddin' onAction='showAbout' screentip='Show Aboutbox with help, version information and project homepage'/></dialogBoxLauncher>" +
        "</group>"
        ' DBAddin Tools Group:
        customUIXml +=
        "<group id='DBAddinToolsGroup' label='DB Addin Tools'>" +
            "<buttonGroup id='buttonGroup2'>" +
                "<dynamicMenu id='DBConfigs' label='DB Configs' imageMso='QueryShowTable' screentip='DB Function Configuration Files quick access' getContent='getDBConfigMenu'/>" +
                "<menu id='DBSheetMenu' label='DBSheet Def'>" +
                    "<button id='DBSheetCreate' label='Create DBsheet definition' screentip='click to create a new or edit an existing DBSheet definition' imageMso='TableDesign' onAction='clickCreateDBSheet'/>" +
                    "<button id='DBSheetAssign' tag='DBSheet' label='Assign DBsheet definition' screentip='click to assign a DBSheet definition to the current cell' imageMso='ChartResetToMatchStyle' onAction='clickAssignDBSheet'/>" +
                "</menu>" +
            "</buttonGroup>" +
            "<buttonGroup id='buttonGroup3'>" +
                "<button id='purgetool' label='Purge' screentip='purges underlying DBtarget/DBsource Names or unhides hidden names' imageMso='BorderErase' onAction='clickpurgetoolbutton' supertip='hold: Ctrl to unhide all DB names and show Name Manager, Shift to purge hidden names'/>" +
                "<button id='showLog' label='Log' screentip='shows Database Addins Diagnostic Display' getImage='getLogsImage' onAction='clickShowLog'/>" +
                "<button id='designmode' label='Buttons' onAction='showToggleDesignMode' getImage='getToggleDesignImage' getScreentip='getToggleDesignScreentip'/>" +
            "</buttonGroup>" +
            "<button id='repairLegacy' label='repair legacy functions' imageMso='ControlWizards' onAction='clickRepairLegacyFunctions' screentip='click to fix legacy functions from old VB6 DBAddin'/>" +
        "</group>"
        ' DBModif Group: maximum three DBModif types possible (depending on existence in current workbook): 
        customUIXml +=
        "<group id='DBModifGroup' label='Execute DBModifier'>"
        For Each DBModifType As String In {"DBSeqnce", "DBMapper", "DBAction"}
            customUIXml += "<dynamicMenu id='" + DBModifType + "' " +
                                                "size='large' getLabel='getDBModifTypeLabel' imageMso='ApplicationOptionsDialog' " +
                                                "getScreentip='getDBModifScreentip' getContent='getDBModifMenuContent' getVisible='getDBModifMenuVisible'/>"
        Next
        customUIXml += "<dialogBoxLauncher><button id='DBModifEdit' label='DBModif design' onAction='showDBModifEdit' screentip='Show/edit DBModif Definitions of current workbook'/></dialogBoxLauncher>"
        customUIXml += "</group>"
        customUIXml += "</tab></tabs></ribbon>"
        ' Context menus for refresh, jump and creation: in cell, row, column and ListRange (area of ListObjects)
        customUIXml += "<contextMenus>" +
        "<contextMenu idMso ='ContextMenuCell'>" +
            "<button id='refreshDataC' label='refresh data (Ctl-Sh-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
            "<button id='gotoDBFuncC' label='jump to DBFunc/target (Ctl-Sh-J)' imageMso='ConvertTextToTable' onAction='clickjumpButton' insertBeforeMso='Cut'/>" +
             "<menu id='createMenu' label='Insert/Edit DBFunc/DBModif' insertBeforeMso='Cut'>" +
                "<button id='DBMapperC' tag='DBMapper' label='DBMapper' imageMso='TableSave' onAction='clickCreateButton'/>" +
                "<button id='DBActionC' tag='DBAction' label='DBAction' imageMso='TableIndexes' onAction='clickCreateButton'/>" +
                "<button id='DBSequenceC' tag='DBSeqnce' label='DBSequence' imageMso='ShowOnNewButton' onAction='clickCreateButton'/>" +
                "<menuSeparator id='separator' />" +
                "<button id='DBListFetchC' tag='DBListFetch' label='DBListFetch' imageMso='GroupLists' onAction='clickCreateButton'/>" +
                "<button id='DBRowFetchC' tag='DBRowFetch' label='DBRowFetch' imageMso='GroupRecords' onAction='clickCreateButton'/>" +
                "<button id='DBSetQueryPivotC' tag='DBSetQueryPivot' label='DBSetQueryPivot' imageMso='AddContentType' onAction='clickCreateButton'/>" +
                "<button id='DBSetQueryListObjectC' tag='DBSetQueryListObject' label='DBSetQueryListObject' imageMso='AddContentType' onAction='clickCreateButton'/>" +
            "</menu>" +
            "<menuSeparator id='MySeparatorC' insertBeforeMso='Cut'/>" +
        "</contextMenu>" +
        "<contextMenu idMso ='ContextMenuPivotTable'>" +
            "<button id='refreshDataPT' label='refresh data (Ctl-Sh-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Copy'/>" +
            "<button id='gotoDBFuncPT' label='jump to DBFunc/target (Ctl-Sh-J)' imageMso='ConvertTextToTable' onAction='clickjumpButton' insertBeforeMso='Copy'/>" +
            "<menuSeparator id='MySeparatorPT' insertBeforeMso='Copy'/>" +
        "</contextMenu>" +
        "<contextMenu idMso ='ContextMenuCellLayout'>" +
            "<button id='refreshDataCL' label='refresh data (Ctl-Sh-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
            "<button id='gotoDBFuncCL' label='jump to DBFunc/target (Ctl-Sh-J)' imageMso='ConvertTextToTable' onAction='clickjumpButton' insertBeforeMso='Cut'/>" +
            "<menu id='createMenuCL' label='Insert/Edit DBFunc/DBModif' insertBeforeMso='Cut'>" +
                "<button id='DBMapperCL' tag='DBMapper' label='DBMapper' imageMso='TableSave' onAction='clickCreateButton'/>" +
                "<button id='DBActionCL' tag='DBAction' label='DBAction' imageMso='TableIndexes' onAction='clickCreateButton'/>" +
                "<button id='DBSequenceCL' tag='DBSeqnce' label='DBSequence' imageMso='ShowOnNewButton' onAction='clickCreateButton'/>" +
                "<menuSeparator id='separatorCL' />" +
                "<button id='DBListFetchCL' tag='DBListFetch' label='DBListFetch' imageMso='GroupLists' onAction='clickCreateButton'/>" +
                "<button id='DBRowFetchCL' tag='DBRowFetch' label='DBRowFetch' imageMso='GroupRecords' onAction='clickCreateButton'/>" +
                "<button id='DBSetQueryPivotCL' tag='DBSetQueryPivot' label='DBSetQueryPivot' imageMso='AddContentType' onAction='clickCreateButton'/>" +
                "<button id='DBSetQueryListObjectCL' tag='DBSetQueryListObject' label='DBSetQueryListObject' imageMso='AddContentType' onAction='clickCreateButton'/>" +
            "</menu>" +
            "<menuSeparator id='MySeparatorCL' insertBeforeMso='Cut'/>" +
        "</contextMenu>" +
        "<contextMenu idMso='ContextMenuRow'>" +
            "<button id='refreshDataR' label='refresh data (Ctl-Sh-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
            "<menuSeparator id='MySeparatorR' insertBeforeMso='Cut'/>" +
        "</contextMenu>" +
        "<contextMenu idMso='ContextMenuColumn'>" +
            "<button id='refreshDataZ' label='refresh data (Ctl-Sh-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
            "<menuSeparator id='MySeparatorZ' insertBeforeMso='Cut'/>" +
        "</contextMenu>" +
        "<contextMenu idMso='ContextMenuListRange'>" +
            "<button id='refreshDataL' label='refresh data (Ctl-Sh-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
            "<button id='gotoDBFuncL' label='jump to DBFunc/target (Ctl-Sh-J)' imageMso='ConvertTextToTable' onAction='clickjumpButton' insertBeforeMso='Cut'/>" +
            "<menu id='createMenuL' label='Insert/Edit DBFunc/DBModif' insertBeforeMso='Cut'>" +
                "<button id='DBMapperL' tag='DBMapper' label='DBMapper' imageMso='TableSave' onAction='clickCreateButton'/>" +
                "<button id='DBSequenceL' tag='DBSeqnce' label='DBSequence' imageMso='ShowOnNewButton' onAction='clickCreateButton'/>" +
            "</menu>" +
            "<menuSeparator id='MySeparatorL' insertBeforeMso='Cut'/>" +
        "</contextMenu>" +
        "</contextMenus></customUI>"
        Return customUIXml
    End Function

#Disable Warning IDE0060 ' Hide not used Parameter warning as this is very often the case with the below callbacks from the ribbon

    Public Sub clickRepairLegacyFunctions(control As CustomUI.IRibbonControl)
        Globals.repairLegacyFunctions(True)
    End Sub



    Public Function getInfoLabel(ByRef control As CustomUI.IRibbonControl)
        getInfoLabel = "DBModif Impl: " + IIf(DBModifs.altDBImpl, "ADO.NET", "ADODB")
    End Function

    ''' <summary>display warning button icon on Cprops change if DBFskip is set...</summary>
    ''' <param name="control"></param>
    ''' <returns></returns>
    Public Function getCPropsImage(control As CustomUI.IRibbonControl) As String
        If Globals.getCustPropertyBool("DBFskip", ExcelDnaUtil.Application.ActiveWorkbook) Then
            Return "DeclineTask"
        Else
            Return "AcceptTask"
        End If
    End Function

    ''' <summary>display warning icon on log button if warning has been logged...</summary>
    ''' <param name="control"></param>
    ''' <returns></returns>
    Public Function getLogsImage(control As CustomUI.IRibbonControl) As String
        If Globals.WarningIssued Then
            Return "IndexUpdate"
        Else
            Return "MailMergeStartLetters"
        End If
    End Function

    ''' <summary>display state of designmode in screentip of dialogBox launcher</summary>
    ''' <param name="control"></param>
    ''' <returns>screentip and the state of designmode</returns>
    Public Function getToggleCPropsScreentip(control As CustomUI.IRibbonControl) As String
        getToggleCPropsScreentip = ""
        If Not ExcelDnaUtil.Application.ActiveWorkbook Is Nothing Then
            Try
                Dim docproperty As Microsoft.Office.Core.DocumentProperty
                For Each docproperty In ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties
                    If Left$(docproperty.Name, 5) = "DBFC" Or docproperty.Name = "DBFskip" Or docproperty.Name = "doDBMOnSave" Or docproperty.Name = "DBFNoLegacyCheck" Then
                        getToggleCPropsScreentip += docproperty.Name + ":" + docproperty.Value.ToString() + vbCrLf
                    End If
                Next
            Catch ex As Exception
                getToggleCPropsScreentip += "exception when collecting docproperties: " + ex.Message
            End Try
        End If
    End Function

    ''' <summary>click on change props: show builtin properties dialog</summary>
    ''' <param name="control"></param>
    Public Sub showCProps(control As CustomUI.IRibbonControl)
        If Not ExcelDnaUtil.Application.ActiveWorkbook Is Nothing Then
            ExcelDnaUtil.Application.Dialogs(Excel.XlBuiltInDialog.xlDialogProperties).Show
            ' to check whether DBFskip has changed:
            Globals.theRibbon.InvalidateControl(control.Id)
        End If
    End Sub

    ''' <summary>show DBModif definitions edit box</summary>
    ''' <param name="control"></param>
    Sub showDBModifEdit(control As CustomUI.IRibbonControl)
        ' only show dialog if there is a workbook and it has the relevant custom XML part.
        If Not ExcelDnaUtil.Application.ActiveWorkbook Is Nothing AndAlso ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef").Count > 0 Then
            Dim theEditDBModifDefDlg As EditDBModifDef = New EditDBModifDef()
            If theEditDBModifDefDlg.ShowDialog() = System.Windows.Forms.DialogResult.OK Then DBModifs.getDBModifDefinitions()
        End If
    End Sub

    ''' <summary>toggle designmode button</summary>
    ''' <param name="control"></param>
    Sub showToggleDesignMode(control As CustomUI.IRibbonControl)
        Dim cbrs As Object = ExcelDnaUtil.Application.CommandBars
        If Not cbrs Is Nothing AndAlso cbrs.GetEnabledMso("DesignMode") Then
            cbrs.ExecuteMso("DesignMode")
        Else
            Globals.UserMsg("Couldn't toggle designmode, because Designmode commandbar button is not available (no button?)", "DBAddin toggle Designmode", MsgBoxStyle.Exclamation)
        End If
        ' update state of designmode in screentip
        Globals.theRibbon.InvalidateControl(control.Id)
    End Sub

    ''' <summary>display state of designmode in screentip of button</summary>
    ''' <param name="control"></param>
    ''' <returns>screentip and the state of designmode</returns>
    Public Function getToggleDesignScreentip(control As CustomUI.IRibbonControl) As String
        Dim cbrs As Object = ExcelDnaUtil.Application.CommandBars
        If Not cbrs Is Nothing AndAlso cbrs.GetEnabledMso("DesignMode") Then
            Return "Designmode is currently " + IIf(cbrs.GetPressedMso("DesignMode"), "on !", "off !")
        Else
            Return "Designmode commandbar button not available (no button on sheet)"
        End If
    End Function

    ''' <summary>display state of designmode in icon of button</summary>
    ''' <param name="control"></param>
    ''' <returns>screentip and the state of designmode</returns>
    Public Function getToggleDesignImage(control As CustomUI.IRibbonControl) As String
        Dim cbrs As Object = ExcelDnaUtil.Application.CommandBars
        If Not cbrs Is Nothing AndAlso cbrs.GetEnabledMso("DesignMode") Then
            If cbrs.GetPressedMso("DesignMode") Then
                Return "ObjectsGroupMenuOutlook"
            Else
                Return "SelectMenuAccess"
            End If
        Else
            Return "SelectMenuAccess"
        End If
    End Function

    ''' <summary>for environment dropdown to get the total number of the entries</summary>
    ''' <returns></returns>
    Public Function GetItemCount(control As CustomUI.IRibbonControl) As Integer
        Return Globals.environdefs.Length
    End Function

    ''' <summary>for environment dropdown to get the label of the entries</summary>
    ''' <returns></returns>
    Public Function GetItemLabel(control As CustomUI.IRibbonControl, index As Integer) As String
        Return Globals.environdefs(index)
    End Function

    ''' <summary>for environment dropdown to get the ID of the entries</summary>
    ''' <returns></returns>
    Public Function GetItemID(control As CustomUI.IRibbonControl, index As Integer) As String
        Return Globals.environdefs(index)
    End Function

    ''' <summary>after selection of environment (using selectEnvironment) used to return the selected environment</summary>
    ''' <returns></returns>
    Public Function GetSelectedEnvironment(control As CustomUI.IRibbonControl) As Integer
        Return Globals.selectedEnvironment
    End Function

    ''' <summary>tooltip for the environment select drop down</summary>
    ''' <param name="control"></param>
    ''' <returns>the tooltip</returns>
    Public Function getSelectedTooltip(control As CustomUI.IRibbonControl) As String
        If CBool(fetchSetting("DontChangeEnvironment", "False")) Then
            Return "DontChangeEnvironment is set, therefore changing the Environment is prevented !"
        Else
            Return "configured for Database Access in Addin config %appdata%\Microsoft\Addins\DBaddin.xll.config (or referenced central/user setting)"
        End If
    End Function

    ''' <summary>whether to enable environment select drop down</summary>
    ''' <param name="control"></param>
    ''' <returns>true if enabled</returns>
    Public Function GetEnabledSelect(control As CustomUI.IRibbonControl) As Integer
        Return Not CBool(fetchSetting("DontChangeEnvironment", "False"))
    End Function

    ''' <summary>Choose environment (configured in registry with ConstConnString(N), ConfigStoreFolder(N))</summary>
    ''' <param name="control"></param>
    Public Sub selectEnvironment(control As CustomUI.IRibbonControl, id As String, index As Integer)
        Globals.selectedEnvironment = index
        Globals.initSettings()
        ' provide a chance to reconnect when switching environment...
        conn = Nothing
        If Not ExcelDnaUtil.Application.ActiveWorkbook Is Nothing Then
            Dim retval As MsgBoxResult = QuestionMsg("ConstConnString:" + Globals.ConstConnString + vbCrLf + "ConfigStoreFolder:" + ConfigFiles.ConfigStoreFolder + vbCrLf + vbCrLf + "Refresh DBFunctions in active workbook to see effects?", MsgBoxStyle.YesNo, "Changed environment to: " + fetchSetting("ConfigName" + Globals.env(), ""))
            If retval = vbYes Then Globals.refreshDBFunctions(ExcelDnaUtil.Application.ActiveWorkbook)
        Else
            Globals.UserMsg("ConstConnString:" + Globals.ConstConnString + vbCrLf + "ConfigStoreFolder:" + ConfigFiles.ConfigStoreFolder, "Changed environment to: " + fetchSetting("ConfigName" + Globals.env(), ""), MsgBoxStyle.Information)
        End If
    End Sub

    ''' <summary>show xll standard config (AppSetting), central config (referenced by App Settings file attr) or user config (referenced by CustomSettings configSource attr)</summary>
    ''' <param name="control"></param>
    Public Sub showAddinConfig(control As CustomUI.IRibbonControl)
        ' if settings (addin, user, central) should not be displayed according to setting then exit...
        If InStr(fetchSetting("disableSettingsDisplay", ""), control.Id) > 0 Then
            Globals.UserMsg("Display of " + control.Id + " settings disabled !", "DBAddin Settings disabled", MsgBoxStyle.Information)
            Exit Sub
        End If
        Dim theEditDBModifDefDlg As EditDBModifDef = New EditDBModifDef()
        theEditDBModifDefDlg.DBFskip.Hide()
        theEditDBModifDefDlg.doDBMOnSave.Hide()
        theEditDBModifDefDlg.Tag = control.Id
        theEditDBModifDefDlg.ShowDialog()
        If control.Id = "addin" Or control.Id = "central" Then
            ConfigurationManager.RefreshSection("appSettings")
        Else
            ConfigurationManager.RefreshSection("UserSettings")
        End If
        ' reflect changes in settings
        Globals.initSettings()
        ' also display in ribbon
        Globals.theRibbon.Invalidate()
    End Sub

    ''' <summary>dialogBoxLauncher of DBAddin settings group: activate about box</summary>
    ''' <param name="control"></param>
    Public Sub showAbout(control As CustomUI.IRibbonControl)
        Dim myAbout As AboutBox = New AboutBox
        myAbout.ShowDialog()
        ' if disabling the addin was chosen, then suicide here..
        If myAbout.disableAddinAfterwards Then
            Try : ExcelDnaUtil.Application.AddIns("DBaddin").Installed = False : Catch ex As Exception : End Try
        End If
    End Sub

    ''' <summary>on demand, refresh the DB Config tree</summary>
    ''' <param name="control"></param>
    Public Sub refreshDBConfigTree(control As CustomUI.IRibbonControl)
        Globals.initSettings()
        ConfigFiles.createConfigTreeMenu()
        Globals.UserMsg("refreshed DB Config Tree Menu", "DBAddin: refresh Config tree...", MsgBoxStyle.Information)
        Globals.theRibbon.Invalidate()
    End Sub

    ''' <summary>get DB Config Menu from File</summary>
    ''' <param name="control"></param>
    ''' <returns></returns>
    Public Function getDBConfigMenu(control As CustomUI.IRibbonControl) As String
        If ConfigFiles.ConfigMenuXML = vbNullString Then ConfigFiles.createConfigTreeMenu()
        Return ConfigFiles.ConfigMenuXML
    End Function

    ''' <summary>load config if config tree menu end-button has been activated (path to config xcl file is in control.Tag)</summary>
    ''' <param name="control"></param>
    Public Sub getConfig(control As CustomUI.IRibbonControl)
        ConfigFiles.loadConfig(control.Tag)
    End Sub

    ''' <summary>set the name of the DBModifType dropdown to the sheet name (for the WB dropdown this is the WB name)</summary>
    ''' <param name="control"></param>
    ''' <returns></returns>
    Public Function getDBModifTypeLabel(control As CustomUI.IRibbonControl) As String
        getDBModifTypeLabel = If(control.Id = "DBSeqnce", "DBSequence", control.Id)
    End Function

    ''' <summary>create the buttons in the DBModif dropdown menu</summary>
    ''' <param name="control"></param>
    ''' <returns>the menu content xml</returns>
    Public Function getDBModifMenuContent(control As CustomUI.IRibbonControl) As String
        Dim xmlString As String = "<menu xmlns='http://schemas.microsoft.com/office/2009/07/customui'>"
        Try
            If Not Globals.DBModifDefColl.ContainsKey(control.Id) Then Return ""
            Dim DBModifTypeName As String = IIf(control.Id = "DBSeqnce", "DBSequence", IIf(control.Id = "DBMapper", "DB Mapper", IIf(control.Id = "DBAction", "DB Action", "undefined DBModifTypeName")))
            For Each nodeName As String In Globals.DBModifDefColl(control.Id).Keys
                Dim descName As String = IIf(nodeName = control.Id, "Unnamed " + DBModifTypeName, Replace(nodeName, DBModifTypeName, ""))
                Dim imageMsoStr As String = IIf(control.Id = "DBSeqnce", "ShowOnNewButton", IIf(control.Id = "DBMapper", "TableSave", IIf(control.Id = "DBAction", "TableIndexes", "undefined imageMso")))
                Dim superTipStr As String = IIf(control.Id = "DBSeqnce", "executes " + DBModifTypeName + " defined in " + nodeName, IIf(control.Id = "DBMapper", "stores data defined in DBMapper (named " + nodeName + ") range on " + Globals.DBModifDefColl(control.Id).Item(nodeName).getTargetRangeAddress(), IIf(control.Id = "DBAction", "executes Action defined in DBAction (named " + nodeName + ") range on " + Globals.DBModifDefColl(control.Id).Item(nodeName).getTargetRangeAddress(), "undefined superTip")))
                xmlString = xmlString + "<button id='_" + nodeName + "' label='do " + descName + "' imageMso='" + imageMsoStr + "' onAction='DBModifClick' tag='" + control.Id + "' screentip='do " + DBModifTypeName + ": " + descName + "' supertip='" + superTipStr + "' />"
            Next
            xmlString += "</menu>"
            Return xmlString
        Catch ex As Exception
            Globals.UserMsg("Exception caught while building xml: " + ex.Message)
            Return ""
        End Try
    End Function

    ''' <summary>show a screentip for the dynamic DBMapper/DBAction/DBSequence Menus (also showing the ID behind)</summary>
    ''' <param name="control"></param>
    ''' <returns>the screentip</returns>
    Public Function getDBModifScreentip(control As CustomUI.IRibbonControl) As String
        Return "Select DBModifier to store/do action/do sequence (" + control.Id + ")"
    End Function

    ''' <summary>to show the DBModif sheet button only if it was collected...</summary>
    ''' <param name="control"></param>
    ''' <returns>true if to be displayed</returns>
    Public Function getDBModifMenuVisible(control As CustomUI.IRibbonControl) As Boolean
        Try
            Return Globals.DBModifDefColl.ContainsKey(control.Id)
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>DBModif button activated, do DB Mapper/DB Action/DB Sequence or define existing (CtrlKey pressed)...</summary>
    ''' <param name="control"></param>
    Public Sub DBModifClick(control As CustomUI.IRibbonControl)
        ' reset noninteractive messages (used for VBA invocations) and hadError for interactive invocations
        nonInteractiveErrMsgs = "" : hadError = False
        Dim nodeName As String = Right(control.Id, Len(control.Id) - 1)
        If Not ExcelDnaUtil.Application.CommandBars.GetEnabledMso("FileNewDefault") Then
            Globals.UserMsg("Cannot execute DB Modifier while cell editing active !", "DB Modifier execution", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Try
            If My.Computer.Keyboard.CtrlKeyDown And My.Computer.Keyboard.ShiftKeyDown Then
                createDBModif(control.Tag, targetDefName:=nodeName)
            Else
                ' DB sequence actions (the sequence to be done) are stored directly in DBMapperDefColl, so different invocation here
                If Not (ExcelDnaUtil.Application.ActiveWorkbook.ReadOnlyRecommended And ExcelDnaUtil.Application.ActiveWorkbook.ReadOnly) Then
                    Globals.DBModifDefColl(control.Tag).Item(nodeName).doDBModif()
                Else
                    Globals.UserMsg("ReadOnlyRecommended is set on active workbook (being readonly), therefore all DB Modifiers are disabled !", "DB Modifier execution", MsgBoxStyle.Exclamation)
                End If
            End If
        Catch ex As Exception
            Globals.UserMsg("Exception: " + ex.Message + ",control.Tag:" + control.Tag + ",nodeName:" + nodeName, "DBModif Click")
        End Try
    End Sub

    ''' <summary>context menu entry refreshData: refresh Data in db function (if area or cell selected) or all db functions</summary>
    ''' <param name="control"></param>
    Public Sub clickrefreshData(control As CustomUI.IRibbonControl)
        Globals.refreshData()
    End Sub

    ''' <summary>context menu entry gotoDBFunc: jumps from DB function to data area and back</summary>
    ''' <param name="control"></param>
    Public Sub clickjumpButton(control As CustomUI.IRibbonControl)
        Globals.jumpButton()
    End Sub

    ''' <summary>purge name tool button, purge names used for dbfunctions from workbook</summary>
    ''' <param name="control"></param>
    Public Sub clickpurgetoolbutton(control As CustomUI.IRibbonControl)
        If Not ExcelDnaUtil.Application.ActiveWorkbook Is Nothing Then
            Globals.purgeNames()
        End If
    End Sub

    ''' <summary>show the trace log</summary>
    ''' <param name="control"></param>
    Public Sub clickShowLog(control As CustomUI.IRibbonControl)
        ExcelDna.Logging.LogDisplay.Show()
        ' reset warning flag
        WarningIssued = False
        theRibbon.InvalidateControl("showLog")
    End Sub

    ''' <summary>ribbon menu button for DBSheet creation start</summary>
    ''' <param name="control"></param>
    Public Sub clickCreateDBSheet(control As CustomUI.IRibbonControl)
        Dim theDBSheetCreateForm As DBSheetCreateForm = New DBSheetCreateForm
        theDBSheetCreateForm.Show()
    End Sub

    ''' <summary>context menu entries in Insert/Edit DBFunc/DBModif and Assign DBSheet: create DB function or DB Modification definition</summary>
    ''' <param name="control"></param>
    Public Sub clickCreateButton(control As CustomUI.IRibbonControl)
        ' check for existing DBMapper or DBAction definition and allow exit
        Dim activeCellDBModifName As String = DBModifs.getDBModifNameFromRange(ExcelDnaUtil.Application.ActiveCell)
        Dim activeCellDBModifType As String = Left(activeCellDBModifName, 8)
        If (activeCellDBModifType = "DBMapper" Or activeCellDBModifType = "DBAction") And activeCellDBModifType <> control.Tag And control.Tag <> "DBSeqnce" Then
            Globals.UserMsg("Active Cell already contains definition for a " + activeCellDBModifType + ", inserting " + IIf(control.Tag = "DBSetQueryPivot" Or control.Tag = "DBSetQueryListObject", "DBSetQuery", control.Tag) + " here will cause trouble !", "Inserting not allowed")
            Exit Sub
        End If
        If control.Tag = "DBListFetch" Then
            ConfigFiles.createFunctionsInCells(ExcelDnaUtil.Application.ActiveCell, {"RC", "=DBListFetch("""","""",R[1]C,,,TRUE,TRUE,TRUE)"})
        ElseIf control.Tag = "DBRowFetch" Then
            ConfigFiles.createFunctionsInCells(ExcelDnaUtil.Application.ActiveCell, {"RC", "=DBRowFetch("""","""",TRUE,R[1]C:R[1]C[10])"})
        ElseIf control.Tag = "DBSetQueryPivot" Then
            ' first create a dummy pivot table
            ConfigFiles.createPivotTable(ExcelDnaUtil.Application.ActiveCell)
            ' then create the DBSetQuery assigning the (yet to be filled) query to the above listobject
            ConfigFiles.createFunctionsInCells(ExcelDnaUtil.Application.ActiveCell, {"RC", "=DBSetQuery("""","""",R[1]C)"})
        ElseIf control.Tag = "DBSetQueryListObject" Then
            ' first create a dummy ListObject
            ConfigFiles.createListObject(ExcelDnaUtil.Application.ActiveCell)
            ' then create the DBSetQuery assigning the (yet to be filled) query to the above listobject
            ConfigFiles.createFunctionsInCells(ExcelDnaUtil.Application.ActiveCell, {"RC", "=DBSetQuery("""","""",RC[1])"})
        ElseIf control.Tag = "DBMapper" Or control.Tag = "DBAction" Or control.Tag = "DBSeqnce" Then
            If activeCellDBModifType = control.Tag Then  ' edit existing definition
                DBModifs.createDBModif(control.Tag, targetDefName:=activeCellDBModifName)
            Else                                         ' create new definition
                DBModifs.createDBModif(control.Tag)
            End If
        End If
    End Sub

    ''' <summary>clicked Assign DBSheet: create DB Mapper with CUD Flags</summary>
    ''' <param name="control"></param>
    Public Sub clickAssignDBSheet(control As CustomUI.IRibbonControl)
        If Not ExcelDnaUtil.Application.ActiveWorkbook Is Nothing Then
            DBSheetConfig.createDBSheet()
        Else
            Globals.UserMsg("Cannot assign DBSheet DB Mapper as there is no Workbook active !", "DB Sheet Assignment", MsgBoxStyle.Exclamation)
        End If
    End Sub

#Enable Warning IDE0060
End Class
