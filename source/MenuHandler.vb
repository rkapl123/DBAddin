Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports System.Configuration
Imports System.Collections.Specialized
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Windows.Forms


''' <summary>handles all Menu related aspects (context menu for building/refreshing, "DBAddin"/"Load Config" tree menu for retrieving stored configuration files, etc.)</summary>
<ComVisible(True)>
Public Class MenuHandler
    Inherits CustomUI.ExcelRibbon

    ''' <summary>callback after Excel loaded the Ribbon, used to initialize data for the Ribbon</summary>
    Public Sub ribbonLoaded(theRibbon As CustomUI.IRibbonUI)
        MenuHandlerGlobals.theRibbon = theRibbon
        initAdhocSQLconfig()
    End Sub

    ''' <summary>used to avoid crashes when closing excel (especially with multiple users of IntelliSenseServer)</summary>
    ''' <param name="custom"></param>
    Public Overrides Sub OnBeginShutdown(ByRef custom As Array)
        'ExcelDna.IntelliSense.IntelliSenseServer.Uninstall()
    End Sub

    ''' <summary>used to clean up com objects to avoid excel crashes on shutdown</summary>
    ''' <param name="RemoveMode"></param>
    ''' <param name="custom"></param>
    Public Overrides Sub OnDisconnection(RemoveMode As ExcelDna.Integration.Extensibility.ext_DisconnectMode, ByRef custom As Array)
        For Each aKey As String In Functions.StatusCollection.Keys
            Dim clearRange As Excel.Range = Functions.StatusCollection(aKey).formulaRange
            If Not IsNothing(clearRange) Then Marshal.ReleaseComObject(clearRange)
        Next
        For Each DBmodifType As String In DBModifDefColl.Keys
            For Each dbmapdefkey As String In DBModifDefColl(DBmodifType).Keys
                Dim clearRange As Excel.Range = DBModifDefColl(DBmodifType).Item(dbmapdefkey).getTargetRange()
                If Not IsNothing(clearRange) Then Marshal.ReleaseComObject(clearRange)
            Next
        Next
        If Not IsNothing(DBSheetConfig.curCell) Then Marshal.ReleaseComObject(DBSheetConfig.curCell)
        If Not IsNothing(DBSheetConfig.createdListObject) Then Marshal.ReleaseComObject(DBSheetConfig.createdListObject)
        If Not IsNothing(DBSheetConfig.lookupWS) Then Marshal.ReleaseComObject(DBSheetConfig.lookupWS)
        Do
            GC.Collect()
            GC.WaitForPendingFinalizers()
        Loop While (Marshal.AreComObjectsAvailableForCleanup())
    End Sub

    ''' <summary>creates the Ribbon (only at startup). any changes to the ribbon can only be done via dynamic menus</summary>
    ''' <returns></returns>
    Public Overrides Function GetCustomUI(RibbonID As String) As String
        ' Ribbon definition XML
        Dim customUIXml As String = "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='ribbonLoaded'><ribbon><tabs><tab id='DBaddinTab' label='DB Addin'>"
        ' DBAddin Group: environment choice, DBConfics selection tree, purge names tool button and dialogBoxLauncher for AboutBox
        customUIXml +=
        "<group id='DBAddinGroup' label='DBAddin settings'>" +
            "<dropDown id='envDropDown' label='Environment:' sizeString='1234567890123456' getEnabled='GetEnvEnabled' getSelectedItemIndex='GetSelectedEnvironment' getItemCount='GetEnvItemCount' getItemID='GetEnvItemID' getItemLabel='GetEnvItemLabel' getSupertip='GetEnvSelectedTooltip' onAction='selectEnvironment'/>" +
            "<buttonGroup id='buttonGroup0'>" +
                "<menu id='configMenu' label='Settings'>" +
                    "<button id='user' label='User settings' onAction='showAddinConfig' imageMso='ControlProperties' screentip='Show/edit user settings for DB Addin' />" +
                    "<button id='central' label='Central settings' onAction='showAddinConfig' imageMso='TablePropertiesDialog' screentip='Show/edit central settings for DB Addin' />" +
                    "<button id='addin' label='DBAddin settings' onAction='showAddinConfig' imageMso='ServerProperties' screentip='Show/edit standard Addin settings for DB Addin' />" +
                "</menu>" +
                "<button id='props' label='Workbook Properties' onAction='showCProps' getImage='getCPropsImage' screentip='Change custom properties relevant for DB Addin:' getSupertip='getToggleCPropsScreentip' />" +
            "</buttonGroup>" +
            "<buttonGroup id='buttonGroup1'>" +
                "<button id='showLog' label='Log' screentip='shows Database Addins Diagnostic Display' getImage='getLogsImage' onAction='clickShowLog'/>" +
                "<button id='preventRefresh' label='DBFunc Refresh' onAction='showTogglePreventRefresh' getImage='getTogglePreventRefreshImage' getScreentip='getTogglePreventRefreshScreentip' />" +
            "</buttonGroup>" +
            "<dialogBoxLauncher><button id='dialog' label='About DBAddin' onAction='showAbout' screentip='Show Aboutbox with help, version information, update check/download and project homepage' getSupertip='getSuperTipInfo'/></dialogBoxLauncher>" +
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
                "<button id='checkpurgetool' label='Check/Purge' screentip='checks, unhides or purges DBFunctions underlying names' imageMso='BorderErase' onAction='clickcheckpurgetoolbutton' supertip='while clicking hold: Ctrl to unhide all DB names and show Name Manager, Shift to purge DBFunc names, both Ctrl and Shift to display name manager. Nothing will just check DB Functions'/>" +
                "<button id='designmode' label='Buttons' onAction='showToggleDesignMode' getImage='getToggleDesignImage' getScreentip='getToggleDesignScreentip'/>" +
            "</buttonGroup>" +
            "<comboBox id='DBAdhocSQL' showLabel='false' sizeString='123456789012345678901234567' getText='GetAdhocSQLText' getItemCount='GetAdhocSQLItemCount' getItemLabel='GetAdhocSQLItemLabel' onChange='showDBAdHocSQL' screentip='enter Ad-hoc SQL statements to execute'/>" +
            "<dialogBoxLauncher><button id='AdHocSQL' label='AdHoc SQL Command' onAction='showDBAdHocSQLDBOX' screentip='Open AdHoc SQL Command Tool'/></dialogBoxLauncher>" +
        "</group>"
        ' DBModif Group: maximum three DBModif types possible (depending on existence in current workbook): 
        customUIXml +=
        "<group id='DBModifGroup' label='Execute DBModifier'>"
        For Each DBModifType As String In {"DBSeqnce", "DBMapper", "DBAction"}
            customUIXml += "<dynamicMenu id='" + DBModifType + "' " +
                                                "size='large' getLabel='getDBModifTypeLabel' imageMso='ApplicationOptionsDialog' " +
                                                "getScreentip='getDBModifScreentip' getContent='getDBModifMenuContent' getVisible='getDBModifMenuVisible'/>"
        Next
        customUIXml += "<dialogBoxLauncher><button id='DBModifEdit' label='DBModif design' onAction='showDBModifEdit' screentip='Show/edit DBModif Definitions of current workbook'/></dialogBoxLauncher>" +
        "</group></tab></tabs></ribbon>"
        ' Context menus for refresh, jump and creation: in cell, row, column, pivot table, ListRange (area of ListObjects) and query area
        customUIXml +=
        "<contextMenus>" +
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
                    "<button id='DBSetPowerQueryC' tag='DBSetPowerQuery' label='DBSetPowerQuery' imageMso='AddContentType' onAction='clickCreateButton'/>" +
                "</menu>" +
                "<menuSeparator id='MySeparatorC' insertBeforeMso='Cut'/>" +
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
                    "<button id='DBSetPowerQueryCL' tag='DBSetPowerQuery' label='DBSetPowerQuery' imageMso='AddContentType' onAction='clickCreateButton'/>" +
                "</menu>" +
                "<menuSeparator id='MySeparatorCL' insertBeforeMso='Cut'/>" +
            "</contextMenu>" +
            "<contextMenu idMso ='ContextMenuPivotTable'>" +
                "<button id='refreshDataPT' label='refresh data (Ctl-Sh-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Copy'/>" +
                "<button id='gotoDBFuncPT' label='jump to DBFunc/target (Ctl-Sh-J)' imageMso='ConvertTextToTable' onAction='clickjumpButton' insertBeforeMso='Copy'/>" +
                "<button id='createChart' label='create Pivot Chart' imageMso='ChartInsert' onAction='clickCreateChart' insertBeforeMso='Copy'/>" +
                "<menuSeparator id='MySeparatorPT' insertBeforeMso='Copy'/>" +
            "</contextMenu>" +
            "<contextMenu idMso='ContextMenuRow'>" +
                "<button id='refreshDataR' label='refresh data (Ctl-Sh-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
                "<menuSeparator id='MySeparatorR' insertBeforeMso='Cut'/>" +
            "</contextMenu>" +
            "<contextMenu idMso='ContextMenuRowLayout'>" +
                "<button id='refreshDataRL' label='refresh data (Ctl-Sh-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
                "<menuSeparator id='MySeparatorRL' insertBeforeMso='Cut'/>" +
            "</contextMenu>" +
            "<contextMenu idMso='ContextMenuColumn'>" +
                "<button id='refreshDataZ' label='refresh data (Ctl-Sh-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
                "<menuSeparator id='MySeparatorZ' insertBeforeMso='Cut'/>" +
            "</contextMenu>" +
            "<contextMenu idMso='ContextMenuColumnLayout'>" +
                "<button id='refreshDataZL' label='refresh data (Ctl-Sh-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
                "<menuSeparator id='MySeparatorZL' insertBeforeMso='Cut'/>" +
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
            "<contextMenu idMso='ContextMenuListRangeLayout'>" +
                "<button id='refreshDataLL' label='refresh data (Ctl-Sh-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
                "<button id='gotoDBFuncLL' label='jump to DBFunc/target (Ctl-Sh-J)' imageMso='ConvertTextToTable' onAction='clickjumpButton' insertBeforeMso='Cut'/>" +
                "<menu id='createMenuLL' label='Insert/Edit DBFunc/DBModif' insertBeforeMso='Cut'>" +
                    "<button id='DBMapperLL' tag='DBMapper' label='DBMapper' imageMso='TableSave' onAction='clickCreateButton'/>" +
                    "<button id='DBSequenceLL' tag='DBSeqnce' label='DBSequence' imageMso='ShowOnNewButton' onAction='clickCreateButton'/>" +
                "</menu>" +
                "<menuSeparator id='MySeparatorLL' insertBeforeMso='Cut'/>" +
            "</contextMenu>" +
            "<contextMenu idMso='ContextMenuQuery'>" +
                "<button id='refreshDataQ' label='refresh data (Ctl-Sh-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
                "<button id='gotoDBFuncQ' label='jump to DBFunc/target (Ctl-Sh-J)' imageMso='ConvertTextToTable' onAction='clickjumpButton' insertBeforeMso='Cut'/>" +
                "<menu id='createMenuQ' label='Insert/Edit DBFunc/DBModif' insertBeforeMso='Cut'>" +
                    "<button id='DBMapperQ' tag='DBMapper' label='DBMapper' imageMso='TableSave' onAction='clickCreateButton'/>" +
                    "<button id='DBSequenceQ' tag='DBSeqnce' label='DBSequence' imageMso='ShowOnNewButton' onAction='clickCreateButton'/>" +
                "</menu>" +
                "<menuSeparator id='MySeparatorQ' insertBeforeMso='Cut'/>" +
            "</contextMenu>" +
            "<contextMenu idMso='ContextMenuQueryLayout'>" +
                "<button id='refreshDataQL' label='refresh data (Ctl-Sh-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
                "<button id='gotoDBFuncQL' label='jump to DBFunc/target (Ctl-Sh-J)' imageMso='ConvertTextToTable' onAction='clickjumpButton' insertBeforeMso='Cut'/>" +
                "<menu id='createMenuQL' label='Insert/Edit DBFunc/DBModif' insertBeforeMso='Cut'>" +
                    "<button id='DBMapperQL' tag='DBMapper' label='DBMapper' imageMso='TableSave' onAction='clickCreateButton'/>" +
                    "<button id='DBSequenceQL' tag='DBSeqnce' label='DBSequence' imageMso='ShowOnNewButton' onAction='clickCreateButton'/>" +
                "</menu>" +
                "<menuSeparator id='MySeparatorQL' insertBeforeMso='Cut'/>" +
            "</contextMenu>" +
        "</contextMenus></customUI>"
        Return customUIXml
    End Function

    ''' <summary>initialize the AdhocSQL ribbon combo-box entries</summary>
    Private Sub initAdhocSQLconfig()
        Dim customSettings As NameValueCollection = Nothing
        Try : customSettings = ConfigurationManager.GetSection("UserSettings") : Catch ex As Exception
            LogWarn("Error in getting User-settings (DBAddinUser.config) for AdhocSQLconfig entries: " + ex.Message)
        End Try
        AdHocSQLStrings = New List(Of String)
        ' getting User-settings might fail (formatting, etc)...
        If Not IsNothing(customSettings) Then
            For Each key As String In customSettings.AllKeys
                If Left(key, 11) = "AdhocSQLcmd" Then AdHocSQLStrings.Add(customSettings(key))
            Next
        End If
        selectedAdHocSQLIndex = 0
    End Sub

#Disable Warning IDE0060 ' Hide not used Parameter warning as this is very often the case with the below callbacks from the ribbon

    ''' <summary>context menu entry: create Chart for underlying Pivot table</summary>
    ''' <param name="control"></param>
    Public Sub clickCreateChart(control As CustomUI.IRibbonControl)
        Dim theChart As Excel.Shape
        Try
            theChart = ExcelDnaUtil.Application.ActiveSheet.Shapes.AddChart2(-1, Excel.XlChartType.xlColumnClustered)
            theChart.Chart.SetSourceData(ExcelDnaUtil.Application.ActiveCell)
        Catch ex As Exception
            UserMsg("Exception setting chart to underlying pivot table: " + ex.Message)
        End Try
    End Sub

    ''' <summary>the collection of all the configured/stored AdHocSQL Strings</summary>
    Private AdHocSQLStrings As Collections.Generic.List(Of String)
    ''' <summary>the index of the currently selected AdHocSQL String</summary>
    Private selectedAdHocSQLIndex As Integer

    ''' <summary>dialogBoxLauncher of DBAddin settings group: activate ad-hoc SQL query dialog</summary>
    ''' <param name="control"></param>
    Public Sub showDBAdHocSQLDBOX(control As CustomUI.IRibbonControl)
        showDBAdHocSQL(Nothing, "")
    End Sub

    ''' <summary>show Ad-hoc SQL Query editor</summary>
    ''' <param name="control"></param>
    Public Sub showDBAdHocSQL(control As CustomUI.IRibbonControl, selectedSQLText As String)
        Dim queryString As String = ""

        theAdHocSQLDlg = New AdHocSQL(selectedSQLText, AdHocSQLStrings.IndexOf(selectedSQLText))
        Dim dialogResult As Windows.Forms.DialogResult = theAdHocSQLDlg.ShowDialog()
        ' reflect potential change in environment...
        theRibbon.InvalidateControl("envDropDown")
        If dialogResult = System.Windows.Forms.DialogResult.OK Then 'OK is set when "transfer" is clicked
            ' "Transfer" was clicked: place SQL String into currently selected Cell/DBFunction and add to combo-box
            queryString = theAdHocSQLDlg.SQLText.Text
            If theAdHocSQLDlg.TransferType.Text = "Cell" Then
                Dim targetFormula As String = ExcelDnaUtil.Application.ActiveCell.Formula
                Dim srchdFunc As String = ""
                ' check whether there is any existing db function other than DBListFetch inside active cell
                For Each srchdFunc In {"DBSETQUERY", "DBROWFETCH", "DBLISTFETCH"}
                    If Left(UCase(targetFormula), Len(srchdFunc) + 2) = "=" + srchdFunc + "(" Then
                        ' for existing theFunction (DBSetQuery or DBRowFetch)...
                        Exit For
                    Else
                        srchdFunc = ""
                    End If
                Next
                If srchdFunc = "" Then
                    ' empty cell, just put query there
                    If targetFormula = "" Then
                        ExcelDnaUtil.Application.ActiveCell = queryString
                    Else
                        If QuestionMsg("Non-empty Cell with no DB function selected, should content be replaced?") = MsgBoxResult.Ok Then
                            ExcelDnaUtil.Application.ActiveCell = queryString
                        End If
                    End If
                Else
                    If QuestionMsg("Cell with DB function selected, should query be replaced?") = MsgBoxResult.Ok Then
                        ' db function, recreate with query inside
                        ' get the parts of the targeted function formula
                        Dim formulaParams As String = Mid$(targetFormula, Len(srchdFunc) + 3)
                        ' replace query-string in existing formula 
                        formulaParams = Left(formulaParams, Len(formulaParams) - 1)
                        Dim tempFormula As String = replaceDelimsWithSpecialSep(formulaParams, ",", """", "(", ")", vbTab)
                        Dim restFormula As String = Mid$(tempFormula, InStr(tempFormula, vbTab))
                        ' ..and put into active cell
                        ExcelDnaUtil.Application.ActiveCell.Formula = "=" + srchdFunc + "(""" + queryString + """" + Replace(restFormula, vbTab, ",") + ")"
                    End If
                End If
            ElseIf theAdHocSQLDlg.TransferType.Text = "Pivot" Then
                If ExcelDnaUtil.Application.ActiveCell.Formula <> "" Then
                    If QuestionMsg("Non-empty Cell selected, should content be replaced?") = MsgBoxResult.Cancel Then Exit Sub
                End If
                createPivotTable(ExcelDnaUtil.Application.ActiveCell)
                createFunctionsInCells(ExcelDnaUtil.Application.ActiveCell, {"RC", "=DBSetQuery(""" + queryString + ""","""",R[1]C)"})
            ElseIf theAdHocSQLDlg.TransferType.Text = "ListObject" Then
                If ExcelDnaUtil.Application.ActiveCell.Formula <> "" Then
                    If QuestionMsg("Non-empty Cell selected, should content be replaced?") = MsgBoxResult.Cancel Then Exit Sub
                End If
                createListObject(ExcelDnaUtil.Application.ActiveCell)
                createFunctionsInCells(ExcelDnaUtil.Application.ActiveCell, {"RC", "=DBSetQuery(""" + queryString + ""","""",RC[1])"})
            ElseIf theAdHocSQLDlg.TransferType.Text = "RowFetch" Then
                If ExcelDnaUtil.Application.ActiveCell.Formula <> "" Then
                    If QuestionMsg("Non-empty Cell selected, should content be replaced?") = MsgBoxResult.Cancel Then Exit Sub
                End If
                createFunctionsInCells(ExcelDnaUtil.Application.ActiveCell, {"RC", "=DBRowFetch(""" + queryString + ""","""",TRUE,R[1]C:R[1]C[10])"})
            ElseIf theAdHocSQLDlg.TransferType.Text = "ListFetch" Then
                If ExcelDnaUtil.Application.ActiveCell.Formula <> "" Then
                    If QuestionMsg("Non-empty Cell selected, should content be replaced?") = MsgBoxResult.Cancel Then Exit Sub
                End If
                createFunctionsInCells(ExcelDnaUtil.Application.ActiveCell, {"RC", "=DBListFetch(""" + queryString + ""","""",R[1]C,,,TRUE,TRUE,TRUE)"})
            End If
        ElseIf dialogResult = System.Windows.Forms.DialogResult.Cancel Then 'Cancel is set when "close" is clicked
            ' "Close" was clicked: only add SQL String to combo-box
            queryString = Strings.Trim(theAdHocSQLDlg.SQLText.Text)
        End If
        If selectedSQLText = "" Then selectedSQLText = queryString ' if ad-hoc SQL dialog was opened with mini button (fresh SQL) then set selectedSQLText here at least to the queryText to get the right index

        ' "Add" or "Transfer": store new sql command into AdHocSQLStrings and settings
        If Not AdHocSQLStrings.Contains(queryString) And Strings.Replace(queryString, " ", "") <> "" Then
            If QuestionMsg("Should the current command be added to the AdHocSQL dropdown?",, "AdHoc SQL Command") = MsgBoxResult.Ok Then
                ' add to AdHocSQLStrings
                AdHocSQLStrings.Add(queryString)
                selectedAdHocSQLIndex = AdHocSQLStrings.Count - 1
                ' change in or add to user settings
                setUserSetting("AdhocSQLcmd" + selectedAdHocSQLIndex.ToString(), queryString)
            Else
                queryString = ""
            End If
        Else ' just update selection index
            selectedAdHocSQLIndex = AdHocSQLStrings.IndexOf(selectedSQLText)
        End If
        If Strings.Replace(queryString, " ", "") <> "" And selectedAdHocSQLIndex >= 0 Then
            ' store the combo-box values for later...
            setUserSetting("AdHocSQLcmdEnv" + selectedAdHocSQLIndex.ToString(), theAdHocSQLDlg.EnvSwitch.SelectedIndex.ToString())
            setUserSetting("AdHocSQLcmdDB" + selectedAdHocSQLIndex.ToString(), theAdHocSQLDlg.Database.Text)
        End If
        setUserSetting("AdHocSQLTransferType", theAdHocSQLDlg.TransferType.Text)
        ' reflect changes in sql combo-box
        theRibbon.InvalidateControl("DBAdhocSQL")
    End Sub

    ''' <summary>get the text for the currently selected AdHocSQL string from the AdHocSQLStrings collection</summary>
    ''' <param name="control"></param>
    ''' <returns></returns>
    Public Function GetAdhocSQLText(control As CustomUI.IRibbonControl)
        If AdHocSQLStrings.Count > 0 And selectedAdHocSQLIndex >= 0 Then
            Return AdHocSQLStrings(selectedAdHocSQLIndex)
        Else
            Return ""
        End If
    End Function

    ''' <summary>get the total item count for the AdHocSQLStrings collection</summary>
    ''' <param name="control"></param>
    ''' <returns></returns>
    Public Function GetAdhocSQLItemCount(control As CustomUI.IRibbonControl) As Integer
        Return AdHocSQLStrings.Count
    End Function

    ''' <summary>get the labels for the AdHocSQL combo-box</summary>
    ''' <param name="control"></param>
    ''' <param name="index"></param>
    ''' <returns></returns>
    Public Function GetAdhocSQLItemLabel(control As CustomUI.IRibbonControl, index As Integer) As String
        If AdHocSQLStrings.Count > 0 Then
            Return AdHocSQLStrings(index)
        Else
            Return ""
        End If
    End Function

    ''' <summary>used for additional information</summary>
    ''' <param name="control"></param>
    ''' <returns></returns>
    Public Function getSuperTipInfo(ByRef control As CustomUI.IRibbonControl)
        getSuperTipInfo = ""
    End Function

    ''' <summary>display warning button icon on custom properties change if DBFskip is set...</summary>
    ''' <param name="control"></param>
    ''' <returns></returns>
    Public Function getCPropsImage(control As CustomUI.IRibbonControl) As String
        Dim actWb As Excel.Workbook
        Try : actWb = ExcelDnaUtil.Application.ActiveWorkbook : Catch ex As Exception
            LogWarn("Exception when trying to get the active workbook for displaying DBFskip status: " + ex.Message + ", this might be due to errors in the VBA Macros (missing references)")
            Return "AcceptTask"
        End Try
        If getCustPropertyBool("DBFskip", actWb) Then
            Return "DeclineTask"
        Else
            Return "AcceptTask"
        End If
    End Function

    ''' <summary>display warning icon on log button if warning has been logged...</summary>
    ''' <param name="control"></param>
    ''' <returns></returns>
    Public Function getLogsImage(control As CustomUI.IRibbonControl) As String
        If WarningIssued Then
            Return "IndexUpdate"
        Else
            Return "MailMergeStartLetters"
        End If
    End Function

    ''' <summary>display DBAddin custom properties/values in screen-tip of dialogBox launcher</summary>
    ''' <param name="control"></param>
    ''' <returns>screen-tip with the DBAddin custom properties/values</returns>
    Public Function getToggleCPropsScreentip(control As CustomUI.IRibbonControl) As String
        getToggleCPropsScreentip = ""
        Dim actWb As Excel.Workbook = Nothing
        Try : actWb = ExcelDnaUtil.Application.ActiveWorkbook : Catch ex As Exception
            getToggleCPropsScreentip = "Exception when trying to get the active workbook for displaying DBAddin custom properties/values: " + ex.Message + ", this might be due to errors in the VBA Macros (missing references)"
        End Try
        If actWb IsNot Nothing Then
            Try
                Dim docproperty As Microsoft.Office.Core.DocumentProperty
                For Each docproperty In actWb.CustomDocumentProperties
                    If Left$(docproperty.Name, 5) = "DBFC" Or docproperty.Name = "DBFskip" Or docproperty.Name = "doDBMOnSave" Or docproperty.Name = "DBFNoLegacyCheck" Then
                        getToggleCPropsScreentip += docproperty.Name + ":" + docproperty.Value.ToString() + vbCrLf
                    End If
                Next
            Catch ex As Exception
                getToggleCPropsScreentip += "exception when collecting DBAddin doc-properties: " + ex.Message
            End Try
        End If
    End Function

    ''' <summary>click on change props: show built-in properties dialog</summary>
    ''' <param name="control"></param>
    Public Sub showCProps(control As CustomUI.IRibbonControl)
        Dim actWb As Excel.Workbook = Nothing
        Try : actWb = ExcelDnaUtil.Application.ActiveWorkbook : Catch ex As Exception
            UserMsg("Exception when trying to get the active workbook for showing custom properties dialog: " + ex.Message + ", this might be due to errors in the VBA Macros (missing references)")
        End Try
        If actWb IsNot Nothing Then
            Try
                ExcelDnaUtil.Application.Dialogs(Excel.XlBuiltInDialog.xlDialogProperties).Show
            Catch ex As Exception
                UserMsg("The properties dialog can't be displayed, maybe you are in the formula/cell editor?", "Properties dialog display")
            End Try
            ' to check whether DBFskip has changed:
            theRibbon.InvalidateControl(control.Id)
        End If
    End Sub

    ''' <summary>converts a Mso Menu ID to a Drawing Image</summary>
    ''' <param name="idMso">the Mso Menu ID to be converted</param>
    ''' <returns>a System.Drawing.Image to be used by </returns>
    Private Function convertFromMso(idMso As String) As System.Drawing.Image
        Try
            Dim p As stdole.IPictureDisp = ExcelDnaUtil.Application.CommandBars.GetImageMso(idMso, 16, 16)
            Dim hPal As IntPtr = p.hPal
            convertFromMso = System.Drawing.Image.FromHbitmap(p.Handle, hPal)
        Catch ex As Exception
            ' in case above image fetching doesn't work then no image is displayed (the image parameter is still required for ContextMenuStrip.Items.Add !)
            convertFromMso = Nothing
        End Try
    End Function

    ''' <summary>menu strip for DB Modifier Definitions dialogBoxLauncher</summary>
    Private WithEvents ctMenuStrip As Windows.Forms.ContextMenuStrip

    ''' <summary>show DBModif definitions edit box</summary>
    ''' <param name="control"></param>
    Sub showDBModifEdit(control As CustomUI.IRibbonControl)
        Dim actWb As Excel.Workbook = Nothing
        Try : actWb = ExcelDnaUtil.Application.ActiveWorkbook : Catch ex As Exception
            UserMsg("Exception when trying to get the active workbook for show DBModif Editor: " + ex.Message + ", this might be due to errors in the VBA Macros (missing references)")
            Exit Sub
        End Try
        ' only show dialog if there is a workbook and it has the relevant custom XML part.
        If actWb IsNot Nothing AndAlso actWb.CustomXMLParts.SelectByNamespace("DBModifDef").Count > 0 Then
            Dim CustomXmlParts As Object = actWb.CustomXMLParts.SelectByNamespace("DBModifDef")
            ' check if any DBModifier exist below root node, only if at least one is defined, open dialog
            If CustomXmlParts(1).SelectNodes("/ns0:root/*").Count > 0 Then
                Dim theEditDBModifDefDlg As New EditDBModifDef()
                If theEditDBModifDefDlg.ShowDialog() = System.Windows.Forms.DialogResult.OK Then getDBModifDefinitions(actWb)
                Exit Sub
            End If
        End If
        ' no existing DB modifier definitions found, offer to create new ones...
        ctMenuStrip = New Windows.Forms.ContextMenuStrip()
        Dim ptLowerLeft As Drawing.Point = System.Windows.Forms.Cursor.Position
        ctMenuStrip.Items().Add("No DBModifier definitions exist in current workbook, do you want to add")
        ctMenuStrip.Items().Add("a DBMapper", convertFromMso("TableSave"), AddressOf ctMenuStrip_Click)
        ctMenuStrip.Items().Add("a DBAction", convertFromMso("TableIndexes"), AddressOf ctMenuStrip_Click)
        ctMenuStrip.Items().Add("a DBSequence", convertFromMso("ShowOnNewButton"), AddressOf ctMenuStrip_Click)
        ctMenuStrip.Show(ptLowerLeft)
    End Sub

    ''' <summary>called after selecting DBMapper/DBAction/DBSequence from the dialogBoxLauncher</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ctMenuStrip_Click(sender As Object, e As EventArgs)
        createDBModif(Replace(Replace(sender.ToString(), "a ", ""), "DBSequence", "DBSeqnce"))
    End Sub

    ''' <summary>toggle PreventRefresh button</summary>
    ''' <param name="control"></param>
    Sub showTogglePreventRefresh(control As CustomUI.IRibbonControl)
        If My.Computer.Keyboard.CtrlKeyDown Then
            If preventRefreshFlagColl.ContainsKey(ExcelDnaUtil.Application.ActiveWorkbook.Name) Then
                Functions.preventRefreshFlagColl(ExcelDnaUtil.Application.ActiveWorkbook.Name) = Not Functions.preventRefreshFlagColl(ExcelDnaUtil.Application.ActiveWorkbook.Name)
            Else
                Functions.preventRefreshFlagColl.Add(ExcelDnaUtil.Application.ActiveWorkbook.Name, True)
            End If
        ElseIf My.Computer.Keyboard.ShiftKeyDown Then
            If preventRefreshFlagColl.ContainsKey(ExcelDnaUtil.Application.ActiveWorkbook.Name) Then Functions.preventRefreshFlagColl.Remove(ExcelDnaUtil.Application.ActiveWorkbook.Name)
        ElseIf Not preventRefreshFlagColl.ContainsKey(ExcelDnaUtil.Application.ActiveWorkbook.Name) Then
            Functions.preventRefreshFlag = Not Functions.preventRefreshFlag
        End If
        ' update state of preventRefreshFlag in screen-tip
        theRibbon.InvalidateControl(control.Id)
    End Sub

    ''' <summary>display state of PreventRefresh in screen-tip of button</summary>
    ''' <param name="control"></param>
    ''' <returns>screen-tip and the state of PreventRefresh</returns>
    Public Function getTogglePreventRefreshScreentip(control As CustomUI.IRibbonControl) As String
        Dim preventStatus, onlyForThisWB As Boolean
        getStatusAndExtent(preventStatus, onlyForThisWB)
        Dim toWhichExtent As String = IIf(onlyForThisWB, "for workbook '" + ExcelDnaUtil.Application.ActiveWorkbook.Name + "'", "globally")
        Return "Refreshing of DB Functions is currently " + IIf(preventStatus, "blocked ", "enabled ") + toWhichExtent + ", click to set/reset preventRefresh globally, Ctrl+click to set/reset preventRefresh on workbook, Shift+click to remove preventRefresh setting on Workbook"
    End Function

    ''' <summary>display state of PreventRefresh in icon of button</summary>
    ''' <param name="control"></param>
    ''' <returns>screen-tip and the state of PreventRefresh</returns>
    Public Function getTogglePreventRefreshImage(control As CustomUI.IRibbonControl) As String
        Dim preventStatus, onlyForThisWB As Boolean
        getStatusAndExtent(preventStatus, onlyForThisWB)
        Return IIf(preventStatus, IIf(onlyForThisWB, "SheetDelete", "CalculateFull"), IIf(onlyForThisWB, "CalculateSheet", "Calculator"))
    End Function

    ''' <summary>common fetching function for getting the Status and the extent of the preventRefresh flag</summary>
    ''' <param name="preventStatus">returns the currently set prevention status</param>
    ''' <param name="onlyForThisWB">returns the currently set flag if only for this Workbook</param>
    Private Sub getStatusAndExtent(ByRef preventStatus As Boolean, ByRef onlyForThisWB As Boolean)
        If preventRefreshFlagColl.ContainsKey(ExcelDnaUtil.Application.ActiveWorkbook.Name) Then
            preventStatus = Functions.preventRefreshFlagColl(ExcelDnaUtil.Application.ActiveWorkbook.Name)
            onlyForThisWB = True
        ElseIf preventRefreshFlagColl.Count > 0 Then
            preventStatus = False
            onlyForThisWB = True
        Else
            preventStatus = Functions.preventRefreshFlag
            onlyForThisWB = False
        End If
    End Sub

    ''' <summary>toggle design mode button</summary>
    ''' <param name="control"></param>
    Sub showToggleDesignMode(control As CustomUI.IRibbonControl)
        Dim cbrs As Object = ExcelDnaUtil.Application.CommandBars
        If cbrs IsNot Nothing AndAlso cbrs.GetEnabledMso("DesignMode") Then
            cbrs.ExecuteMso("DesignMode")
        Else
            UserMsg("Couldn't toggle design mode, because Design mode command-bar button is not available (no button?)", "DBAddin toggle Design mode", MsgBoxStyle.Exclamation)
        End If
        ' update state of design mode in screen-tip
        theRibbon.InvalidateControl(control.Id)
    End Sub

    ''' <summary>display state of design mode in screen-tip of button</summary>
    ''' <param name="control"></param>
    ''' <returns>screen-tip and the state of design mode</returns>
    Public Function getToggleDesignScreentip(control As CustomUI.IRibbonControl) As String
        Dim cbrs As Object = ExcelDnaUtil.Application.CommandBars
        If cbrs IsNot Nothing AndAlso cbrs.GetEnabledMso("DesignMode") Then
            Return "Design mode is currently " + IIf(cbrs.GetPressedMso("DesignMode"), "on !", "off !")
        Else
            Return "Design mode command-bar button not available (no button on sheet)"
        End If
    End Function

    ''' <summary>display state of design mode in icon of button</summary>
    ''' <param name="control"></param>
    ''' <returns>screen-tip and the state of design mode</returns>
    Public Function getToggleDesignImage(control As CustomUI.IRibbonControl) As String
        Dim cbrs As Object = ExcelDnaUtil.Application.CommandBars
        If cbrs IsNot Nothing AndAlso cbrs.GetEnabledMso("DesignMode") Then
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
    Public Function GetEnvItemCount(control As CustomUI.IRibbonControl) As Integer
        Return environdefs.Length
    End Function

    ''' <summary>for environment dropdown to get the label of the entries</summary>
    ''' <returns></returns>
    Public Function GetEnvItemLabel(control As CustomUI.IRibbonControl, index As Integer) As String
        Return environdefs(index)
    End Function

    ''' <summary>for environment dropdown to get the ID of the entries</summary>
    ''' <returns></returns>
    Public Function GetEnvItemID(control As CustomUI.IRibbonControl, index As Integer) As String
        Return environdefs(index)
    End Function

    ''' <summary>after selection of environment (using selectEnvironment) used to return the selected environment</summary>
    ''' <returns></returns>
    Public Function GetSelectedEnvironment(control As CustomUI.IRibbonControl) As Integer
        Return selectedEnvironment
    End Function

    ''' <summary>tool-tip for the environment select drop down</summary>
    ''' <param name="control"></param>
    ''' <returns>the tool-tip</returns>
    Public Function GetEnvSelectedTooltip(control As CustomUI.IRibbonControl) As String
        If fetchSettingBool("DontChangeEnvironment", "False") Then
            Return "DontChangeEnvironment is set, therefore changing the Environment is prevented !"
        Else
            Return "configured for Database Access in Addin config %appdata%\Microsoft\Addins\DBaddin.xll.config (or referenced central/user setting)"
        End If
    End Function

    ''' <summary>whether to enable environment selection drop down</summary>
    ''' <param name="control"></param>
    ''' <returns>true if enabled</returns>
    Public Function GetEnvEnabled(control As CustomUI.IRibbonControl) As Integer
        Return Not fetchSettingBool("DontChangeEnvironment", "False")
    End Function

    ''' <summary>Choose environment (configured in settings with ConstConnString(N), ConfigStoreFolder(N))</summary>
    ''' <param name="control"></param>
    Public Sub selectEnvironment(control As CustomUI.IRibbonControl, id As String, index As Integer)
        selectedEnvironment = index
        initSettings()
        ' provide a chance to reconnect when switching environment...
        conn = Nothing
        Dim actWb As Excel.Workbook = Nothing
        Try : actWb = ExcelDnaUtil.Application.ActiveWorkbook : Catch ex As Exception
            UserMsg("Exception when trying to get the active workbook for selecting environment: " + ex.Message + ", this might be due to errors in the VBA Macros (missing references)")
        End Try
        If actWb IsNot Nothing Then
            Dim retval As MsgBoxResult = QuestionMsg("ConstConnString:" + ConstConnString + vbCrLf + "ConfigStoreFolder:" + ConfigFiles.ConfigStoreFolder + vbCrLf + vbCrLf + "Refresh DBFunctions in active workbook to see effects?", MsgBoxStyle.OkCancel, "Changed environment to: " + fetchSetting("ConfigName" + env(), ""))
            If retval = MsgBoxResult.Ok Then refreshDBFunctions(actWb)
        Else
            UserMsg("ConstConnString:" + ConstConnString + vbCrLf + "ConfigStoreFolder:" + ConfigFiles.ConfigStoreFolder, "Changed environment to: " + fetchSetting("ConfigName" + env(), ""), MsgBoxStyle.Information)
        End If
    End Sub

    ''' <summary>show xll standard config (AppSetting), central config (referenced by App Settings file attr) or user config (referenced by CustomSettings configSource attr)</summary>
    ''' <param name="control"></param>
    Public Sub showAddinConfig(control As CustomUI.IRibbonControl)
        ' if settings (addin, user, central) should not be displayed according to setting then exit...
        If InStr(fetchSetting("disableSettingsDisplay", ""), control.Id) > 0 Then
            UserMsg("Display of " + control.Id + " settings disabled !", "DBAddin Settings disabled", MsgBoxStyle.Information)
            Exit Sub
        End If
        Dim theEditDBModifDefDlg As New EditDBModifDef()
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
        initSettings()
        initAdhocSQLconfig()
        ' also display in ribbon
        theRibbon.Invalidate()
    End Sub

    ''' <summary>dialogBoxLauncher of DBAddin settings group: activate about box</summary>
    ''' <param name="control"></param>
    Public Sub showAbout(control As CustomUI.IRibbonControl)
        Dim myAbout As New AboutBox
        myAbout.ShowDialog()
        ' if quitting was chosen, then quit excel here..
        If myAbout.quitExcelAfterwards Then ExcelDnaUtil.Application.Quit()
    End Sub

    ''' <summary>on demand, refresh the DB Config tree</summary>
    ''' <param name="control"></param>
    Public Sub refreshDBConfigTree(control As CustomUI.IRibbonControl)
        initSettings()
        ConfigFiles.createConfigTreeMenu()
        UserMsg("refreshed DB Config Tree Menu", "DBAddin: refresh Config tree...", MsgBoxStyle.Information)
        theRibbon.Invalidate()
    End Sub

    ''' <summary>get DB Config Menu from File</summary>
    ''' <param name="control"></param>
    ''' <returns></returns>
    Public Function getDBConfigMenu(control As CustomUI.IRibbonControl) As String
        If ConfigFiles.ConfigMenuXML = vbNullString Then ConfigFiles.createConfigTreeMenu()
        Return ConfigFiles.ConfigMenuXML
    End Function

    ''' <summary>load config if config tree menu end-button has been activated (path to config xcl file is in control.Tag), also provides access to documentation if provided in setting ConfigDocQuery</summary>
    ''' <param name="control"></param>
    Public Sub getConfig(control As CustomUI.IRibbonControl)
        If My.Computer.Keyboard.CtrlKeyDown Or My.Computer.Keyboard.ShiftKeyDown Then
            ConfigFiles.showTheDocs(control.Tag)
        Else
            ConfigFiles.loadConfig(control.Tag)
        End If
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
            If Not DBModifDefColl.ContainsKey(control.Id) Then Return ""
            Dim DBModifTypeName As String = IIf(control.Id = "DBSeqnce", "DBSequence", IIf(control.Id = "DBMapper", "DB Mapper", IIf(control.Id = "DBAction", "DB Action", "undefined DBModifTypeName")))
            For Each nodeName As String In DBModifDefColl(control.Id).Keys
                Dim descName As String = IIf(nodeName = control.Id, "Unnamed " + DBModifTypeName, Replace(nodeName, DBModifTypeName, ""))
                Dim imageMsoStr As String = IIf(control.Id = "DBSeqnce", "ShowOnNewButton", IIf(control.Id = "DBMapper", "TableSave", IIf(control.Id = "DBAction", "TableIndexes", "undefined imageMso")))
                Dim superTipStr As String = IIf(control.Id = "DBSeqnce", "executes " + DBModifTypeName + " defined in " + nodeName, IIf(control.Id = "DBMapper", "stores data defined in DBMapper (named " + nodeName + ") range on " + DBModifDefColl(control.Id).Item(nodeName).getTargetRangeAddress(), IIf(control.Id = "DBAction", "executes Action defined in DBAction (named " + nodeName + ") range on " + DBModifDefColl(control.Id).Item(nodeName).getTargetRangeAddress(), "undefined superTip")))
                xmlString = xmlString + "<button id='_" + nodeName + "' label='do " + descName + "' imageMso='" + imageMsoStr + "' onAction='DBModifClick' tag='" + control.Id + "' screentip='do " + DBModifTypeName + ": " + descName + "' supertip='" + superTipStr + "' />"
            Next
            xmlString += "</menu>"
            Return xmlString
        Catch ex As Exception
            UserMsg("Exception caught while building xml: " + ex.Message)
            Return ""
        End Try
    End Function

    ''' <summary>show a screen-tip for the dynamic DBMapper/DBAction/DBSequence Menus (also showing the ID behind)</summary>
    ''' <param name="control"></param>
    ''' <returns>the screen-tip</returns>
    Public Function getDBModifScreentip(control As CustomUI.IRibbonControl) As String
        Return "Select DBModifier to store/do action/do sequence (" + control.Id + ")"
    End Function

    ''' <summary>to show the DBModif sheet button only if it was collected...</summary>
    ''' <param name="control"></param>
    ''' <returns>true if to be displayed</returns>
    Public Function getDBModifMenuVisible(control As CustomUI.IRibbonControl) As Boolean
        Try
            Return DBModifDefColl.ContainsKey(control.Id)
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>DBModif button activated, do DB Mapper/DB Action/DB Sequence or define existing (CtrlKey pressed)...</summary>
    ''' <param name="control"></param>
    Public Sub DBModifClick(control As CustomUI.IRibbonControl)
        Dim actWb As Excel.Workbook = Nothing
        Try : actWb = ExcelDnaUtil.Application.ActiveWorkbook : Catch ex As Exception
            UserMsg("Exception when trying to get the active workbook for DB Modifier activation: " + ex.Message + ", this might be due to errors in the VBA Macros (missing references)")
        End Try
        ' reset non-interactive messages (used for VBA invocations) and hadError for interactive invocations
        nonInteractiveErrMsgs = "" : DBModifHelper.hadError = False
        Dim nodeName As String = Right(control.Id, Len(control.Id) - 1)
        If Not ExcelDnaUtil.Application.CommandBars.GetEnabledMso("FileNewDefault") Then
            UserMsg("Cannot execute DB Modifier while cell editing active !", "DB Modifier execution", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Try
            If My.Computer.Keyboard.CtrlKeyDown And My.Computer.Keyboard.ShiftKeyDown Then
                createDBModif(control.Tag, targetDefName:=nodeName)
            Else
                ' DB sequence actions (the sequence to be done) are stored directly in DBMapperDefColl, so different invocation here
                If Not (actWb.ReadOnlyRecommended And actWb.ReadOnly) Then
                    DBModifDefColl(control.Tag).Item(nodeName).doDBModif()
                Else
                    UserMsg("ReadOnlyRecommended is set on active workbook (being readonly), therefore all DB Modifiers are disabled !", "DB Modifier execution", MsgBoxStyle.Exclamation)
                End If
            End If
        Catch ex As Exception
            UserMsg("Exception: " + ex.Message + ",control.Tag:" + control.Tag + ",nodeName:" + nodeName, "DBModif Click")
        End Try
    End Sub

    ''' <summary>context menu entry refreshData: refresh Data in db function (if area or cell selected) or all db functions</summary>
    ''' <param name="control"></param>
    Public Sub clickrefreshData(control As CustomUI.IRibbonControl)
        refreshData()
    End Sub

    ''' <summary>context menu entry gotoDBFunc: jumps from DB function to data area and back</summary>
    ''' <param name="control"></param>
    Public Sub clickjumpButton(control As CustomUI.IRibbonControl)
        jumpButton()
    End Sub

    ''' <summary>check/purge name tool button, purge names used for db-functions from workbook</summary>
    ''' <param name="control"></param>
    Public Sub clickcheckpurgetoolbutton(control As CustomUI.IRibbonControl)
        checkpurgeNames()
        checkHiddenExcelInstance()
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
        Dim theDBSheetCreateForm As New DBSheetCreateForm
        theDBSheetCreateForm.Show()
    End Sub

    ''' <summary>context menu entries in Insert/Edit DBFunc/DBModif and Assign DBSheet: create DB function or DB Modification definition</summary>
    ''' <param name="control"></param>
    Public Sub clickCreateButton(control As CustomUI.IRibbonControl)
        ' check for existing DBMapper or DBAction definition and allow exit
        Dim activeCellDBModifName As String = getDBModifNameFromRange(ExcelDnaUtil.Application.ActiveCell)
        Dim activeCellDBModifType As String = Left(activeCellDBModifName, 8)
        If (activeCellDBModifType = "DBMapper" Or activeCellDBModifType = "DBAction") And activeCellDBModifType <> control.Tag And control.Tag <> "DBSeqnce" Then
            UserMsg("Active Cell already contains definition for a " + activeCellDBModifType + ", inserting " + IIf(control.Tag = "DBSetQueryPivot" Or control.Tag = "DBSetQueryListObject", "DBSetQuery", control.Tag) + " here will cause trouble !", "Inserting not allowed")
            Exit Sub
        End If
        If control.Tag = "DBListFetch" Then
            createFunctionsInCells(ExcelDnaUtil.Application.ActiveCell, {"RC", "=DBListFetch("""","""",R[1]C,,,TRUE,TRUE,TRUE)"})
        ElseIf control.Tag = "DBRowFetch" Then
            createFunctionsInCells(ExcelDnaUtil.Application.ActiveCell, {"RC", "=DBRowFetch("""","""",TRUE,R[1]C:R[1]C[10])"})
        ElseIf control.Tag = "DBSetQueryPivot" Then
            ' first create a dummy pivot table
            createPivotTable(ExcelDnaUtil.Application.ActiveCell)
            ' then create the DBSetQuery assigning the (yet to be filled) query to the above list-object
            createFunctionsInCells(ExcelDnaUtil.Application.ActiveCell, {"RC", "=DBSetQuery("""","""",R[1]C)"})
        ElseIf control.Tag = "DBSetQueryListObject" Then
            ' first create a dummy ListObject
            createListObject(ExcelDnaUtil.Application.ActiveCell)
            ' then create the DBSetQuery assigning the (yet to be filled) query to the above list-object
            createFunctionsInCells(ExcelDnaUtil.Application.ActiveCell, {"RC", "=DBSetQuery("""","""",RC[1])"})
        ElseIf control.Tag = "DBMapper" Or control.Tag = "DBAction" Or control.Tag = "DBSeqnce" Then
            If activeCellDBModifType = control.Tag Then  ' edit existing definition
                createDBModif(control.Tag, targetDefName:=activeCellDBModifName)
            Else                                         ' create new definition
                createDBModif(control.Tag)
            End If
        ElseIf control.Tag = "DBSetPowerQuery" Then
            Dim wbQueries As Object = Nothing
            Try : wbQueries = ExcelDnaUtil.Application.ActiveWorkbook.Queries
            Catch ex As Exception
                UserMsg("Error getting power queries: " + ex.Message + ", this might be due to errors in the VBA Macros (missing references)")
                Exit Sub
            End Try
            If IsNothing(wbQueries) Or wbQueries.Count = 0 Then
                UserMsg("No power queries available...", "DBAddin getting power queries", MsgBoxStyle.Exclamation)
                Exit Sub
            End If
            ctMenuStrip2 = New Windows.Forms.ContextMenuStrip()
            Dim ptLowerLeft As Drawing.Point = System.Windows.Forms.Cursor.Position
            ctMenuStrip2.Items().Add("Select a power query below to be referenced by DBSetPowerQuery (hold Ctrl to restore the last query)").Enabled = False
            For Each qry As Object In wbQueries
                ctMenuStrip2.Items().Add(qry.Name, convertFromMso("TableExcelSpreadsheetInsert"), AddressOf ctMenuStrip2_Click)
            Next
            ctMenuStrip2.Show(ptLowerLeft)
        End If
    End Sub

    ''' <summary>menu strip for displaying available power queries in current workbook</summary>
    Private WithEvents ctMenuStrip2 As Windows.Forms.ContextMenuStrip

    ''' <summary>called after selecting a power-query for insert DBSetPowerQuery</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ctMenuStrip2_Click(sender As Object, e As EventArgs)
        Dim actWb As Excel.Workbook = Nothing
        Try : actWb = ExcelDnaUtil.Application.ActiveWorkbook : Catch ex As Exception
            UserMsg("Exception when trying to get the active workbook for power query selection: " + ex.Message + ", this might be due to errors in the VBA Macros (missing references)")
        End Try
        ' restore previously stored query with Ctrl..
        If My.Computer.Keyboard.CtrlKeyDown Then
            actWb.Queries(sender.ToString()).Formula = Functions.queryBackupColl(sender.ToString())
            UserMsg("Last power query restored for " + sender.ToString())
            Exit Sub
        End If
        Dim theFormulaStr As String() = actWb.Queries(sender.ToString()).Formula.ToString().Split(vbCrLf)
        Dim i As Integer = 1
        Dim currentCell As Excel.Range = ExcelDnaUtil.Application.ActiveCell
        Functions.avoidRequeryDuringEdit = True
        For Each formulaPart As String In theFormulaStr
            If currentCell.Offset(i, 0).Value <> "" Then
                currentCell.Offset(i, 0).Select()
                If QuestionMsg("Cell not empty (would be overwritten), continue?") <> MsgBoxResult.Ok Then Exit Sub
            End If
            currentCell.Offset(i, 0).Value = formulaPart.Replace(vbLf, "")
            i += 1
        Next
        createFunctionsInCells(currentCell, {"RC", "=DBSetPowerQuery(R[1]C:R[" + (i - 1).ToString() + "]C,""" + sender.ToString() + """)"})
        Functions.avoidRequeryDuringEdit = False
    End Sub

    ''' <summary>clicked Assign DBSheet: create DB Mapper with CUD Flags</summary>
    ''' <param name="control"></param>
    Public Sub clickAssignDBSheet(control As CustomUI.IRibbonControl)
        Dim actWb As Excel.Workbook = Nothing
        Try : actWb = ExcelDnaUtil.Application.ActiveWorkbook : Catch ex As Exception
            LogWarn("Exception when trying to get the active workbook for assigning DBSheet: " + ex.Message + ", this might be due to errors in the VBA Macros (missing references)")
        End Try
        If actWb IsNot Nothing Then
            DBSheetConfig.createDBSheet()
        Else
            UserMsg("Cannot assign DBSheet DB Mapper as there is no Workbook active !", "DB Sheet Assignment", MsgBoxStyle.Exclamation)
        End If
    End Sub

#Enable Warning IDE0060
End Class


''' <summary>global ribbon variable and procedures for MenuHandler</summary>
Public Module MenuHandlerGlobals
    ''' <summary>reference object for the ribbon</summary>
    Public theRibbon As CustomUI.IRibbonUI
    ''' <summary>This needs to be public accessible for click handler to pass results into SQLText control</summary>
    Public theAdHocSQLDlg As AdHocSQL


    ''' <summary>refresh DB Functions (and - if called from outside any db function area - all other external data ranges)</summary>
    <ExcelCommand(Name:="refreshData")>
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
                        If Not fetchSettingBool("AvoidUpdateQueryTables_Refresh", "False") Then
                            For Each qrytbl As Excel.QueryTable In ws.QueryTables
                                ' avoid double refreshing of query tables
                                If getUnderlyingDBNameFromRange(qrytbl.ResultRange) = "" Then qrytbl.Refresh()
                            Next
                        End If
                        If Not fetchSettingBool("AvoidUpdatePivotTables_Refresh", "False") Then
                            For Each pivottbl As Excel.PivotTable In ws.PivotTables
                                ' avoid double refreshing of dbsetquery list objects
                                If getUnderlyingDBNameFromRange(pivottbl.TableRange1) = "" Then pivottbl.PivotCache.Refresh()
                            Next
                        End If
                        If Not fetchSettingBool("AvoidUpdateListObjects_Refresh", "False") Then
                            For Each listobj As Excel.ListObject In ws.ListObjects
                                ' avoid double refreshing of dbsetquery list objects
                                If getUnderlyingDBNameFromRange(listobj.Range) = "" Then listobj.QueryTable.Refresh()
                            Next
                        End If
                    Next
                    If Not fetchSettingBool("AvoidUpdateLinks_Refresh", "False") Then
                        For Each linksrc As String In actWb.LinkSources(Excel.XlLink.xlExcelLinks)
                            ' first check file existence to not run into excel opening a file dialog for missing files....
                            If System.IO.File.Exists(linksrc) Then
                                actWb.UpdateLink(Name:=linksrc, Type:=Excel.XlLink.xlExcelLinks)
                            Else
                                LogWarn("File " + linksrc + " doesn't exist, couldn't update excel links to it!")
                            End If
                        Next
                    End If
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
    <ExcelCommand(Name:="jumpButton")>
    Public Sub jumpButton()
        If checkMultipleDBRangeNames(ExcelDnaUtil.Application.ActiveCell) Then
            UserMsg("Multiple hidden DB Function names in selected cell (making 'jump' ambiguous/impossible), please use purge names tool!", "DB-Addin Jump")
            Exit Sub
        End If
        ' show underlying name(s) of dbfunc target areas
        If My.Computer.Keyboard.CtrlKeyDown Or My.Computer.Keyboard.ShiftKeyDown Then
            Dim nm As Excel.Name
            Dim rng, testRng As Excel.Range
            Dim underlyingDBName As String = getUnderlyingDBNameFromRange(ExcelDnaUtil.Application.ActiveCell)
            Dim underlyingNameFromRange As String = ""
            Try
                testRng = Nothing
                Try : testRng = ExcelDnaUtil.Application.Range(underlyingDBName) : Catch ex As Exception : End Try
                For Each nm In ExcelDnaUtil.Application.ActiveWorkbook.Names
                    rng = Nothing
                    Try : rng = nm.RefersToRange : Catch ex As Exception : End Try
                    If rng IsNot Nothing Then
                        Dim testQryTblName As String = " " ' space cannot exist in a name, so this default will never be found
                        Try : Dim testQryTbl As Excel.QueryTable = ExcelDnaUtil.Application.ActiveCell.QueryTable : testQryTblName = testQryTbl.Name : Catch ex As Exception : End Try
                        If testRng IsNot Nothing AndAlso nm.Visible AndAlso InStr(nm.Name, "DBFtarget") = 0 AndAlso InStr(nm.Name, "DBFsource") = 0 AndAlso nm.Name <> "'" + rng.Parent.Name + "'!_FilterDatabase" AndAlso InStr(nm.Name, testQryTblName) = 0 AndAlso testRng.Cells(1, 1).Address = rng.Cells(1, 1).Address And testRng.Parent Is rng.Parent Then
                            underlyingNameFromRange += nm.Name + ","
                        End If
                    End If
                Next
            Catch ex As Exception : End Try
            If underlyingNameFromRange <> "" Then
                UserMsg("Underlying name(s): " + Left(underlyingNameFromRange, Len(underlyingNameFromRange) - 1), "DBAddin Information", MsgBoxStyle.Information)
            End If
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

    ''' <summary>check for multiple excel instances</summary>
    Public Sub checkHiddenExcelInstance()
        Try
            If Diagnostics.Process.GetProcessesByName("Excel").Length > 1 Then
                For Each procinstance As Diagnostics.Process In Diagnostics.Process.GetProcessesByName("Excel")
                    If procinstance.MainWindowTitle = "" Then
                        UserMsg("Another hidden excel instance detected (PID: " + procinstance.Id + "), this may cause problems with querying DB Data")
                    End If
                Next
            End If
        Catch ex As Exception : End Try
    End Sub

    ''' <summary>check if multiple (hidden, containing DBtarget or DBsource) DB Function names exist in theRange</summary>
    ''' <param name="theRange">given range</param>
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

End Module