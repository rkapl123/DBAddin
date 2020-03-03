Imports ExcelDna.Integration.CustomUI
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices


''' <summary>handles all Menu related aspects (context menu for building/refreshing, "DBAddin"/"Load Config" tree menu for retrieving stored configuration files</summary>
<ComVisible(True)>
Public Class MenuHandler
    Inherits ExcelRibbon

    ''' <summary>callback after Excel loaded the Ribbon, used to initialize data for the Ribbon</summary>
    Public Sub ribbonLoaded(theRibbon As IRibbonUI)
        Globals.theRibbon = theRibbon
    End Sub

    ''' <summary>creates the Ribbon (only at startup). any changes to the ribbon can only be done via dynamic menus</summary>
    ''' <returns></returns>
    Public Overrides Function GetCustomUI(RibbonID As String) As String
        ' Ribbon definition XML
        Dim customUIXml As String = "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='ribbonLoaded' ><ribbon><tabs><tab id='DBaddinTab' label='DB Addin'>"
        ' DBAddin Group: environment choice, DBConfics selection tree, purge names tool button and dialogBoxLauncher for AboutBox
        customUIXml += "<group id='DBAddinGroup' label='General settings'>" +
              "<dropDown id='envDropDown' label='Environment:' sizeString='123456789012345' getSelectedItemIndex='GetSelItem' getItemCount='GetItemCount' getItemID='GetItemID' getItemLabel='GetItemLabel' onAction='selectItem'/>" +
              "<dynamicMenu id='DBConfigs' size='normal' label='DB Configs' imageMso='QueryShowTable' screentip='DB Function Configuration Files quick access' getContent='getDBConfigMenu'/>" +
              "<buttonGroup id='buttonGroup'>" +
              "<button id='purgetool' label='purge tool' screentip='purges underlying DBtarget/DBsource Names' imageMso='BorderErase' onAction='clickpurgetoolbutton'/>" +
              "<button id='showLog' label='show log' screentip='shows Database Addins Diagnostic Display' imageMso='ZoomOnePage' onAction='clickShowLog'/>" +
              "</buttonGroup>" +
              "<dialogBoxLauncher><button id='dialog' label='About DBAddin' onAction='showAbout' tag='3' screentip='Show Aboutbox with help, version information, homepage and access to log'/></dialogBoxLauncher></group>"
        ' DBModif Group: maximum three DBModif types possible (depending on existence in current workbook): 
        customUIXml += "<group id='DBModifGroup' label='Store DBModif Data'>"
        For Each DBModifType As String In {"DBSeqnce", "DBMapper", "DBAction"}
            customUIXml += "<dynamicMenu id='" + DBModifType + "' " +
                                            "size='large' getLabel='getDBModifTypeLabel' imageMso='ApplicationOptionsDialog' " +
                                            "getScreentip='getDBModifScreentip' getContent='getDBModifMenuContent' getVisible='getDBModifMenuVisible'/>"
        Next
        customUIXml += "<dialogBoxLauncher><button id='designmode' label='DBModif design' onAction='toggleDesignMode' tag='4' getScreentip='getToggleDesignScreentip'/></dialogBoxLauncher></group></tab></tabs></ribbon>"
        ' Context menus for refresh, jump and creation: in cell, row, column and ListRange (area of ListObjects)
        customUIXml += "<contextMenus>" +
         "<contextMenu idMso ='ContextMenuCell'>" +
         "<button id='refreshDataC' label='refresh data (Ctl-Sh-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
         "<button id='gotoDBFuncC' label='jump to DBFunc/target (Ctl-Sh-J)' imageMso='ConvertTextToTable' onAction='clickjumpButton' insertBeforeMso='Cut'/>" +
         "<button id='DeleteRowC' label='delete Row (Ctl-Sh-D)' imageMso='SlicerDelete' onAction='deleteRowButton' insertBeforeMso='Cut'/>" +
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
         "<button id='DeleteRowCL' label='delete Row (Ctl-Sh-D)' imageMso='SlicerDelete' onAction='deleteRowButton' insertBeforeMso='Cut'/>" +
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
         "<button id='DeleteRowR' label='delete Row (Ctl-Sh-D)' imageMso='SlicerDelete' onAction='deleteRowButton' insertBeforeMso='Cut'/>" +
         "<menuSeparator id='MySeparatorR' insertBeforeMso='Cut'/>" +
         "</contextMenu>" +
         "<contextMenu idMso='ContextMenuColumn'>" +
         "<button id='refreshDataZ' label='refresh data (Ctl-Sh-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
         "<menuSeparator id='MySeparatorZ' insertBeforeMso='Cut'/>" +
         "</contextMenu>" +
         "<contextMenu idMso='ContextMenuListRange'>" +
             "<button id='refreshDataL' label='refresh data (Ctl-Sh-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
             "<button id='gotoDBFuncL' label='jump to DBFunc/target (Ctl-Sh-J)' imageMso='ConvertTextToTable' onAction='clickjumpButton' insertBeforeMso='Cut'/>" +
             "<button id='DeleteRowL' label='delete Row (Ctl-Sh-D)' imageMso='SlicerDelete' onAction='deleteRowButton' insertBeforeMso='Cut'/>" +
             "<menu id='createMenuL' label='Insert/Edit DBFunc/DBModif' insertBeforeMso='Cut'>" +
               "<button id='DBMapperL' tag='DBMapper' label='DBMapper' imageMso='TableSave' onAction='clickCreateButton'/>" +
               "<button id='DBSequenceL' tag='DBSeqnce' label='DBSequence' imageMso='ShowOnNewButton' onAction='clickCreateButton'/>" +
             "</menu>" +
             "<menuSeparator id='MySeparatorL' insertBeforeMso='Cut'/>" +
         "</contextMenu>" +
         "</contextMenus></customUI>"
        Return customUIXml
    End Function

    ''' <summary>toggle designmode button (actually dialogBox launcher on DBModif Menu)</summary>
    ''' <param name="control"></param>
    Sub toggleDesignMode(control As IRibbonControl)

        ' Ctrl-Shift starts the CustomXML display
        If My.Computer.Keyboard.CtrlKeyDown And My.Computer.Keyboard.ShiftKeyDown Then
            Dim theEditDBModifDefDlg As EditDBModifDef = New EditDBModifDef()
            If theEditDBModifDefDlg.ShowDialog() = System.Windows.Forms.DialogResult.OK Then DBModifs.getDBModifDefinitions()
            Exit Sub
        End If

        Dim cbrs As Object = ExcelDnaUtil.Application.CommandBars
        If Not IsNothing(cbrs) AndAlso cbrs.GetEnabledMso("DesignMode") Then
            cbrs.ExecuteMso("DesignMode")
        Else
            ' this should actually never be reached...
            MsgBox("Couldn't toggle designmode, because Designmode commandbar button is not available !")
        End If
        ' update state of designmode in screentip
        theRibbon.InvalidateControl(control.Id)
    End Sub

    ''' <summary>display state of designmode in screentip of dialogBox launcher</summary>
    ''' <param name="control"></param>
    ''' <returns>screentip and the state of designmode</returns>
    Public Function getToggleDesignScreentip(control As IRibbonControl) As String
        Dim cbrs As Object = ExcelDnaUtil.Application.CommandBars
        If Not IsNothing(cbrs) AndAlso cbrs.GetEnabledMso("DesignMode") Then
            Return "Designmode is currently " & IIf(cbrs.GetPressedMso("DesignMode"), "on !", "off !") & "; Ctrl-Shift-click to inspect/edit DBModifier definitions here"
        Else
            ' this should actually never be reached...
            Return "Designmode commandbar button not available; Ctrl-Shift-click to inspect/edit DBModifier definitions here"
        End If
    End Function

    ''' <summary>for environment dropdown to get the total number of the entries</summary>
    ''' <returns></returns>
    Public Function GetItemCount(control As IRibbonControl) As Integer
        Return Globals.environdefs.Length
    End Function

    ''' <summary>for environment dropdown to get the label of the entries</summary>
    ''' <returns></returns>
    Public Function GetItemLabel(control As IRibbonControl, index As Integer) As String
        Return Globals.environdefs(index)
    End Function

    ''' <summary>for environment dropdown to get the ID of the entries</summary>
    ''' <returns></returns>
    Public Function GetItemID(control As IRibbonControl, index As Integer) As String
        Return Globals.environdefs(index)
    End Function

    ''' <summary>after selection of environment (using selectItem) used to return the selected environment</summary>
    ''' <returns></returns>
    Public Function GetSelItem(control As IRibbonControl) As Integer
        Return Globals.selectedEnvironment
    End Function

    ''' <summary>Choose environment (configured in registry with ConstConnString(N), ConfigStoreFolder(N))</summary>
    Public Sub selectItem(control As IRibbonControl, id As String, index As Integer)
        Globals.selectedEnvironment = index
        Dim env As String = (index + 1).ToString()

        If GetSetting("DBAddin", "Settings", "DontChangeEnvironment", String.Empty) = "Y" Then
            MsgBox("Setting DontChangeEnvironment is set to Y, therefore changing the Environment is prevented !")
            Exit Sub
        End If
        storeSetting("ConstConnString", fetchSetting("ConstConnString" & env, String.Empty))
        storeSetting("ConfigStoreFolder", fetchSetting("ConfigStoreFolder" & env, String.Empty))
        storeSetting("ConfigName", fetchSetting("ConfigName" & env, String.Empty))
        initSettings()
        Dim retval As MsgBoxResult = MsgBox("ConstConnString" & ConstConnString & vbCrLf & "ConfigStoreFolder:" & ConfigStoreFolder & vbCrLf & vbCrLf & "Refresh DBFunctions to see effects?", vbYesNo, "Changed environment to: " & fetchSetting("ConfigName" & env, String.Empty))
        ' provide a chance to reconnect when switching environment...
        conn = Nothing
        If retval = vbYes Then Globals.refreshData()
    End Sub

    ''' <summary>dialogBoxLauncher of leftmost group: activate about box</summary>
    Public Sub showAbout(control As IRibbonControl)
        Dim myAbout As AboutBox = New AboutBox
        myAbout.ShowDialog()
        ' if disabling the addin was chosen, then suicide here..
        If myAbout.disableAddinAfterwards Then
            Try : ExcelDnaUtil.Application.AddIns("DBaddin").Installed = False : Catch ex As Exception : End Try
        End If
    End Sub

    ''' <summary>on demand, refresh the DB Config tree</summary>
    Public Sub refreshDBConfigTree(control As IRibbonControl)
        initSettings()
        createConfigTreeMenu()
        MsgBox("refreshed DB Config Tree Menu", vbInformation + vbOKOnly, "DBAddin: refresh Config tree...")
        theRibbon.Invalidate()
    End Sub

    ''' <summary>get DB Config Menu from File</summary>
    ''' <returns></returns>
    Public Function getDBConfigMenu(control As IRibbonControl) As String
        If ConfigMenuXML = vbNullString Then createConfigTreeMenu()
        Return ConfigMenuXML
    End Function

    ''' <summary>load config if config tree menu end-button has been activated (path to config xcl file is in control.Tag)</summary>
    Public Sub getConfig(control As IRibbonControl)
        loadConfig(control.Tag)
    End Sub

    ''' <summary>set the name of the DBModifType dropdown to the sheet name (for the WB dropdown this is the WB name)</summary>
    ''' <returns></returns>
    Public Function getDBModifTypeLabel(control As IRibbonControl) As String
        getDBModifTypeLabel = IIf(control.Id = "DBSeqnce", "DBSequences", control.Id)
    End Function

    ''' <summary>create the buttons in the DBModif sheet dropdown menu</summary>
    ''' <returns></returns>
    Public Function getDBModifMenuContent(control As IRibbonControl) As String
        Dim xmlString As String = "<menu xmlns='http://schemas.microsoft.com/office/2009/07/customui'>"
        Try
            If Not Globals.DBModifDefColl.ContainsKey(control.Id) Then Return ""
            Dim DBModifTypeName As String = IIf(control.Id = "DBSeqnce", "DBSequences", IIf(control.Id = "DBMapper", "DB Mapper", IIf(control.Id = "DBAction", "DB Action", "undefined DBModifTypeName")))
            For Each nodeName As String In Globals.DBModifDefColl(control.Id).Keys
                Dim descName As String = IIf(nodeName = control.Id, "Unnamed " + DBModifTypeName, Replace(nodeName, DBModifTypeName, ""))
                Dim imageMsoStr As String = IIf(control.Id = "DBSeqnce", "ShowOnNewButton", IIf(control.Id = "DBMapper", "TableSave", IIf(control.Id = "DBAction", "TableIndexes", "undefined imageMso")))
                Dim superTipStr As String = IIf(control.Id = "DBSeqnce", "executes " + DBModifTypeName + " defined in docproperty: " + nodeName, IIf(control.Id = "DBMapper", "stores data defined in DBMapper (named " + nodeName + ") range on " + Globals.DBModifDefColl(control.Id).Item(nodeName).getTargetRangeAddress(), IIf(control.Id = "DBAction", "executes Action defined in DBAction (named " + nodeName + ") range on " + Globals.DBModifDefColl(control.Id).Item(nodeName).getTargetRangeAddress(), "undefined superTip")))
                xmlString = xmlString + "<button id='_" + nodeName + "' label='do " + descName + "' imageMso='" & imageMsoStr & "' onAction='DBModifClick' tag='" + control.Id + "' screentip='do " & DBModifTypeName & ": " + descName + "' supertip='" + superTipStr + "' />"
            Next
            xmlString += "</menu>"
            Return xmlString
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function

    ''' <summary>show a screentip for the dynamic DBMapper/DBAction/DBSequence Menus (also showing the ID behind)</summary>
    ''' <returns></returns>
    Public Function getDBModifScreentip(control As IRibbonControl) As String
        Return "Select DBModifier to store/do action/do sequence (" & control.Id & ")"
    End Function

    ''' <summary>shows the DBModif sheet button only if it was collected...</summary>
    ''' <returns></returns>
    Public Function getDBModifMenuVisible(control As IRibbonControl) As Boolean
        Try
            Return Globals.DBModifDefColl.ContainsKey(control.Id)
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>DBModif button activated, do DB Mapper/DB Action/DB Sequence or define existing (CtrlKey pressed)...</summary>
    Public Sub DBModifClick(control As IRibbonControl)
        Dim nodeName As String = Right(control.Id, Len(control.Id) - 1)
        If Not ExcelDnaUtil.Application.CommandBars.GetEnabledMso("FileNewDefault") Then
            MsgBox("Cannot execute DB Modifier while cell editing active !", MsgBoxStyle.Critical)
            Exit Sub
        End If
        Try
            If My.Computer.Keyboard.CtrlKeyDown And My.Computer.Keyboard.ShiftKeyDown Then
                createDBModif(control.Tag, targetDefName:=nodeName)
            Else
                ' DB sequence actions (the sequence to be done) are stored directly in DBMapperDefColl, so different invocation here
                Globals.DBModifDefColl(control.Tag).Item(nodeName).doDBModif()
            End If
        Catch ex As Exception
            LogError(ex.Message & ",control.Tag:" & control.Tag & ",nodeName:" & nodeName)
        End Try
    End Sub

    ''' <summary>context menu entry refreshData: refresh Data in db function (if area or cell selected) or all db functions</summary>
    Public Sub deleteRowButton(control As IRibbonControl)
        deleteRow()
    End Sub

    ''' <summary>context menu entry refreshData: refresh Data in db function (if area or cell selected) or all db functions</summary>
    Public Sub clickrefreshData(control As IRibbonControl)
        refreshData()
    End Sub

    ''' <summary>context menu entry gotoDBFunc: jumps from DB function to data area and back</summary>
    Public Sub clickjumpButton(control As IRibbonControl)
        jumpButton()
    End Sub

    ''' <summary>purge name tool button, purge names used for dbfunctions from workbook</summary>
    Public Sub clickpurgetoolbutton(control As IRibbonControl)
        purgeNames()
    End Sub

    ''' <summary>show the trace log</summary>
    Public Sub clickShowLog(control As IRibbonControl)
        ExcelDna.Logging.LogDisplay.Show()
    End Sub

    ''' <summary>context menu entries below create...: create DB function or DB Modification definition</summary>
    Public Sub clickCreateButton(control As IRibbonControl)
        ' check for existing DBMapper or DBAction definition and allow exit
        Dim activeCellDBModifName As String = getDBModifNameFromRange(ExcelDnaUtil.Application.ActiveCell)
        Dim activeCellDBModifType As String = Left(activeCellDBModifName, 8)
        If (activeCellDBModifType = "DBMapper" Or activeCellDBModifType = "DBAction") And activeCellDBModifType <> control.Tag And control.Tag <> "DBSeqnce" Then
            Dim exitMe As Boolean = True
            MsgBox("Active Cell already contains definition for a " & activeCellDBModifType & ", inserting " & IIf(control.Tag = "DBSetQueryPivot" Or control.Tag = "DBSetQueryListObject", "DBSetQuery", control.Tag) & " here will cause trouble !", vbCritical, "Inserting not allowed")
            Exit Sub
        End If
        If control.Tag = "DBListFetch" Then
            createFunctionsInCells(ExcelDnaUtil.Application.ActiveCell, {"RC", "=DBListFetch("""","""",R[1]C,,,TRUE,TRUE,TRUE)"})
        ElseIf control.Tag = "DBRowFetch" Then
            createFunctionsInCells(ExcelDnaUtil.Application.ActiveCell, {"RC", "=DBRowFetch("""","""",TRUE,R[1]C:R[1]C[10])"})
        ElseIf control.Tag = "DBSetQueryPivot" Then
            ' insert dummy cmd sql definition for pivot table
            Dim pivotcache As Excel.PivotCache = ExcelDnaUtil.Application.ActiveWorkbook.PivotCaches().Add(Excel.XlPivotTableSourceType.xlExternal)
            pivotcache.Connection = "OLEDB;" & Globals.ConstConnString
            pivotcache.MaintainConnection = False
            pivotcache.CommandText = "select CURRENT_TIMESTAMP" ' this should be sufficient for most databases
            pivotcache.CommandType = Excel.XlCmdType.xlCmdSql
            Dim pivotTables As Excel.PivotTables = ExcelDnaUtil.Application.ActiveSheet.PivotTables()
            Try
                pivotTables.Add(pivotcache, ExcelDnaUtil.Application.ActiveCell.Offset(1, 0), "PivotTable1")
            Catch ex As Exception
                LogWarn("Exception caught when adding pivot table:" & ex.Message)
                Exit Sub
            End Try
            createFunctionsInCells(ExcelDnaUtil.Application.ActiveCell, {"RC", "=DBSetQuery("""","""",R[1]C)"})
        ElseIf control.Tag = "DBSetQueryListObject" Then
            Try
                ' insert dummy cmd sql definition for listobject table
                With ExcelDnaUtil.Application.ActiveSheet.ListObjects.Add(SourceType:=Excel.XlListObjectSourceType.xlSrcQuery, Source:="OLEDB;" & Globals.ConstConnString, Destination:=ExcelDnaUtil.Application.ActiveCell.Offset(0, 1)).QueryTable
                    .CommandType = Excel.XlCmdType.xlCmdSql
                    .CommandText = "select CURRENT_TIMESTAMP" ' this should be sufficient for all ansi sql compliant databases
                    .RowNumbers = False
                    .FillAdjacentFormulas = False
                    .PreserveFormatting = True
                    .BackgroundQuery = True
                    .RefreshStyle = Excel.XlCellInsertionMode.xlInsertDeleteCells
                    .SavePassword = False
                    .SaveData = True
                    .AdjustColumnWidth = True
                    .RefreshPeriod = 0
                    .PreserveColumnInfo = True
                    .Refresh(BackgroundQuery:=False)
                End With
            Catch ex As Exception
                LogWarn("Exception caught when adding listobject table:" & ex.Message)
                Exit Sub
            End Try
            createFunctionsInCells(ExcelDnaUtil.Application.ActiveCell, {"RC", "=DBSetQuery("""","""",RC[1])"})
        ElseIf control.Tag = "DBMapper" Or control.Tag = "DBAction" Or control.Tag = "DBSeqnce" Then
            If activeCellDBModifType = control.Tag Then                 ' edit existing definition
                createDBModif(control.Tag, targetDefName:=activeCellDBModifName)
            Else                                                     ' create new definition
                createDBModif(control.Tag)
            End If
        End If
    End Sub

End Class
