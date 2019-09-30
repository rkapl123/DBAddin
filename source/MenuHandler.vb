Imports ExcelDna.Integration.CustomUI
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports System.Linq

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
        ' DBModif Group: max. 15 sheets with DBModif definitions possible: 
        customUIXml += "<group id='DBModifGroup' label='Store DBModif Data'>"
        For i As Integer = 0 To 14
            customUIXml += "<dynamicMenu id='ID" + i.ToString() + "' " +
                                            "size='large' getLabel='getSheetLabel' imageMso='ApplicationOptionsDialog' " +
                                            "getScreentip='getDBModifScreentip' getContent='getDBModifMenuContent' getVisible='getDBModifMenuVisible'/>"
        Next
        ' Context menus for refresh, jump and creation: in cell, row, column and ListRange (area of ListObjects)
        customUIXml += "</group></tab></tabs></ribbon>" +
         "<contextMenus>" +
         "<contextMenu idMso ='ContextMenuCell'>" +
         "<menu id='createMenu' label='Insert/Edit DBFunc/DBModif' insertBeforeMso='Cut'>" +
           "<button id='DBMapper' tag='DBMapper' label='DBMapper' imageMso='TableSave' onAction='clickCreateButton'/>" +
           "<button id='DBAction' tag='DBAction' label='DBAction' imageMso='TableIndexes' onAction='clickCreateButton'/>" +
           "<button id='DBSequence' tag='DBSeqnce' label='DBSequence' imageMso='ShowOnNewButton' onAction='clickCreateButton'/>" +
           "<button id='DBListFetch' tag='DBListFetch' label='DBListFetch' imageMso='GroupLists' onAction='clickCreateButton'/>" +
           "<button id='DBRowFetch' tag='DBRowFetch' label='DBRowFetch' imageMso='GroupRecords' onAction='clickCreateButton'/>" +
           "<button id='DBSetQueryPivot' tag='DBSetQueryPivot' label='DBSetQueryPivot' imageMso='AddContentType' onAction='clickCreateButton'/>" +
           "<button id='DBSetQueryListObject' tag='DBSetQueryListObject' label='DBSetQueryListObject' imageMso='AddContentType' onAction='clickCreateButton'/>" +
         "</menu>" +
         "<button id='refreshDataC' label='refresh data (Ctrl-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
         "<button id='gotoDBFuncC' label='jump to DBFunc/target (Ctrl-J)' imageMso='ConvertTextToTable' onAction='clickjumpButton' insertBeforeMso='Cut'/>" +
         "<menuSeparator id='MySeparatorC' insertBeforeMso='Cut'/>" +
         "</contextMenu>" +
         "<contextMenu idMso ='ContextMenuCellLayout'>" +
         "<button id='refreshDataCL' label='refresh data (Ctrl-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
         "<button id='gotoDBFuncCL' label='jump to DBFunc/target (Ctrl-J)' imageMso='ConvertTextToTable' onAction='clickjumpButton' insertBeforeMso='Cut'/>" +
         "<menu id='createMenuCL' label='Insert/Edit DBFunc/DBModif' insertBeforeMso='Cut'>" +
            "<button id='DBMapperCL' tag='DBMapper' label='DBMapper' imageMso='TableSave' onAction='clickCreateButton'/>" +
            "<button id='DBActionCL' tag='DBAction' label='DBAction' imageMso='TableIndexes' onAction='clickCreateButton'/>" +
            "<button id='DBSequenceCL' tag='DBSeqnce' label='DBSequence' imageMso='ShowOnNewButton' onAction='clickCreateButton'/>" +
            "<button id='DBListFetchCL' tag='DBListFetch' label='DBListFetch' imageMso='GroupLists' onAction='clickCreateButton'/>" +
            "<button id='DBRowFetchCL' tag='DBRowFetch' label='DBRowFetch' imageMso='GroupRecords' onAction='clickCreateButton'/>" +
            "<button id='DBSetQueryPivotCL' tag='DBSetQueryPivot' label='DBSetQueryPivot' imageMso='AddContentType' onAction='clickCreateButton'/>" +
            "<button id='DBSetQueryListObjectCL' tag='DBSetQueryListObject' label='DBSetQueryListObject' imageMso='AddContentType' onAction='clickCreateButton'/>" +
         "</menu>" +
         "<menuSeparator id='MySeparatorCL' insertBeforeMso='Cut'/>" +
         "</contextMenu>" +
         "<contextMenu idMso='ContextMenuRow'>" +
         "<button id='refreshDataR' label='refresh data (Ctrl-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
         "<menuSeparator id='MySeparatorR' insertBeforeMso='Cut'/>" +
         "</contextMenu>" +
         "<contextMenu idMso='ContextMenuColumn'>" +
         "<button id='refreshDataZ' label='refresh data (Ctrl-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
         "<menuSeparator id='MySeparatorZ' insertBeforeMso='Cut'/>" +
         "</contextMenu>" +
         "<contextMenu idMso='ContextMenuListRange'>" +
             "<button id='refreshDataL' label='refresh data (Ctrl-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
             "<button id='gotoDBFuncL' label='jump to DBFunc/target (Ctrl-J)' imageMso='ConvertTextToTable' onAction='clickjumpButton' insertBeforeMso='Cut'/>" +
             "<menu id='createMenuL' label='Insert/Edit DBFunc/DBModif' insertBeforeMso='Cut'>" +
               "<button id='DBMapperL' tag='DBMapper' label='DBMapper' imageMso='TableSave' onAction='clickCreateButton'/>" +
               "<button id='DBSequenceL' tag='DBSeqnce' label='DBSequence' imageMso='ShowOnNewButton' onAction='clickCreateButton'/>" +
             "</menu>" +
             "<menuSeparator id='MySeparatorL' insertBeforeMso='Cut'/>" +
         "</contextMenu>" +
         "</contextMenus></customUI>"
        Return customUIXml
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
        MsgBox("ConstConnString" & ConstConnString & vbCrLf & "ConfigStoreFolder:" & ConfigStoreFolder & vbCrLf & vbCrLf & "Please refresh DBFunctions to see effects...", vbOKOnly, "set defaults to: ")
        ' provide a chance to reconnect when switching environment...
        conn = Nothing
    End Sub

    ''' <summary>dialogBoxLauncher of leftmost group: activate about box</summary>
    Public Sub showAbout(control As IRibbonControl)
        Dim myAbout As AboutBox = New AboutBox
        myAbout.ShowDialog()
        ' if disabling the addin was chosen, then suicide here...
        If myAbout.disableAddinAfterwards Then
            Try : hostApp.AddIns("DBAddin-AddIn-packed").Installed = False : Catch ex As Exception : End Try
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

    ''' <summary>set the name of the WB/sheet dropdown to the sheet name (for the WB dropdown this is the WB name)</summary>
    ''' <returns></returns>
    Public Function getSheetLabel(control As IRibbonControl) As String
        getSheetLabel = IIf(getSheetNameForMenuID(control.Id) = "ID0", "DBSequences", getSheetNameForMenuID(control.Id))
    End Function

    ''' <summary>create the buttons in the DBModif sheet dropdown menu</summary>
    ''' <returns></returns>
    Public Function getDBModifMenuContent(control As IRibbonControl) As String
        Dim xmlString As String = "<menu xmlns='http://schemas.microsoft.com/office/2009/07/customui'>"
        Try
            If control.Id = "ID0" Then
                ' special menu for sequences
                For Each nodeName As String In DBModifDefColl(control.Id).Keys
                    Dim descName As String = IIf(nodeName = "", "Unnamed DBSequence", nodeName)
                    xmlString = xmlString + "<button id='_" + nodeName + "' label='do " + descName + "' imageMso='ShowOnNewButton' onAction='DBSeqnceClick' tag='" + control.Id + "' screentip='do Sequence: " + descName + "' supertip='executes DB Sequence defined in docproperty DBSeqnce: " + descName + "' />"
                Next
            Else
                Dim curWsName As String = getSheetNameForMenuID(control.Id)
                If Not DBModifDefColl.ContainsKey(curWsName) Then Return ""

                For Each nodeName As String In DBModifDefColl(curWsName).Keys
                    Dim rngName As String = getDBModifNameFromRange(DBModifDefColl(curWsName).Item(nodeName))
                    Dim descName As String = IIf(nodeName = "", "Unnamed " + Left(rngName, 8), nodeName)
                    If Left(rngName, 8) = "DBMapper" Then
                        xmlString = xmlString + "<button id='_" + nodeName + "' label='store " + descName + "' imageMso='TableSave' onAction='DBMapperClick' tag='" + curWsName + "' screentip='store DBMapper: " + descName + "' supertip='stores data defined in DBMapper (named " + descName + ") range on " + DBModifDefColl(curWsName).Item(nodeName).Parent.Name + "!" + DBModifDefColl(curWsName).Item(nodeName).Address + "' />"
                    ElseIf Left(rngName, 8) = "DBAction" Then
                        xmlString = xmlString + "<button id='_" + nodeName + "' label='do " + descName + "' imageMso='TableIndexes' onAction='DBActionClick' tag='" + curWsName + "' screentip='do DBAction: " + descName + "' supertip='executes Action defined in DBAction (named " + descName + ") range on " + DBModifDefColl(curWsName).Item(nodeName).Parent.Name + "!" + DBModifDefColl(curWsName).Item(nodeName).Address + "' />"
                    End If
                Next
            End If
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
            Return DBModifDefColl.ContainsKey(getSheetNameForMenuID(control.Id))
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>get the worksheet name for a given control ID (via the index of the sheet)</summary>
    ''' <param name="controlId"></param>
    ''' <returns></returns>
    Private Function getSheetNameForMenuID(controlId As String) As String
        If controlId = "ID0" Then Return controlId
        Dim curIndex As Integer = CInt(Replace(controlId, "ID", ""))
        Dim curWsName As String = ""
        For Each ws In hostApp.ActiveWorkbook.Worksheets
            If ws.Index = curIndex Then Return ws.Name
        Next
        Return ""
    End Function

    ''' <summary>DBMapper store button activated, save Range to DB or define existing (CtrlKey pressed)...</summary>
    Public Sub DBMapperClick(control As IRibbonControl)
        Dim nodeName As String = Right(control.Id, Len(control.Id) - 1) ' remove underscore at beginning of id
        If My.Computer.Keyboard.CtrlKeyDown And My.Computer.Keyboard.ShiftKeyDown Then
            createDBModif("DBMapper", targetRange:=DBModifDefColl(control.Tag).Item(nodeName))
        Else
            doDBMapper(DBModifDefColl(control.Tag).Item(nodeName))
        End If
    End Sub

    ''' <summary>DBAction button activated, do DB Action or define existing (CtrlKey pressed)...</summary>
    Public Sub DBActionClick(control As IRibbonControl)
        Dim nodeName As String = Right(control.Id, Len(control.Id) - 1)
        If My.Computer.Keyboard.CtrlKeyDown And My.Computer.Keyboard.ShiftKeyDown Then
            createDBModif("DBAction", targetRange:=DBModifDefColl(control.Tag).Item(nodeName))
        Else
            doDBAction(DBModifDefColl(control.Tag).Item(nodeName))
        End If
    End Sub

    ''' <summary>DBSequence button activated, do DB Sequence or define existing (CtrlKey pressed)...</summary>
    Public Sub DBSeqnceClick(control As IRibbonControl)
        Dim nodeName As String = Right(control.Id, Len(control.Id) - 1)
        If My.Computer.Keyboard.CtrlKeyDown And My.Computer.Keyboard.ShiftKeyDown Then
            createDBModif("DBSeqnce", targetDefName:=nodeName, DBSequenceText:=DBModifDefColl(control.Tag).Item(nodeName))
        Else
            ' DB sequence actions (the sequence to be done) are stored directly in DBMapperDefColl, so different invocation here
            doDBSeqnce(nodeName, DBModifDefColl(control.Tag).Item(nodeName))
        End If
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
        Dim currentDBModifName As String = Left(getDBModifNameFromRange(hostApp.ActiveCell), 8)
        If (currentDBModifName = "DBMapper" Or currentDBModifName = "DBAction") And currentDBModifName <> control.Tag And control.Tag <> "DBSeqnce" Then
            Dim exitMe As Boolean = True
            ErrorMsg("Active Cell already contains definition for a " & currentDBModifName & ", inserting " & IIf(control.Tag = "DBSetQueryPivot" Or control.Tag = "DBSetQueryListObject", "DBSetQuery", control.Tag) & " here will probably cause trouble !", exitMe, "")
            If exitMe Then Exit Sub
        End If
        If control.Tag = "DBListFetch" Then
            createFunctionsInCells(hostApp.ActiveCell, {"RC", "=DBListFetch("""","""",R[1]C,,,TRUE,TRUE,TRUE)"})
        ElseIf control.Tag = "DBRowFetch" Then
            createFunctionsInCells(hostApp.ActiveCell, {"RC", "=DBRowFetch("""","""",TRUE,R[1]C:R[1]C[10])"})
        ElseIf control.Tag = "DBSetQueryPivot" Then
            Dim pivotcache As Excel.PivotCache = hostApp.ActiveWorkbook.PivotCaches().Add(Excel.XlPivotTableSourceType.xlExternal)
            pivotcache.Connection = "OLEDB;" & Globals.ConstConnString
            pivotcache.MaintainConnection = False
            pivotcache.CommandText = "select CURRENT_TIMESTAMP" ' this should be sufficient for most databases
            pivotcache.CommandType = Excel.XlCmdType.xlCmdSql
            Dim pivotTables As Excel.PivotTables = hostApp.ActiveSheet.PivotTables()
            Try
                pivotTables.Add(pivotcache, hostApp.ActiveCell.Offset(1, 0), "PivotTable1")
            Catch ex As Exception
                LogWarn("Exception caught when adding pivot table:" & ex.Message)
                Exit Sub
            End Try
            createFunctionsInCells(hostApp.ActiveCell, {"RC", "=DBSetQuery("""","""",R[1]C)"})
        ElseIf control.Tag = "DBSetQueryListObject" Then
            Try
                With hostApp.ActiveSheet.ListObjects.Add(SourceType:=Excel.XlListObjectSourceType.xlSrcQuery, Source:="OLEDB;" & Globals.ConstConnString, Destination:=hostApp.ActiveCell.Offset(0, 1)).QueryTable
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
            createFunctionsInCells(hostApp.ActiveCell, {"RC", "=DBSetQuery("""","""",RC[1])"})
        ElseIf control.Tag = "DBMapper" Or control.Tag = "DBAction" Or control.Tag = "DBSeqnce" Then
            If currentDBModifName = control.Tag Then                 ' edit existing definition
                createDBModif(control.Tag, hostApp.ActiveCell)
            Else                                                     ' create new definition
                createDBModif(control.Tag)
            End If
        End If
    End Sub

End Class
