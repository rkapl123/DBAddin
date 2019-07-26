Imports ExcelDna.Integration
Imports ExcelDna.Integration.CustomUI
Imports System.Runtime.InteropServices

''' <summary>handles all Menu related aspects (context menu for building/refreshing, "DBAddin"/"Load Config" tree menu for retrieving stored configuration files</summary>
<ComVisible(True)>
Public Class MenuHandler
    Inherits ExcelRibbon

    ''' <summary>callback after Excel loaded the Ribbon, used to initialize data for the Ribbon</summary>
    Public Sub ribbonLoaded(theRibbon As ExcelDna.Integration.CustomUI.IRibbonUI)
        DBAddin.theRibbon = theRibbon
    End Sub

    ''' <summary>creates the Ribbon (only at startup). any changes to the ribbon can only be done via dynamic menus</summary>
    Public Overrides Function GetCustomUI(RibbonID As String) As String
        Dim customUIXml As String = "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='ribbonLoaded' ><ribbon><tabs><tab id='DBaddinTab' label='DB Addin'>" +
            "<group id='DBAddinGroup' label='General settings'>" +
              "<dropDown id='envDropDown' label='Environment:' sizeString='12345678901234567890' getSelectedItemIndex='GetSelItem' getItemCount='GetItemCount' getItemID='GetItemID' getItemLabel='GetItemLabel' onAction='selectItem'/>" +
              "<dynamicMenu id='DBConfigs' size='normal' label='DB Configs' imageMso='Refresh' screentip='DB Function Configuration Files quick access' getContent='getDBConfigMenu'/>" +
              "<dialogBoxLauncher><button id='dialog' label='About DBAddin' onAction='showAbout' tag='3' screentip='Show Aboutbox with help, version information and homepage'/></dialogBoxLauncher></group>" +
              "<group id='DBMapperGroup' label='Store DBMapper Data'>"
        ' max. 15 sheets with DBMapper definitions possible:
        For i As Integer = 0 To 14
            customUIXml = customUIXml + "<dynamicMenu id='ID" + i.ToString() + "' " +
                                            "size='large' getLabel='getSheetLabel' imageMso='SignatureLineInsert' " +
                                            "screentip='Select DBMapper range to store' " +
                                            "getContent='getDBMapperMenuContent' getVisible='getDBMapperMenuVisible'/>"
        Next
        ' context menus for refresh, jump and creation: in cell, row, column and ListRange (in a ListObject area)
        customUIXml = customUIXml + "</group></tab></tabs></ribbon>" +
         "<contextMenus><contextMenu idMso='ContextMenuCell'>" +
         "<button id='refreshDataC' label='refresh data (Ctrl-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
         "<button id='gotoDBFuncC' label='jump to DBFunc/target (Ctrl-J)' imageMso='ConvertTextToTable' onAction='clickjumpButton' insertBeforeMso='Cut'/>" +
         "<menu id='createMenuC' label='DBAddin create ...' insertBeforeMso='Cut'>" +
            "<button id='DBMapper' tag='DBMapper' label='DBMapper' imageMso='AddToolGallery' onAction='clickCreateButton'/>" +
            "<button id='DBListFetch' tag='DBListFetch' label='DBListFetch' imageMso='AddCalendarMenu' onAction='clickCreateButton'/>" +
            "<button id='DBRowFetch' tag='DBRowFetch' label='DBRowFetch' imageMso='DataFormAddRecord' onAction='clickCreateButton'/>" +
            "<button id='DBSetQuery' tag='DBSetQuery' label='DBSetQuery' imageMso='AddContentType' onAction='clickCreateButton'/></menu>" +
         "<menuSeparator id='MySeparatorC' insertBeforeMso='Cut'/></contextMenu>" +
         "<contextMenu idMso='ContextMenuRow'>" +
         "<button id='refreshDataR' label='refresh data (Ctrl-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
         "<menuSeparator id='MySeparatorR' insertBeforeMso='Cut'/></contextMenu>" +
         "<contextMenu idMso='ContextMenuColumn'>" +
         "<button id='refreshDataZ' label='refresh data (Ctrl-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
         "<menuSeparator id='MySeparatorZ' insertBeforeMso='Cut'/></contextMenu>" +
         "<contextMenu idMso='ContextMenuListRange'>" +
         "<button id='refreshDataL' label='refresh data (Ctrl-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
         "<button id='gotoDBFuncL' label='jump to DBFunc/target (Ctrl-J)' imageMso='ConvertTextToTable' onAction='clickjumpButton' insertBeforeMso='Cut'/>" +
         "<menu id='createMenuL' label='DBAddin create ...' insertBeforeMso='Cut'>" +
            "<button id='DBMapperC' tag='DBMapper' label='DBMapper' imageMso='AddToolGallery' onAction='clickCreateButton'/>" +
            "<button id='DBListFetchC' tag='DBListFetch' label='DBListFetch' imageMso='AddCalendarMenu' onAction='clickCreateButton'/>" +
            "<button id='DBRowFetchC' tag='DBRowFetch' label='DBRowFetch' imageMso='DataFormAddRecord' onAction='clickCreateButton'/>" +
            "<button id='DBSetQueryC' tag='DBSetQuery' label='DBSetQuery' imageMso='AddContentType' onAction='clickCreateButton'/></menu>" +
         "<menuSeparator id='MySeparatorL' insertBeforeMso='Cut'/></contextMenu>" +
         "</contextMenus></customUI>"
        Return customUIXml
    End Function

    ''' <summary>for environment dropdown to get the total number of the entries</summary>
    Public Function GetItemCount(control As IRibbonControl) As Integer
        Return environdefs.Length
    End Function

    ''' <summary>for environment dropdown to get the label of the entries</summary>
    Public Function GetItemLabel(control As IRibbonControl, index As Integer) As String
        Return environdefs(index)
    End Function

    ''' <summary>for environment dropdown to get the ID of the entries</summary>
    Public Function GetItemID(control As IRibbonControl, index As Integer) As String
        Return environdefs(index)
    End Function

    ''' <summary>after selection of environment (using selectItem) used to return the selected environment</summary>
    Public Function GetSelItem(control As IRibbonControl) As Integer
        Return selectedEnvironment
    End Function

    ''' <summary>Choose environment (configured in registry with ConstConnString(N), ConfigStoreFolder(N))</summary>
    Public Sub selectItem(control As IRibbonControl, id As String, index As Integer)
        selectedEnvironment = index
        Dim env As String = (index + 1).ToString()

        If GetSetting("DBAddin", "Settings", "DontChangeEnvironment", String.Empty) = "Y" Then
            MsgBox("Setting DontChangeEnvironment Is set to Y, therefore changing the Environment Is prevented !")
            Exit Sub
        End If
        storeSetting("ConstConnString", fetchSetting("ConstConnString" & env, String.Empty))
        storeSetting("ConfigStoreFolder", fetchSetting("ConfigStoreFolder" & env, String.Empty))
        storeSetting("ConfigName", fetchSetting("ConfigName" & env, String.Empty))
        initSettings()
        MsgBox("ConstConnString" & ConstConnString & vbCrLf & "ConfigStoreFolder:" & ConfigStoreFolder & vbCrLf & vbCrLf & "Please refresh DBSheets Or DBFuncs to see effects...", vbOKOnly, "set defaults to: ")
        ' provide a chance to reconnect when switching environment...
        theDBFuncEventHandler.cnn = Nothing
        dontTryConnection = False
    End Sub

    ''' <summary>dialogBoxLauncher of leftmost group: activate about box</summary>
    Public Sub showAbout(control As IRibbonControl)
        Dim myAbout As AboutBox1 = New AboutBox1
        myAbout.ShowDialog()
    End Sub

    ''' <summary>on demand, refresh the DB Config tree</summary>
    Public Sub refreshDBConfigTree(control As IRibbonControl)
        initSettings()
        createConfigTreeMenu()
        MsgBox("refreshed DB Config Tree Menu", vbInformation + vbOKOnly, "DBAddin: refresh Config tree...")
        theRibbon.Invalidate()
    End Sub

    ''' <summary>get DB Config Menu from File</summary>
    Public Function getDBConfigMenu(control As IRibbonControl) As String
        If ConfigMenuXML = vbNullString Then createConfigTreeMenu()
        Return ConfigMenuXML
    End Function

    ''' <summary>load config if config tree menu end-button has been activated (path to config xcl file is in control.Tag)</summary>
    Public Sub getConfig(control As IRibbonControl)
        loadConfig(control.Tag)
    End Sub

    ''' <summary>set the name of the WB/sheet dropdown to the sheet name (for the WB dropdown this is the WB name)</summary>
    Public Function getSheetLabel(control As IRibbonControl) As String
        getSheetLabel = vbNullString
        If DBMapperDefColl.ContainsKey(control.Id) Then
            ' get parent name of first stored DB Mapper range
            getSheetLabel = DBMapperDefColl(control.Id).Item(DBMapperDefColl(control.Id).Keys.First).Parent.Name
        End If
    End Function

    ''' <summary>create the buttons in the DBMapper sheet dropdown menu</summary>
    Public Function getDBMapperMenuContent(control As IRibbonControl) As String
        Dim xmlString As String = "<menu xmlns='http://schemas.microsoft.com/office/2009/07/customui'>"

        If Not DBMapperDefColl.ContainsKey(control.Id) Then Return ""

        For Each nodeName As String In DBMapperDefColl(control.Id).Keys
            xmlString = xmlString + "<button id='" + nodeName + "' label='store " + nodeName + "' imageMso='SignatureLineInsert' onAction='saveRangeToDBClick' tag='" + control.Id + "' screentip='store " + nodeName + "' supertip='stores data defined in " + nodeName + " Mapper range on " + DBMapperDefColl(control.Id).Item(nodeName).Parent.Name + "![" + DBMapperDefColl(control.Id).Item(nodeName).Address + "]' />"
        Next
        xmlString += "</menu>"
        Return xmlString
    End Function

    ''' <summary>shows the DBMapper sheet button only if it was collected...</summary>
    Public Function getDBMapperMenuVisible(control As IRibbonControl) As Boolean
        Return DBMapperDefColl.ContainsKey(control.Id)
    End Function

    ''' <summary>DBMapper store button activated, save Range to DB...</summary>
    Public Sub saveRangeToDBClick(control As IRibbonControl)
        saveRangeToDB(DBMapperDefColl(control.Tag).Item(control.Id), control.Id)
    End Sub

    ''' <summary>context menu entry refreshData: refresh Data in db function (if area or cell selected) or all db functions</summary>
    Public Sub clickrefreshData(control As IRibbonControl)
        refreshData()
    End Sub

    ''' <summary>context menu entry gotoDBFunc: jumps from DB function to data area and back</summary>
    Public Sub clickjumpButton(control As IRibbonControl)
        jumpButton()
    End Sub

    ''' <summary>context menu entries below create...: create DB function or DB Mapper</summary>
    Public Sub clickCreateButton(control As IRibbonControl)
        If control.Tag = "DBListFetch" Then
            createFunctionsInCells(hostApp.ActiveCell, {"RC", "=DBListFetch("""","""",R[1]C,,,TRUE,TRUE,TRUE)"})
        ElseIf control.Tag = "DBRowFetch" Then
            createFunctionsInCells(hostApp.ActiveCell, {"RC", "=DBRowFetch("""","""",TRUE,R[1]C:R[1]C[10])"})
        ElseIf control.Tag = "DBSetQuery" Then
            createFunctionsInCells(hostApp.ActiveCell, {"RC", "=DBSetQuery("""","""",R[1]C)"})
        ElseIf control.Tag = "DBMapper" Then
            createDBMapper()
        End If
    End Sub

End Class
