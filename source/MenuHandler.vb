Imports ExcelDna.Integration
Imports ExcelDna.Integration.CustomUI
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel
Imports System.IO

''
'  handles all Menu related aspects (context menu for building/refreshing,
'             "DBAddin"/"Load Config" tree menu for retrieving stored configuration files
<ComVisible(True)>
Public Class MenuHandler
    Inherits ExcelRibbon

    Private specialConfigFoldersTempColl As Collection
    Private selectedEnvironment As Integer

    Public Sub ribbonLoaded(theRibbon As ExcelDna.Integration.CustomUI.IRibbonUI)
        DBAddin.theRibbon = theRibbon
        hostApp = ExcelDnaUtil.Application
        defsheetColl = New Dictionary(Of String, Dictionary(Of String, Range))
        defsheetMap = New Dictionary(Of String, String)

        ' load environments
        Dim i As Integer = 1
        Dim ConfigName As String
        Do
            ConfigName = fetchSetting("ConstConnStringName" + i.ToString(), vbNullString)
            If Len(ConfigName) > 0 Then
                ReDim Preserve defnames(defnames.Length)
                defnames(defnames.Length - 1) = ConfigName + " - " + i.ToString()
            End If
            ' set selectedEnvironment
            If fetchSetting("ConstConnString" + i.ToString(), vbNullString) = ConstConnString Then
                selectedEnvironment = i - 1
            End If
            i = i + 1
        Loop Until Len(ConfigName) = 0
    End Sub

    Public Sub showAbout(control As IRibbonControl)
        Dim myAbout As AboutBox1 = New AboutBox1
        myAbout.ShowDialog()
    End Sub

    Public Function GetItemCount(control As IRibbonControl) As Integer
        Return defnames.Length
    End Function

    Public Function GetItemLabel(control As IRibbonControl, index As Integer) As String
        Return defnames(index)
    End Function

    Public Function GetItemID(control As IRibbonControl, index As Integer) As String
        Return defnames(index)
    End Function

    Public Function GetSelItem(control As IRibbonControl) As Integer
        Return selectedEnvironment
    End Function

    ' Choose environment (configured in registry with ConstConnString<N>, ConfigStoreFolder<N>)
    Public Sub selectItem(control As IRibbonControl, id As String, index As Integer)
        selectedEnvironment = index
        Dim env As String = (index + 1).ToString()

        If GetSetting("DBAddin", "Settings", "DontChangeEnvironment", String.Empty) = "Y" Then
            MsgBox("Setting DontChangeEnvironment is set to Y, therefore changing the Environment is prevented !")
            Exit Sub
        End If
        storeSetting("ConstConnString", fetchSetting("ConstConnString" & env, String.Empty))
        storeSetting("ConfigStoreFolder", fetchSetting("ConfigStoreFolder" & env, String.Empty))

        initSettings()
        dontTryConnection = False  ' provide a chance to reconnect when switching environment...
    End Sub

    Public Sub refreshDBConfigTree(control As IRibbonControl)
        initSettings()
        createConfigTreeMenu()
        theRibbon.Invalidate()
        MsgBox("refreshed DB Config Tree Menu", vbInformation + vbOKOnly, "DBAddin: refresh Config tree...")
    End Sub

    ' creates the Ribbon
    Public Overrides Function GetCustomUI(RibbonID As String) As String
        Dim customUIXml As String = "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='ribbonLoaded' ><ribbon><tabs><tab id='DBaddinTab' label='DB Addin'>" +
            "<group id='DBaddinGroup' label='General settings'>" +
              "<dropDown id='envDropDown' label='Environment:' sizeString='12345678901234567890' getSelectedItemIndex='GetSelItem' getItemCount='GetItemCount' getItemID='GetItemID' getItemLabel='GetItemLabel' onAction='selectItem'/>" +
              "<dynamicMenu id='DBConfigs' size='normal' label='DB Configs' imageMso='Refresh' screentip='DB Function Configuration Files quick access' getContent='getDBConfigMenu'/>" +
              "<dialogBoxLauncher><button id='dialog' label='About DBAddin' onAction='showAbout' tag='3' screentip='Show Aboutbox and refresh configs if wanted'/></dialogBoxLauncher></group>" +
              "<group id='RscriptsGroup' label='Store Data defined with saveRangeToDB'>"
        For i As Integer = 0 To 10
            customUIXml = customUIXml + "<dynamicMenu id='ID" + i.ToString() + "' " +
                                            "size='normal' getLabel='getSheetLabel' imageMso='SignatureLineInsert' " +
                                            "screentip='Select script to run' " +
                                            "getContent='getDynMenContent' getVisible='getDynMenVisible'/>"
        Next
        customUIXml = customUIXml + "</group></tab></tabs></ribbon>" +
         "<contextMenus><contextMenu idMso='ContextMenuCell'>" +
         "<button id='refreshData' label='refresh data' imageMso='Refresh' onAction='refreshData' insertBeforeMso='Cut'/>" +
         "<button id='gotoDBFunc' label='GoTo DBFunc/target' imageMso='ConvertTextToTable' onAction='jumpButton' insertBeforeMso='Cut'/>" +
         "<menuSeparator id='MySeparator' insertBeforeMso='Cut'/>" +
         "</contextMenu></contextMenus></customUI>"
        Return customUIXml
    End Function

    Public Function getDBConfigMenu(control As IRibbonControl) As String
        If Not File.Exists(menufilename) Then
            Return "<menu xmlns='http://schemas.microsoft.com/office/2009/07/customui'><button id='refreshConfig' label='refresh DBConfig Tree' imageMso='Refresh' onAction='refreshDBConfigTree'/></menu>"
        End If

        Dim menufile As StreamReader = Nothing
        Try
            menufile = New StreamReader(menufilename, System.Text.Encoding.GetEncoding(1252))
        Catch ex As Exception
            MsgBox("Exception occured When trying To open menufile " + menufilename + ": " + ex.Message)
        End Try
        getDBConfigMenu = menufile.ReadToEnd()
        menufile.Close()
    End Function

    ' set the name of the WB/sheet dropdown to the sheet name (for the WB dropdown this is the WB name) 
    Public Function getSheetLabel(control As IRibbonControl) As String
        getSheetLabel = vbNullString
        If defsheetMap.ContainsKey(control.Id) Then getSheetLabel = defsheetMap(control.Id)
    End Function

    ' create the buttons in the WB/sheet dropdown
    Public Function getDynMenContent(control As IRibbonControl) As String
        Dim xmlString As String = "<menu xmlns='http://schemas.microsoft.com/office/2009/07/customui'>"

        If Not defsheetMap.ContainsKey(control.Id) Then Return ""

        Dim currentSheet As String = defsheetMap(control.Id)
        For Each nodeName As String In defsheetColl(currentSheet).Keys
            xmlString = xmlString + "<button id='" + nodeName + "' label='store " + nodeName + "' imageMso='SignatureLineInsert' onAction='saveRangeToDBClick' tag ='" + currentSheet + "' screentip='store " + nodeName + "' supertip='stores data defined in " + nodeName + " Mapper range on sheet " + currentSheet + "' />"
        Next
        xmlString = xmlString + "</menu>"
        Return xmlString
    End Function

    ' shows the sheet button only if it was collected...
    Public Function getDynMenVisible(control As IRibbonControl) As Boolean
        Return defsheetMap.ContainsKey(control.Id)
    End Function

    ' load config if config tree menu has been activated (name stored in Ctrl.Parameter)
    Public Sub getConfig(control As IRibbonControl)
        loadConfig(control.Tag)
    End Sub

    Public Sub refreshData(control As IRibbonControl)
        doRefreshData()
    End Sub

    Public Sub jumpButton(control As IRibbonControl)
        doJumpButton()
    End Sub

    Public Shared Sub saveRangeToDBClick(control As IRibbonControl)
        If saveRangeToDB(hostApp.ActiveCell) Then MsgBox("hooray!")
    End Sub

    Public Sub saveAllWSheetRangesToDBClick(control As IRibbonControl)

    End Sub

    Public Sub saveAllWBookRangesToDBClick(control As IRibbonControl)

    End Sub
End Class
