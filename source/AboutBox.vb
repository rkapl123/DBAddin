Imports System.Diagnostics
Imports ExcelDna.Integration


''' <summary>About box: used to provide information about version/buildtime and links for local help and project homepage</summary>
Public NotInheritable Class AboutBox
    Public disableAddinAfterwards As Boolean = False
    Private dontChangeEventLevels As Boolean

    ''' <summary>set up Aboutbox</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AboutBox1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim sModuleInfo As String = vbNullString

        ' get module info for buildtime (FileDateTime of xll):
        For Each tModule As ProcessModule In Process.GetCurrentProcess().Modules
            Dim sModule As String = tModule.FileName
            If sModule.ToUpper.Contains("DBADDIN") Then
                sModuleInfo = FileDateTime(sModule).ToString()
                Exit For
            End If
        Next
        ' set the UI elements
        Me.Text = String.Format("About {0}", My.Application.Info.Title)
        Me.LabelProductName.Text = "DB Addin Help"
        Me.LabelVersion.Text = String.Format("Version {0} Buildtime {1}", My.Application.Info.Version.ToString, sModuleInfo)
        Me.LabelCopyright.Text = My.Application.Info.Copyright
        Me.LabelCompanyName.Text = "Information: " + My.Application.Info.CompanyName
        Me.TextBoxDescription.Text = My.Application.Info.Description
        dontChangeEventLevels = True
        Me.EventLevels.SelectedItem = Globals.EventLevelSelected
        dontChangeEventLevels = False
    End Sub

    ''' <summary>Close Aboutbox</summary>
    Private Sub OKButton_Click(ByVal sender As Object, ByVal e As EventArgs) Handles OKButton.Click
        Me.Close()
    End Sub

    ''' <summary>Click on Project homepage: activate hyperlink in browser</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub LabelCompanyName_Click(sender As Object, e As EventArgs) Handles LabelCompanyName.Click
        Try
            Process.Start(My.Application.Info.CompanyName)
        Catch ex As Exception
            LogWarn(ex.Message)
        End Try
    End Sub

    ''' <summary>Click on Local help: activate hyperlink in browser</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub LabelProductName_Click(sender As Object, e As EventArgs) Handles LabelProductName.Click
        Try
            Process.Start(fetchSetting("LocalHelp", ""))
        Catch ex As Exception
            LogWarn(ex.Message)
        End Try
    End Sub

    ''' <summary>select event levels: filter events by selected level (from now on)</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub EventLevels_SelectedValueChanged(sender As Object, e As EventArgs) Handles EventLevels.SelectedValueChanged
        If dontChangeEventLevels Then Exit Sub
        Select Case EventLevels.SelectedItem
            Case "Information"
                Globals.theLogListener.Filter = New EventTypeFilter(SourceLevels.Information)
            Case "Warning"
                Globals.theLogListener.Filter = New EventTypeFilter(SourceLevels.Warning)
            Case "Error"
                Globals.theLogListener.Filter = New EventTypeFilter(SourceLevels.Error)
            Case "Verbose"
                Globals.theLogListener.Filter = New EventTypeFilter(SourceLevels.Verbose)
            Case "All"
                Globals.theLogListener.Filter = New EventTypeFilter(SourceLevels.All)
        End Select
        Trace.Refresh()
        ' by refreshing the Trace with a different filter, the LogListener gets lost sometimes...
        If Not Trace.Listeners.Contains(Globals.theLogListener) Then Trace.Listeners.Add(Globals.theLogListener)
        Globals.EventLevelSelected = EventLevels.SelectedItem
    End Sub

    Private Sub FixLegacyFunctions_Click(sender As Object, e As EventArgs) Handles FixLegacyFunctions.Click
        repairLegacyFunctions(True)
    End Sub

    Private Sub DisableAddin_Click(sender As Object, e As EventArgs) Handles disableAddin.Click
        ' first reactivate legacy Addin
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\DBAddin.Connection", "LoadBehavior", 3)
        ExcelDnaUtil.Application.AddIns("DBAddin.Functions").Installed = True
        ErrorMsg("Please restart Excel to make changes effective...", "Disable DBAddin and re-enable Legacy DBAddin", MsgBoxStyle.Exclamation)
        Try : ExcelDnaUtil.Application.AddIns("OebfaFuncs").Installed = False : Catch ex As Exception : End Try
        disableAddinAfterwards = True
        Me.Close()
    End Sub
End Class