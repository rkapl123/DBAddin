''' <summary>About box: used to provide information about version/buildtime and links for local help and project homepage</summary>
Public NotInheritable Class AboutBox1

    ''' <summary>set up Aboutbox</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AboutBox1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim sModuleInfo As String = vbNullString
        ' get module info for buildtime (FileDateTime of xll):
        For Each tModule As ProcessModule In Process.GetCurrentProcess().Modules
            Dim sModule As String = tModule.FileName
            If sModule.ToUpper.Contains("DBADDIN-ADDIN-PACKED.XLL") Or sModule.ToUpper.Contains("DBADDIN-ADDIN64-PACKED.XLL") Then
                sModuleInfo = FileDateTime(sModule).ToString()
            End If
        Next

        ' set the UI elements
        Me.Text = String.Format("About {0}", My.Application.Info.Title)
        Me.LabelProductName.Text = "DB Addin Help (click here)..."
        Me.LabelVersion.Text = String.Format("Version {0} Buildtime {1}", My.Application.Info.Version.ToString, sModuleInfo)
        Me.LabelCopyright.Text = My.Application.Info.Copyright
        Me.LabelCompanyName.Text = "Information: " + My.Application.Info.CompanyName
        Me.TextBoxDescription.Text = My.Application.Info.Description
    End Sub

    ''' <summary>Close Aboutbox</summary>
    Private Sub OKButton_Click(ByVal sender As Object, ByVal e As EventArgs) Handles OKButton.Click
        Me.Close()
    End Sub

    ''' <summary>Click on Project homepage: activate hyperlink in browser</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub LabelCompanyName_Click(sender As Object, e As EventArgs) Handles LabelCompanyName.Click
        Process.Start(My.Application.Info.CompanyName)
    End Sub

    ''' <summary>Click on Local help: activate hyperlink in browser</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub LabelProductName_Click(sender As Object, e As EventArgs) Handles LabelProductName.Click
        Process.Start(fetchSetting("LocalHelp", String.Empty))
    End Sub

End Class
