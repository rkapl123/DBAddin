﻿Imports ExcelDna.Integration
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.IO

''' <summary>About box: used to provide information about version/buildtime and links for local help and project homepage</summary>
Public NotInheritable Class AboutBox
    ''' <summary>flag for disabling addin after closing (set on DisableAddin_Click)</summary>
    Public disableAddinAfterwards As Boolean = False
    ''' <summary>flag for quitting excel after closing (set on CheckForUpdates_Click)</summary>
    Public quitExcelAfterwards As Boolean = False
    ''' <summary>when setting EventLevels Listitem at load, prevent event from being fired with this</summary>
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
        Me.LabelVersion.Text = String.Format("Version {0} Buildtime {1}", My.Application.Info.Version.ToString(), sModuleInfo)
        Me.LabelCopyright.Text = My.Application.Info.Copyright
        Me.LabelCompanyName.Text = "Information: " + My.Application.Info.CompanyName
        Me.TextBoxDescription.Text = My.Application.Info.Description
        dontChangeEventLevels = True
        Dim theEventTypeFilter As EventTypeFilter = theLogDisplaySource.Listeners(0).Filter
        Me.EventLevels.SelectedItem = theEventTypeFilter.EventType.ToString()
        dontChangeEventLevels = False
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    ''' <summary>only display result of check (false) or actually perform the update and download new version (true)</summary>
    Dim doUpdate As Boolean = False
    Const AddinName = "DBAddin-"
    Const updateFilenameZip = "downloadedVersion.zip"
    Dim localUpdateFolder As String = fetchSetting("localUpdateFolder", "")
    Dim localUpdateMessage As String = fetchSetting("localUpdateMessage", "A new version is available in the local update folder, quit Excel and open explorer to start deployAddin.cmd ?")
    Dim updatesMajorVersion As String = fetchSetting("updatesMajorVersion", "1.0.0.")
    Dim updatesDownloadFolder As String = fetchSetting("updatesDownloadFolder", "C:\temp\")
    Dim updatesUrlBase As String = fetchSetting("updatesUrlBase", "https://github.com/rkapl123/DBAddin/archive/refs/tags/")
    Dim response As Net.HttpWebResponse = Nothing
    Dim urlFile As String = ""
    ' check for zip file of next higher revision
    Dim curRevision As Integer
    Dim foundARevision As Boolean = False

    ''' <summary>checks for updates of DB-Addin, asks for download and downloads them</summary>
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        If doUpdate Then Exit Sub
        curRevision = My.Application.Info.Version.Revision
        ' try with highest possible Security protocol
        Try
            Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12 Or Net.SecurityProtocolType.SystemDefault
        Catch ex As Exception
            UserMsg("Error setting the SecurityProtocol: " + ex.Message())
            Exit Sub
        End Try

        ' always accept url certificate as valid
        Net.ServicePointManager.ServerCertificateValidationCallback = AddressOf ValidationCallbackHandler
        Dim revisionNotFoundTries As Integer = 0
        Dim triedRevision As Integer = curRevision
        Do
            urlFile = updatesUrlBase + updatesMajorVersion + triedRevision.ToString() + ".zip"
            Dim request As Net.HttpWebRequest
            Try
                request = Net.WebRequest.Create(urlFile)
                response = Nothing
                request.Method = "HEAD"
                response = request.GetResponse()
            Catch ex As Exception
            End Try
            ' if nothing is found this can mean: there is no higher revision available or the available revision is still higher than the tried one...
            If response Is Nothing Then
                revisionNotFoundTries += 1
            Else
                curRevision = triedRevision
                foundARevision = True
                response.Close()
            End If
            triedRevision += 1
        Loop Until revisionNotFoundTries = CInt(fetchSetting("maxTriesForRevisionFind", "10"))
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        ' get out if no newer version found
        If curRevision = My.Application.Info.Version.Revision Then
            If foundARevision Then
                Me.TextBoxDescription.Text = My.Application.Info.Description + vbCrLf + vbCrLf + "You have the latest version (" + updatesMajorVersion + curRevision.ToString() + ")."
                Me.TextBoxDescription.BackColor = Drawing.Color.FromKnownColor(Drawing.KnownColor.Control)
            Else
                Me.TextBoxDescription.Text = My.Application.Info.Description + vbCrLf + vbCrLf + "Version " + updatesMajorVersion + curRevision.ToString() +
                    " was not found on Github, it is probably more than 10 releases behind, reopen the Aboutbox to retry with maxTriesForRevisionFind (currently " +
                    fetchSetting("maxTriesForRevisionFind", "10") + ") increased by 10."
                setUserSetting("maxTriesForRevisionFind", (CInt(fetchSetting("maxTriesForRevisionFind", "10")) + 10).ToString())
                Me.TextBoxDescription.BackColor = Drawing.Color.Violet
            End If
            Me.CheckForUpdates.Text = "no Update ..."
            Me.CheckForUpdates.Enabled = False
            Me.Refresh()
            Exit Sub
        Else
            setUserSetting("maxTriesForRevisionFind", "10")
            Me.TextBoxDescription.Text = My.Application.Info.Description + vbCrLf + vbCrLf + "A new version (" + updatesMajorVersion + curRevision.ToString() + ") is available " +
                IIf(localUpdateFolder <> "", "in " + localUpdateFolder, "on Github")
            Me.TextBoxDescription.BackColor = Drawing.Color.DarkOrange
            Me.CheckForUpdates.Text = "get Update ..."
            Me.CheckForUpdates.Enabled = True
            Me.Refresh()
            If Not doUpdate Then Exit Sub
        End If
        ' if there is a maintained local update folder, open it and let user update from there...
        If localUpdateFolder <> "" Then
            Try
                If QuestionMsg(localUpdateMessage, MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                    System.Diagnostics.Process.Start("explorer.exe", localUpdateFolder)
                    Me.quitExcelAfterwards = True
                    Me.Close()
                End If
            Catch ex As Exception
                UserMsg("Error when opening local update folder: " + ex.Message())
            End Try
            Exit Sub
        End If

        ' continue with download
        urlFile = updatesUrlBase + updatesMajorVersion + curRevision.ToString() + ".zip"

        ' create the download folder
        Try
            IO.Directory.CreateDirectory(updatesDownloadFolder)
        Catch ex As Exception
            UserMsg("Couldn't create file download folder (" + updatesDownloadFolder + "): " + ex.Message())
            Exit Sub
        End Try

        Me.TextBoxDescription.Text = My.Application.Info.Description + vbCrLf + vbCrLf + "Downloading new version from " + urlFile
        Me.Refresh()
        ' get the new version zip-file
        Dim requestGet As Net.HttpWebRequest = Net.WebRequest.Create(urlFile)
        requestGet.Method = "GET"
        Try
            response = requestGet.GetResponse()
        Catch ex As Exception
            UserMsg("Error when downloading new version: " + ex.Message())
            Exit Sub
        End Try
        ' save the version as zip file
        If response IsNot Nothing Then
            Dim receiveStream As Stream = response.GetResponseStream()
            Using downloadFile As IO.FileStream = File.Create(updatesDownloadFolder + updateFilenameZip)
                receiveStream.CopyTo(downloadFile)
            End Using
        End If
        response.Close()
        Me.TextBoxDescription.Text = My.Application.Info.Description + vbCrLf + vbCrLf + "Extracting " + urlFile + " to " + updatesDownloadFolder
        Me.Refresh()
        ' now extract the downloaded file and open the Distribution folder, first remove any existing folder...
        Try
            Directory.Delete(updatesDownloadFolder + AddinName + updatesMajorVersion + curRevision.ToString(), True)
        Catch ex As Exception : End Try
        Try
            Compression.ZipFile.ExtractToDirectory(updatesDownloadFolder + updateFilenameZip, updatesDownloadFolder)
        Catch ex As Exception
            UserMsg("Error when extracting new version: " + ex.Message())
        End Try
        Me.TextBoxDescription.Text = My.Application.Info.Description + vbCrLf + vbCrLf + "New version in " + updatesDownloadFolder + AddinName + updatesMajorVersion + curRevision.ToString() + "\Distribution, start deployAddin.cmd to install the new Version."
        Me.Refresh()
        Try
            System.Diagnostics.Process.Start("explorer.exe", updatesDownloadFolder + AddinName + updatesMajorVersion + curRevision.ToString() + "\Distribution")
        Catch ex As Exception
            UserMsg("Error when opening Distribution folder of new version: " + ex.Message())
        End Try
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
        Dim theEventTypeFilter As New EventTypeFilter(SourceLevels.Off)
        Select Case EventLevels.SelectedItem
            Case "Information"
                theEventTypeFilter = New EventTypeFilter(SourceLevels.Information)
            Case "Warning"
                theEventTypeFilter = New EventTypeFilter(SourceLevels.Warning)
            Case "Error"
                theEventTypeFilter = New EventTypeFilter(SourceLevels.Error)
            Case "Verbose"
                theEventTypeFilter = New EventTypeFilter(SourceLevels.Verbose)
            Case "All"
                theEventTypeFilter = New EventTypeFilter(SourceLevels.All)
        End Select
        theLogDisplaySource.Listeners(0).Filter = theEventTypeFilter
        theLogFileSource.Listeners("FileLogger").Filter = theEventTypeFilter
    End Sub

    Private Sub CheckForUpdates_Click(sender As Object, e As EventArgs) Handles CheckForUpdates.Click
        If Not BackgroundWorker1.IsBusy Then
            doUpdate = True
            BackgroundWorker1.RunWorkerAsync()
        End If
    End Sub

    Private Sub DisableAddin_Click(sender As Object, e As EventArgs) Handles disableAddin.Click
        Try
            ExcelDnaUtil.Application.AddIns("DBAddin.Functions").Installed = True
        Catch ex As Exception
            UserMsg("Legacy DB-Addin not available in Excel-Addins, can't reactivate it, so disabling this Add-in not possible !")
            Exit Sub
        End Try
        ' first reactivate legacy Addin
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\DBAddin.Connection", "LoadBehavior", 3)
        UserMsg("Please restart Excel to make changes effective...", "Disable DBAddin and re-enable Legacy DBAddin", MsgBoxStyle.Exclamation)
        Try : ExcelDnaUtil.Application.AddIns("OebfaFuncs").Installed = False : Catch ex As Exception : End Try
        disableAddinAfterwards = True
        Me.Close()
    End Sub

    Private Function ValidationCallbackHandler() As Boolean
        Return True
    End Function

End Class