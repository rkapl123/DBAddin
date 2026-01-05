Imports System.Collections.Generic
Imports System.Data
Imports System.Data.Common
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms


''' <summary>User-form for ad-hoc SQL execution</summary>
Public Class AdHocSQL
    ''' <summary>common connection settings factored in helper class</summary>
    Private myDBConnHelper As DBConnHelper
    ''' <summary>stored environment to reset after change</summary>
    Private ReadOnly storedUserSetEnv As String = ""
    ''' <summary>stored database to reset after change</summary>
    Private ReadOnly userSetDB As String = ""
    ''' <summary>needed to avoid escape key pressed in DBDocumentation from propagating to main AdHocSQL dialog (and closing this dialog therefore)</summary>
    Public propagatedFromDoc As Boolean = False
    ''' <summary>for cancelling the asynchronous execution of the SQL command</summary>
    Private cts As CancellationTokenSource
    ''' <summary>fetch elapsed time in Timer to show after completion</summary>
    Private elapsedTime As DateTime
    ''' <summary>Timer for progress display</summary>
    Public Timer As System.Timers.Timer
    ''' <summary>for counting received rows during execution of query commands</summary>
    Private rowsCount As Integer
    ''' <summary></summary>
    Private batchSize As Integer = 1000

    ''' <summary>create new AdHocSQL dialog</summary>
    ''' <param name="SQLString">SQL string passed from combo-box</param>
    ''' <param name="AdHocSQLStringsIndex">index of SQLstring in combo-box needed to get the environment for this string</param>
    Public Sub New(SQLString As String, AdHocSQLStringsIndex As Integer)
        ' This call is required by the designer.
        InitializeComponent()
        createConfigTreeMenu()
        Me.SQLText.Text = SQLString
        Me.TransferType.Items.Clear()
        For Each TransType As String In {"Cell", "ListFetch", "RowFetch", "ListObject", "Pivot"}
            Me.TransferType.Items.Add(TransType)
        Next
        Me.TransferType.SelectedIndex = Me.TransferType.Items.IndexOf(fetchSetting("AdHocSQLTransferType", "Cell"))

        Me.EnvSwitch.Items.Clear()
        For Each env As String In environdefs
            Me.EnvSwitch.Items.Add(env)
        Next
        Dim userSetEnv As String = fetchSetting("AdHocSQLcmdEnv" + AdHocSQLStringsIndex.ToString(), env(baseZero:=True))
        ' get settings for connection
        myDBConnHelper = New DBConnHelper((Integer.Parse(userSetEnv) + 1).ToString())
        ' issue warning if current selected environment not same as that stored for command (prod/test !)
        If userSetEnv <> env(baseZero:=True) Then
            If QuestionMsg("Current selected environment different from the environment stored for AdHocSQLcmd, change to this environment?",, "AdHoc SQL Command", MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2) = MsgBoxResult.Ok Then
                storedUserSetEnv = userSetEnv
                userSetEnv = env(baseZero:=True)
                myDBConnHelper = New DBConnHelper(env())
            End If
        End If
        userSetDB = fetchSetting("AdHocSQLcmdDB" + AdHocSQLStringsIndex.ToString(), fetchSubstr(myDBConnHelper.dbsheetConnString, myDBConnHelper.dbidentifier, ";"))
        fillDatabasesAndSetDropDown()
        Me.EnvSwitch.SelectedIndex = Integer.Parse(userSetEnv)
        Me.Database.SelectedIndex = Me.Database.Items.IndexOf(userSetDB)
    End Sub

    ''' <summary>execution of the selected ribbon dropdown command right after dialog has been set up, otherwise GUI elements are not available</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AdHocSQL_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        ' only if SQL command not empty and not consisting of spaces only...
        If Strings.Replace(Me.SQLText.Text, " ", "") <> "" Then executeSQL().ConfigureAwait(False)
    End Sub

    ''' <summary>fill the Database dropdown</summary>
    Private Sub fillDatabasesAndSetDropDown()
        Try
            fillDatabases()
        Catch ex As System.Exception
            UserMsg("fillDatabasesAndSetDropDown:" + ex.Message)
            Exit Sub
        End Try
    End Sub

    ''' <summary>fills all possible databases of current connection using db proprietary code in dbGetAllStr, data coming from field DBGetAllFieldName</summary>
    Private Sub fillDatabases()
        Dim dbs As IDataReader

        Me.Database.Items.Clear()
        ' do not catch exception here, as it should be handled by fillDatabasesAndSetDropDown
        myDBConnHelper.openConnection()
        Dim sqlCommand As IDbCommand = myDBConnHelper.getCommand(myDBConnHelper.dbGetAllStr)
        Try
            dbs = sqlCommand.ExecuteReader()
        Catch ex As Exception
            Throw New Exception("Could not retrieve schema information for databases in connection string: '" + myDBConnHelper.dbsheetConnString + "',error: " + ex.Message)
        End Try
        If dbs.FieldCount > 0 Then
            Try
                While dbs.Read()
                    Dim addVal As String
                    If Strings.Len(myDBConnHelper.DBGetAllFieldName) = 0 Then
                        addVal = dbs(0)
                    Else
                        addVal = dbs(myDBConnHelper.DBGetAllFieldName)
                    End If
                    Me.Database.Items.Add(addVal)
                End While
                dbs.Close()
            Catch ex As Exception
                Throw New Exception("Exception when filling DatabaseComboBox: " + ex.Message)
            End Try
        Else
            Throw New Exception("Could not retrieve any databases with: " + myDBConnHelper.dbGetAllStr + "!")
        End If
    End Sub

    ''' <summary>Change Environment in AdHocSQL</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Environment_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles EnvSwitch.SelectionChangeCommitted
        ' reset connection, recreate DB Connection helper and refill database dropdown
        myDBConnHelper.dbshcnn = Nothing
        myDBConnHelper = New DBConnHelper((Me.EnvSwitch.SelectedIndex + 1).ToString())
        SettingsTools.selectedEnvironment = Me.EnvSwitch.SelectedIndex
        theRibbon.InvalidateControl("envDropDown")
        Dim PrevSelDB As String = Me.Database.Text
        fillDatabasesAndSetDropDown()
        ' reset previously set database
        If PrevSelDB = "" Then Exit Sub
        If Me.Database.Items.IndexOf(PrevSelDB) = -1 Then
            UserMsg("Previously selected database '" + PrevSelDB + "' doesn't exist in this environment !", "AdHoc SQL Command")
        End If
        Me.Database.SelectedIndex = Me.Database.Items.IndexOf(PrevSelDB)
    End Sub

    ''' <summary>database changed</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Database_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Database.SelectionChangeCommitted
        ' add database information in calling openConnection to signal a change in connection string to selected database !
        Try
            myDBConnHelper.openConnection(Me.Database.Text)
        Catch ex As System.Exception
            UserMsg("Exception in Database_SelectionChangeCommitted: " + ex.Message)
        End Try
    End Sub

    ''' <summary>executing the SQL command and passing the results to the results pane</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Async Sub Execute_Click(sender As Object, e As EventArgs) Handles Execute.Click
        Await executeSQL().ConfigureAwait(False)
    End Sub

    ''' <summary>after confirmation for non select statements (DML), execute the command given in SQLText by running fill_dgv_Async</summary>
    Private Async Function executeSQL() As Task
        ' only select commands are executed immediately, others are asked for (with default button being cancel)
        If InStr(Strings.LTrim(SQLText.Text.ToLower()), "insert ") > 1 Or InStr(Strings.LTrim(SQLText.Text.ToLower()), "update ") > 1 Or InStr(Strings.LTrim(SQLText.Text.ToLower()), "delete ") > 1 Then
            If LCase(fetchSetting("DMLStatementsAllowed", "False")) <> "true" Then
                UserMsg("Non Select Statements (DML) are forbidden (DMLStatementsAllowed needs to be True) !", "AdHoc SQL Command")
                Exit Function
            End If
            If QuestionMsg("Do you really want to execute the command ?",, "AdHoc SQL Command", MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2) = vbCancel Then Exit Function
        End If

        elapsedTime = New DateTime(0)
        Timer = New System.Timers.Timer(1000)
        AddHandler Timer.Elapsed, AddressOf Timer_Tick
        Timer.Enabled = True
        Timer.Start()

        Execute.Enabled = False
        CloseBtn.Text = "Cancel"
        Me.RowsReturned.Text = ""
        cts = New CancellationTokenSource()
        Try
            Await fill_dgv_Async(cts.Token).ConfigureAwait(False)
            If cts.Token.IsCancellationRequested Then
                Me.RowsReturned.Text = Me.RowsReturned.Text + " (cancelled)"
            End If
        Catch oce As OperationCanceledException
            Me.RowsReturned.Text = Me.RowsReturned.Text + " (cancelled)"
        Catch ex As Exception
            Me.AdHocSQLQueryResult.Rows.Clear()
            Me.AdHocSQLQueryResult.Columns.Clear()
            Me.AdHocSQLQueryResult.Columns.Add("result", "command_result:")
            If rowsCount = 0 And cts.Token.IsCancellationRequested Then
                Me.AdHocSQLQueryResult.Rows.Add("DML execution cancelled.")
            Else
                Me.AdHocSQLQueryResult.Rows.Add("Error: " & ex.Message)
            End If
            Me.RowsReturned.Text = ""
        End Try
        cts.Dispose()
        cts = Nothing

        Timer.Stop()
        Timer.Dispose()
        Timer = Nothing
        ' resize after all columns are available
        AdHocSQLQueryResult.AutoResizeColumns(DataGridViewAutoSizeColumnMode.DisplayedCells)
        Execute.Enabled = True
        CloseBtn.Text = "Close"
    End Function

    ''' <summary>opens connection, creates command from SQLText and calls ExecuteReaderAsync to asynchronously fill the AdHocSQLQueryResult datagridview</summary>
    ''' <param name="ct"></param>
    ''' ''' <returns></returns>
    Public Async Function fill_dgv_Async(ct As CancellationToken) As Task
        Me.AdHocSQLQueryResult.BeginInvoke(Sub()
                                               Me.AdHocSQLQueryResult.Rows.Clear()
                                               Me.AdHocSQLQueryResult.Columns.Clear()
                                           End Sub)
        Try
            myDBConnHelper.openConnection(Me.Database.Text)
        Catch ex As System.Exception
            UserMsg("Exception in fill_dgv_Async (opening Database connection): " + ex.Message)
        End Try
        Dim SqlCmd As DbCommand = myDBConnHelper.getCommand(SQLText.Text)
        SqlCmd.CommandTimeout = 0 ' infinite timeout as the command can be cancelled anytime.
        SqlCmd.CommandType = CommandType.Text

        rowsCount = 0
        Using reader = Await SqlCmd.ExecuteReaderAsync(CommandBehavior.SequentialAccess Or CommandBehavior.CloseConnection, ct).ConfigureAwait(False)
            If reader.FieldCount > 0 Then
                Dim columnsCreated As Boolean = False
                Dim batch As New List(Of Object())()
                Dim statusText As String = ""
                While Await reader.ReadAsync(ct).ConfigureAwait(False)
                    If ct.IsCancellationRequested Then Exit While
                    Dim vals(reader.FieldCount - 1) As Object
                    reader.GetValues(vals)
                    ' create column headers only once
                    If Not columnsCreated Then
                        columnsCreated = True
                        Dim colNames(reader.FieldCount - 1) As String
                        For i As Integer = 0 To reader.FieldCount - 1
                            colNames(i) = reader.GetName(i)
                        Next
                        Dim namesCopy = CType(colNames.Clone(), String())
                        Me.AdHocSQLQueryResult.BeginInvoke(Sub()
                                                               For i As Integer = 0 To namesCopy.Length - 1
                                                                   Me.AdHocSQLQueryResult.Columns.Add("c" & i.ToString(), namesCopy(i))
                                                               Next
                                                           End Sub)
                    End If
                    ' adding batches of data to the datagrid view
                    batch.Add(CType(vals.Clone(), Object()))
                    rowsCount += 1
                    statusText = "returned rows: " + rowsCount.ToString("N0") + " (" + elapsedTime.ToString("T") + ")"
                    If batch.Count >= batchSize Then
                        Me.AdHocSQLQueryResult.BeginInvoke(New delegate_refresh(AddressOf refresh_dgv), batch)
                        ' important to separately set RowsReturned.Text as in the AdHocSQLQueryResult.BeginInvoke it's blocking the UI
                        Me.RowsReturned.Invoke(Sub()
                                                   Me.RowsReturned.Text = statusText
                                               End Sub)
                        batch = New List(Of Object())()
                    End If
                End While
                ' flush remainder
                If batch.Count > 0 And Not ct.IsCancellationRequested Then
                    Me.AdHocSQLQueryResult.BeginInvoke(New delegate_refresh(AddressOf refresh_dgv), batch)
                    Me.RowsReturned.Invoke(Sub()
                                               Me.RowsReturned.Text = statusText
                                           End Sub)
                End If
            Else
                Me.AdHocSQLQueryResult.BeginInvoke(Sub()
                                                       Me.AdHocSQLQueryResult.SuspendLayout()
                                                       Me.AdHocSQLQueryResult.Columns.Add("result", "command_result:")
                                                       Me.AdHocSQLQueryResult.Rows.Add(reader.RecordsAffected.ToString("N0") + " record(s) affected.")
                                                       Me.AdHocSQLQueryResult.ResumeLayout()
                                                   End Sub)
            End If
        End Using
    End Function

    Delegate Sub delegate_refresh(batch As List(Of Object()))

    ''' <summary>refresh the datagrid view in the UI thread (using AdHocSQLQueryResult.BeginInvoke)</summary>
    ''' <param name="batch">the list of Object arrays with the data passed from fill_dgv_Async thread</param>
    Private Sub refresh_dgv(batch As List(Of Object()))
        Me.AdHocSQLQueryResult.SuspendLayout()
        For Each r As Object() In batch
            Me.AdHocSQLQueryResult.Rows.Add(r)
        Next
        Me.AdHocSQLQueryResult.ResumeLayout()
    End Sub

    ''' <summary>to show progress during execution, add elapsedTime (accessed in fill_dgv_Async via Me.RowsReturned.Invoke)</summary>
    ''' <param name="source"></param>
    ''' <param name="e"></param>
    Sub Timer_Tick(source As Object, e As EventArgs)
        elapsedTime = elapsedTime.AddSeconds(1.0)
        ' in case of DML just show time progressing
        If rowsCount = 0 Then
            Me.RowsReturned.Invoke(Sub()
                                       Me.RowsReturned.Text = "DML executing (" + elapsedTime.ToString("T") + ")"
                                   End Sub)
        End If
    End Sub


    ''' <summary>"Transfer": close dialog with OK result</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Transfer_Click(sender As Object, e As EventArgs) Handles Transfer.Click
        finishForm(DialogResult.OK)
    End Sub

    ''' <summary>"Close": close dialog with Cancel result</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CloseBtn_Click(sender As Object, e As EventArgs) Handles CloseBtn.Click
        finishForm(DialogResult.Cancel)
    End Sub

    ''' <summary>common procedure to close the form, regarding (canceling) a running excution</summary>
    Private Sub finishForm(theDialogResult As DialogResult)
        ' Close button (esc) is used as Cancel button during execution
        If theDialogResult = DialogResult.Cancel AndAlso cts IsNot Nothing Then
            cts.Cancel()
            Exit Sub
        End If

        ' get rid of leading and trailing blanks for dropdown and combo box presets
        Me.SQLText.Text = Strings.Trim(Me.SQLText.Text)
        ' if the user environment was changed to the currently selected (global) one, reset it here to the passed one...
        If storedUserSetEnv <> "" Then
            Me.EnvSwitch.SelectedIndex = Integer.Parse(storedUserSetEnv)
            Me.Database.SelectedIndex = Me.Database.Items.IndexOf(userSetDB)
        End If
        Me.DialogResult = theDialogResult
        myDBConnHelper = Nothing
        Me.Hide()
    End Sub

    ''' <summary>keyboard shortcuts for executing (ctrl-return), Transfer (shift-return) and maybe other things in the future (auto-complete)</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub SQLText_KeyDown(sender As Object, e As KeyEventArgs) Handles SQLText.KeyDown
        If e.KeyCode = Keys.Return And e.Modifiers = Keys.Control Then
            e.SuppressKeyPress = True
            executeSQL().ConfigureAwait(False)
        ElseIf e.KeyCode = Keys.Return And e.Modifiers = Keys.Shift Then
            e.SuppressKeyPress = True
            finishForm(DialogResult.OK)
        End If
        ' override paste key combinations to avoid pasting rich text into edit box
        If (e.Modifiers = Keys.Control And e.KeyCode = Keys.V) Then
            Me.SQLText.Paste(DataFormats.GetFormat("Text"))
            e.Handled = True
        End If
    End Sub

    ''' <summary>when being on the database also allow Ctrl-Enter</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Database_KeyDown(sender As Object, e As KeyEventArgs) Handles Database.KeyDown
        If e.KeyCode = Keys.Return And e.Modifiers = Keys.Control Then
            e.SuppressKeyPress = True
            executeSQL().ConfigureAwait(False)
        End If
    End Sub

    ''' <summary>when being on the TransferType selection also allow Shift-Enter</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TransferType_KeyDown(sender As Object, e As KeyEventArgs) Handles TransferType.KeyDown
        If e.KeyCode = Keys.Return And e.Modifiers = Keys.Shift Then
            e.SuppressKeyPress = True
            finishForm(DialogResult.OK)
        End If
    End Sub

    ''' <summary>show context menu for SQLText, displaying config menu as a MenuStrip</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub SQLText_MouseDown(sender As Object, e As MouseEventArgs) Handles SQLText.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then
            Me.ConfigMenuStrip = ConfigFiles.ConfigContextMenu
            Me.ConfigMenuStrip.Show(DirectCast(sender, RichTextBox).PointToScreen(e.Location))
        End If
    End Sub

    ''' <summary>needed together with KeyPreview=True on form to simulate ESC canceling the form and catching this successfully
    ''' (preventing closing when canceling an ongoing sql-command), also see DBDocumentation.vb
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AdHocSQL_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        If e.KeyCode = Keys.Escape And Not propagatedFromDoc Then finishForm(DialogResult.Cancel)
        propagatedFromDoc = False
    End Sub
End Class