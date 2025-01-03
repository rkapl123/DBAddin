Imports System.ComponentModel ' for BackgroundWorker callback handling
Imports System.Data
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

    ''' <summary>execution of ribbon entered command after dialog has been set up, otherwise GUI elements are not available</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AdHocSQL_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        ' only if SQL command not empty and not consisting of spaces only...
        If Strings.Replace(Me.SQLText.Text, " ", "") <> "" Then executeSQL()
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
    Private Sub Execute_Click(sender As Object, e As EventArgs) Handles Execute.Click
        executeSQL()
    End Sub

    ''' <summary>after confirmation for non select statements (DML), execute the SQL by running the BackgroundWorker1</summary>
    Private Sub executeSQL()
        If Not BackgroundWorker1.IsBusy Then
            ' only select commands are executed immediately, others are asked for (with default button being cancel)
            If InStr(Strings.LTrim(SQLText.Text.ToLower()), "select") <> 1 Then
                If LCase(fetchSetting("DMLStatementsAllowed", "False")) <> "true" Then
                    UserMsg("Non Select Statements (DML) are forbidden (DMLStatementsAllowed needs to be True) !", "AdHoc SQL Command")
                    Exit Sub
                End If
                If QuestionMsg("Do you really want to execute the command ?",, "AdHoc SQL Command", MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2) = vbCancel Then Exit Sub
            End If
            elapsedTime = New DateTime(0)
            Timer1.Interval = 1000
            Timer1.Enabled = True
            Timer1.Start()
            BackgroundWorker1.RunWorkerAsync()
        End If
    End Sub

    ' variables needed for passing data between background worker and main thread
    ''' <summary>the command object for executing the AdHocSQL command text</summary>
    Private SqlCmd As IDbCommand
    ''' <summary>used to pass non row results between BackgroundWorker1_DoWork and BackgroundWorker1_RunWorkerCompleted</summary>
    Private nonRowResult As String
    ''' <summary>the resulting data table object loaded by BackgroundWorker1_DoWork and displayed in BackgroundWorker1_RunWorkerCompleted</summary>
    Private dt As DataTable

    ''' <summary>start sql command and load data into data table in the background (to show progress and have cancellation control)</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim theResult As IDataReader
        Try
            myDBConnHelper.openConnection(Me.Database.Text)
        Catch ex As System.Exception
            UserMsg("Exception in BackgroundWorker1_DoWork (opening Database connection): " + ex.Message)
        End Try

        ' set up command
        SqlCmd = myDBConnHelper.getCommand(SQLText.Text)
        SqlCmd.CommandTimeout = CmdTimeout
        SqlCmd.CommandType = CommandType.Text
        ' execute command on DB Server
        nonRowResult = ""
        Try
            theResult = SqlCmd.ExecuteReader()
        Catch ex As Exception
            ' interruption of SqlCmd leads to exception which is misleading...
            nonRowResult = If(Not BackgroundWorker1.CancellationPending, ex.Message + " (" + elapsedTime.ToString("T") + ")", "Execution was interrupted..")
            theResult = Nothing
        End Try
        ' get results (into data table)
        If Not IsNothing(theResult) Then
            ' for row returning results (select/stored procedures)
            If theResult.FieldCount > 0 Then
                dt = New DataTable()
                Try
                    dt.Load(theResult)
                Catch ex As Exception
                    ' interruption of SqlCmd leads to exception which is misleading...
                    nonRowResult = If(Not BackgroundWorker1.CancellationPending, ex.Message + " (" + elapsedTime.ToString("T") + ")", "Execution was interrupted..")
                End Try
            Else
                ' DML: insert/update/delete returns no rows, only records affected
                nonRowResult = theResult.RecordsAffected.ToString() + " record(s) affected. (" + elapsedTime.ToString("T") + ")"
            End If
            theResult.Close()
        End If
    End Sub

    ''' <summary>sql command finished, show results. All GUI related work needs to be done in the main thread</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        ' for non row returning results (DML/errors) show returned message
        If nonRowResult <> "" Then
            AdHocSQLQueryResult.DataSource = Nothing
            AdHocSQLQueryResult.Columns.Clear()
            AdHocSQLQueryResult.Columns.Add("result", "command_result:")
            AdHocSQLQueryResult.Rows.Clear()
            AdHocSQLQueryResult.Rows.Add(nonRowResult)
            Me.RowsReturned.Text = ""
        Else
            ' row returning results: display row count and elapsed time and pass data table to data grid
            Me.RowsReturned.Text = dt.Rows.Count.ToString() + " rows returned. (" + elapsedTime.ToString("T") + ")"
            AdHocSQLQueryResult.Columns.Clear()
            AdHocSQLQueryResult.AutoGenerateColumns = True
            AdHocSQLQueryResult.DataSource = dt
            AdHocSQLQueryResult.Refresh()
        End If
        AdHocSQLQueryResult.AutoResizeColumns(DataGridViewAutoSizeColumnMode.DisplayedCells)
        Timer1.Enabled = False
        Timer1.Stop()
    End Sub

    ''' <summary>fetch elapsed time in Timer to show after completion</summary>
    Private elapsedTime As DateTime

    ''' <summary>show progress during BackgroundWorker1 execution</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If BackgroundWorker1.CancellationPending Then Exit Sub
        elapsedTime = elapsedTime.AddSeconds(1.0)
        Me.RowsReturned.Text = "(" + elapsedTime.ToString("T") + ")"
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
    Private Sub Close_Click(sender As Object, e As EventArgs) Handles CloseBtn.Click
        finishForm(DialogResult.Cancel)
    End Sub

    ''' <summary>common procedure to close the form, regarding (canceling) a busy background worker = sqlcmd)</summary>
    Private Sub finishForm(theDialogResult As DialogResult)
        If BackgroundWorker1.IsBusy Then
            If QuestionMsg("Cancel the running SQL Command ?",, "AdHoc SQL Command") = MsgBoxResult.Cancel Then Exit Sub
            SqlCmd.Cancel()
            BackgroundWorker1.CancelAsync()
            If QuestionMsg("Also close the Ad-hoc SQL Command Tool now ?",, "AdHoc SQL Command") = MsgBoxResult.Cancel Then Exit Sub
        End If
        ' get rid of leading and trailing blanks for dropdown and combo box presets
        Me.SQLText.Text = Strings.Trim(Me.SQLText.Text)
        ' if the user environment was changed to the currently selected (global) one, reset it here to the passed one...
        If storedUserSetEnv <> "" Then
            Me.EnvSwitch.SelectedIndex = Integer.Parse(storedUserSetEnv)
            Me.Database.SelectedIndex = Me.Database.Items.IndexOf(userSetDB)
        End If
        Me.DialogResult = theDialogResult
        Me.Hide()
    End Sub

    ''' <summary>keyboard shortcuts for executing (ctrl-return), Transfer (shift-return) and maybe other things in the future (auto-complete)</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub SQLText_KeyDown(sender As Object, e As KeyEventArgs) Handles SQLText.KeyDown
        If e.KeyCode = Keys.Return And e.Modifiers = Keys.Control Then
            e.SuppressKeyPress = True
            executeSQL()
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
            executeSQL()
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

    ''' <summary>For non displayable data (blobs, etc.) that raise an exception, write out the exception in the data-grid cell tool-tip instead of lots of pop-ups</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AdHocSQLQueryResult_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles AdHocSQLQueryResult.DataError
        sender.CurrentRow.Cells(e.ColumnIndex).TooltipText = "Data raised exception: " + e.Exception.Message + " (" + e.Context.ToString() + ")"
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

    ''' <summary>needed together with KeyPreview=True on form to simulate ESC canceling the form and catching this successfully (preventing closing when canceling an ongoing sql-command)</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AdHocSQL_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        If e.KeyCode = Keys.Escape And Not propagatedFromDoc Then finishForm(DialogResult.Cancel)
        propagatedFromDoc = False
    End Sub
End Class