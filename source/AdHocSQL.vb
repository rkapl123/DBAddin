Imports System.Data
Imports System.Data.Odbc
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Windows.Forms
Imports System.ComponentModel

Public Class AdHocSQL
    ''' <summary>common connection settings factored in helper class</summary>
    Private myDBConnHelper As DBSheetConnHelper

    Public Sub New(SQLString As String)
        ' This call is required by the designer.
        InitializeComponent()
        Me.SQLText.Text = SQLString
        ' get settings for connection
        myDBConnHelper = New DBSheetConnHelper()

        Me.TransferType.Items.Clear()
        For Each TransType As String In {"Cell", "ListFetch", "RowFetch", "ListObject", "Pivot"}
            Me.TransferType.Items.Add(TransType)
        Next
        Me.TransferType.SelectedIndex = Me.TransferType.Items.IndexOf(fetchSetting("AdHocSQLTransferType", "Cell"))

        Me.EnvSwitch.Items.Clear()
        For Each env As String In Globals.environdefs
            Me.EnvSwitch.Items.Add(env)
        Next
        Dim userSetEnv As String = fetchSetting("AdHocSQLEnvironment", "")
        Me.EnvSwitch.SelectedIndex = If(userSetEnv = "", Globals.selectedEnvironment, Integer.Parse(userSetEnv))
        fillDatabasesAndSetDropDown()
    End Sub

    ''' <summary>execution of ribbon entered command after dialog has been set up, otherwise GUI elements are not available</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AdHocSQL_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        ' only if SQL command not empty and not consisting of spaces only...
        If Strings.Replace(Me.SQLText.Text, " ", "") <> "" Then executeSQL()
    End Sub

    ''' <summary>fill the Database dropdown and set dropdown to database set in connection string</summary>
    Private Sub fillDatabasesAndSetDropDown()
        Try
            fillDatabases()
        Catch ex As System.Exception
            Globals.UserMsg(ex.Message)
            Exit Sub
        End Try
        Me.Database.SelectedIndex = Me.Database.Items.IndexOf(fetchSetting("AdHocSQLDatabase" + Me.EnvSwitch.SelectedIndex.ToString(), Globals.fetch(myDBConnHelper.dbsheetConnString, myDBConnHelper.dbidentifier, ";")))
    End Sub

    ''' <summary>fills all possible databases of current connection using db proprietary code in dbGetAllStr, data coming from field DBGetAllFieldName</summary>
    Private Sub fillDatabases()
        Dim dbs As IDataReader

        Me.Database.Items.Clear()
        ' do not catch exception here, as it should be handled by fillDatabasesAndSetDropDown
        myDBConnHelper.openConnection()
        Dim sqlCommand As IDbCommand
        If TypeName(myDBConnHelper.dbshcnn) = "SqlConnection" Then
            sqlCommand = New SqlCommand(myDBConnHelper.dbGetAllStr, myDBConnHelper.dbshcnn)
        ElseIf TypeName(myDBConnHelper.dbshcnn) = "OleDbConnection" Then
            sqlCommand = New OleDbCommand(myDBConnHelper.dbGetAllStr, myDBConnHelper.dbshcnn)
        Else
            sqlCommand = New OdbcCommand(myDBConnHelper.dbGetAllStr, myDBConnHelper.dbshcnn)
        End If
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
        Globals.selectedEnvironment = Me.EnvSwitch.SelectedIndex
        Globals.ConstConnString = fetchSetting("ConstConnString" + Globals.env(), "")
        ' reset connection and refill database dropdown
        myDBConnHelper.dbshcnn = Nothing
        fillDatabasesAndSetDropDown()
    End Sub

    ''' <summary>database changed</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Database_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Database.SelectionChangeCommitted
        ' add database information in calling openConnection to signal a change in connection string to selected database !
        Try
            myDBConnHelper.openConnection(Me.Database.Text)
        Catch ex As System.Exception
            Globals.UserMsg("Exception in Database_SelectionChangeCommitted: " + ex.Message)
        End Try
    End Sub

    ''' <summary>executing the SQL command and passing the results to the results pane</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Execute_Click(sender As Object, e As EventArgs) Handles Execute.Click
        executeSQL()
    End Sub

    Private Sub executeSQL()
        If Not BackgroundWorker1.IsBusy Then
            ' only select commands are executed immediately, others are asked for (with default button being cancel)
            If InStr(Strings.LTrim(SQLText.Text.ToLower()), "select") <> 1 Then
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
    Private SqlCmd As IDbCommand
    Private nonRowResult As String
    Private dt As DataTable

    ''' <summary>start sql command and load data into datatable in the background (to show progress and have cancellation control)</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim theResult As IDataReader
        Try
            myDBConnHelper.openConnection(Me.Database.Text)
        Catch ex As System.Exception
            Globals.UserMsg("Exception in BackgroundWorker1_DoWork (opeining Database connection): " + ex.Message)
        End Try

        ' set up command
        If TypeName(myDBConnHelper.dbshcnn) = "SqlConnection" Then
            SqlCmd = New SqlCommand(SQLText.Text, myDBConnHelper.dbshcnn)
        ElseIf TypeName(myDBConnHelper.dbshcnn) = "OleDbConnection" Then
            SqlCmd = New OleDbCommand(SQLText.Text, myDBConnHelper.dbshcnn)
        Else
            SqlCmd = New OdbcCommand(SQLText.Text, myDBConnHelper.dbshcnn)
        End If
        SqlCmd.CommandTimeout = Globals.CmdTimeout
        SqlCmd.CommandType = CommandType.Text
        ' execute command on DB Server
        nonRowResult = ""
        Try
            theResult = SqlCmd.ExecuteReader()
        Catch ex As Exception
            nonRowResult = ex.Message + " (" + elapsedTime.ToString("T") + ")"
            theResult = Nothing
        End Try
        ' get results (into data table)
        If Not IsNothing(theResult) Then
            ' for row returning results (select/storedprocedures)
            If theResult.FieldCount > 0 Then
                dt = New DataTable()
                Try
                    dt.Load(theResult)
                Catch ex As Exception
                    nonRowResult = ex.Message
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
            ' row returning results: display row count and elapsed time and pass datatable to datagrid
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

    Private elapsedTime As DateTime

    ''' <summary>show progress during BackgroundWorker1 execution</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
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

    ''' <summary>common procedure to close the form, regarding (cancelling) a busy backgroundworker = sqlcmd)</summary>
    Private Sub finishForm(theDialogResult As DialogResult)
        If BackgroundWorker1.IsBusy Then
            If Globals.QuestionMsg("Cancel the running SQL Command ?",, "AdHoc SQL Command") = MsgBoxResult.Cancel Then Exit Sub
            SqlCmd.Cancel()
            BackgroundWorker1.CancelAsync()
        End If
        ' store the combobox values for later...
        setUserSetting("AdHocEnvironment", Me.EnvSwitch.SelectedIndex.ToString())
        setUserSetting("AdHocSQLTransferType", Me.TransferType.Text)
        setUserSetting("AdHocSQLDatabase" + Me.EnvSwitch.SelectedIndex.ToString(), Me.Database.Text)
        Me.DialogResult = theDialogResult
        Me.Hide()
    End Sub

    ''' <summary>keyboard shortcuts for executing (ctrl-return), Transfer (shift-return) and maybe other things in the future (autocomplete)</summary>
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

    ''' <summary>needed together with KeyPreview=True on form to simulate ESC cancelling the form and catching this successfully (preventing closing when cancelling an ongoing sqlcommand)</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AdHocSQL_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        If e.KeyCode = Keys.Escape Then finishForm(DialogResult.Cancel)
    End Sub
End Class