Imports System.Data
Imports System.Data.Odbc
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Windows.Forms

Public Class AdHocSQL
    ''' <summary>common connection settings factored in helper class</summary>
    Private myDBConnHelper As DBSheetConnHelper

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
    End Sub

    Public Sub New(SQLString As String)
        ' This call is required by the designer.
        InitializeComponent()
        Me.SQLText.Text = SQLString
        ' get settings for connection
        myDBConnHelper = New DBSheetConnHelper()
        fillDatabasesAndSetDropDown()
    End Sub

    ''' <summary>fill the Database dropdown and set dropdown to database set in connection string</summary>
    Private Sub fillDatabasesAndSetDropDown()
        Try
            fillDatabases()
        Catch ex As System.Exception
            Globals.UserMsg(ex.Message)
            Exit Sub
        End Try
        Database.SelectedIndex = Database.Items.IndexOf(Globals.fetch(myDBConnHelper.dbsheetConnString, myDBConnHelper.dbidentifier, ";"))
    End Sub

    ''' <summary>fills all possible databases of current connection using db proprietary code in dbGetAllStr, data coming from field DBGetAllFieldName</summary>
    Private Sub fillDatabases()
        Dim dbs As OdbcDataReader

        ' do not catch exception here, as it should be handled by fillDatabasesAndSetDropDown
        myDBConnHelper.openConnection()
        Database.Items.Clear()
        Dim sqlCommand As OdbcCommand = New OdbcCommand(myDBConnHelper.dbGetAllStr, myDBConnHelper.dbshcnn)
        Try
            dbs = sqlCommand.ExecuteReader()
        Catch ex As OdbcException
            Throw New Exception("Could not retrieve schema information for databases in connection string: '" + myDBConnHelper.dbsheetConnString + "',error: " + ex.Message)
        End Try
        If dbs.HasRows Then
            Try
                While dbs.Read()
                    Dim addVal As String
                    If Strings.Len(myDBConnHelper.DBGetAllFieldName) = 0 Then
                        addVal = dbs(0)
                    Else
                        addVal = dbs(myDBConnHelper.DBGetAllFieldName)
                    End If
                    Database.Items.Add(addVal)
                End While
                dbs.Close()
            Catch ex As System.Exception
                Throw New Exception("Exception when filling DatabaseComboBox: " + ex.Message)
            End Try
        Else
            Throw New Exception("Could not retrieve any databases with: " + myDBConnHelper.dbGetAllStr + "!")
        End If
    End Sub

    ''' <summary>database changed, initialize everything else (Tables, DBSheetCols definition) from scratch</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Database_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Database.SelectedIndexChanged
        ' add database information in calling openConnection to signal a change in connection string to selected database !
        Try : myDBConnHelper.openConnection(Database.Text) : Catch ex As Exception : Globals.UserMsg(ex.Message) : Exit Sub : End Try
        Try
        Catch ex As System.Exception
            Globals.UserMsg("Exception in Database_SelectedIndexChanged: " + ex.Message)
        End Try
    End Sub

    ''' <summary>executing the SQL command and passing the results to the results pane</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Execute_Click(sender As Object, e As EventArgs) Handles Execute.Click
        ' only select commands are executed immediately, others are asked for (with default button being cancel)
        If InStr(SQLText.Text.ToLower(), "select") <> 1 Then
            If QuestionMsg("Do you really want to execute the command ?",, "AdHoc SQL Command", MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2) = vbCancel Then Exit Sub
        End If
        Dim result As IDataReader = Nothing
        Dim SqlCmd As IDbCommand
        If TypeName(idbcnn) = "SqlConnection" Then
            SqlCmd = New SqlClient.SqlCommand(SQLText.Text, myDBConnHelper.dbshcnn)
        ElseIf TypeName(idbcnn) = "OleDbConnection" Then
            SqlCmd = New OleDb.OleDbCommand(SQLText.Text, myDBConnHelper.dbshcnn)
        Else
            SqlCmd = New Odbc.OdbcCommand(SQLText.Text, myDBConnHelper.dbshcnn)
        End If
        SqlCmd.CommandType = CommandType.Text
        Dim nonRowResult As String = ""
        Try
            result = SqlCmd.ExecuteReader()
        Catch ex As Exception
            nonRowResult = ex.Message
            result = Nothing
        End Try

        If Not IsNothing(result) Then
            ' for row returning results (select/storedprocedures)
            If result.FieldCount > 0 Then
                Dim dt = New DataTable()
                Try
                    dt.Load(result)
                Catch ex As Exception
                    nonRowResult = ex.Message
                End Try
                If nonRowResult = "" Then
                    AdHocSQLQueryResult.Columns.Clear()
                    AdHocSQLQueryResult.AutoGenerateColumns = True
                    AdHocSQLQueryResult.DataSource = dt
                    AdHocSQLQueryResult.Refresh()
                End If
            Else
                ' DML: insert/update/delete returns no rows, only records affected
                nonRowResult = result.RecordsAffected.ToString() + " record(s) affected."
            End If
            result.Close()
        End If
        ' for non row returning results (DML/errors) return message
        If nonRowResult <> "" Then
            AdHocSQLQueryResult.DataSource = Nothing
            AdHocSQLQueryResult.Columns.Clear()
            AdHocSQLQueryResult.Columns.Add("result", "command_result:")
            AdHocSQLQueryResult.Rows.Clear()
            AdHocSQLQueryResult.Rows.Add(nonRowResult)
        End If
        AdHocSQLQueryResult.AutoResizeColumns(DataGridViewAutoSizeColumnMode.DisplayedCells)
    End Sub

    ''' <summary>only close dialog here, OK result is set to Transfer button in the designer</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Transfer_Click(sender As Object, e As EventArgs) Handles Transfer.Click
        Me.Hide()
    End Sub

    ''' <summary>execution of ribbon entered command after dialog has been set up, otherwise GUI elements are not available</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AdHocSQL_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        ' only if SQL command not empty and not consisting of spaces only...
        If Strings.Replace(Me.SQLText.Text, " ", "") <> "" Then Execute_Click(Nothing, Nothing)
    End Sub

    ''' <summary>only close dialog here, Cancel result is set to Transfer button in the designer</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Close_Click(sender As Object, e As EventArgs) Handles CloseBtn.Click
        Me.Hide()
    End Sub

    ''' <summary>keyboard shortcuts for executing (ctrl-return), Transfer (shift-return) and other things</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub SQLText_KeyDown(sender As Object, e As KeyEventArgs) Handles SQLText.KeyDown
        If e.KeyCode = Keys.Return And e.Modifiers = Keys.Control Then
            e.SuppressKeyPress = True
            Execute_Click(Nothing, Nothing)
        ElseIf e.KeyCode = Keys.Return And e.Modifiers = Keys.Shift Then
            e.SuppressKeyPress = True
            Me.DialogResult = DialogResult.OK
            Me.Hide()
        End If
    End Sub
End Class