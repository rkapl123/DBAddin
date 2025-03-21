Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.Odbc
Imports System.IO
Imports System.Windows.Forms

''' <summary>Form for defining/creating DBSheet definitions</summary>
Public Class DBSheetCreateForm
    Inherits System.Windows.Forms.Form
    ''' <summary>whether the form fields should react to changes (set if making changes within code)....</summary>
    Private FormDisabled As Boolean
    ''' <summary>placeholder used in lookup queries to identify the current field's lookup table</summary>
    Private tblPlaceHolder As String = "!T!"
    ''' <summary>character prepended before field name to specify non null-able fields</summary>
    Private specialNonNullableChar As String = "*"
    ''' <summary>common connection settings factored in helper class</summary>
    Private myDBConnHelper As DBConnHelper

#Region "Initialization of DBSheetDefs"
    ''' <summary>entry point of form, invoked by clicking "create/edit DBSheet definition"</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DBSheetCreateForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        ' set up columns for DBSheetCols grid-view
        Dim nameCB As New DataGridViewComboBoxColumn With {
                    .Name = "name",
                    .DataSource = New List(Of String),
                    .HeaderText = "name",
                    .DataPropertyName = "name",
                    .SortMode = DataGridViewColumnSortMode.NotSortable
                }
        Dim ftableCB As New DataGridViewComboBoxColumn With {
                    .Name = "ftable",
                    .DataSource = New List(Of String),
                    .HeaderText = "ftable",
                    .DataPropertyName = "ftable",
                    .SortMode = DataGridViewColumnSortMode.NotSortable
                }
        Dim fkeyCB As New DataGridViewComboBoxColumn With {
                    .Name = "fkey",
                    .DataSource = New List(Of String),
                    .HeaderText = "fkey",
                    .DataPropertyName = "fkey",
                    .SortMode = DataGridViewColumnSortMode.NotSortable
                }
        Dim flookupCB As New DataGridViewComboBoxColumn With {
                    .Name = "flookup",
                    .DataSource = New List(Of String),
                    .HeaderText = "flookup",
                    .DataPropertyName = "flookup",
                    .SortMode = DataGridViewColumnSortMode.NotSortable
                }
        Dim outerCB As New DataGridViewCheckBoxColumn With {
                    .Name = "outer",
                    .HeaderText = "outer",
                    .DataPropertyName = "outer",
                    .SortMode = DataGridViewColumnSortMode.NotSortable
                }
        Dim primkeyCB As New DataGridViewCheckBoxColumn With {
                    .Name = "primkey",
                    .HeaderText = "primkey",
                    .DataPropertyName = "primkey",
                    .SortMode = DataGridViewColumnSortMode.NotSortable
                }
        Dim typeTB As New DataGridViewTextBoxColumn With {
                    .Name = "type",
                    .HeaderText = "type",
                    .DataPropertyName = "type",
                    .[ReadOnly] = True,
                    .SortMode = DataGridViewColumnSortMode.NotSortable
                }
        Dim sortCB As New DataGridViewComboBoxColumn With {
                    .Name = "sort",
                    .DataSource = New List(Of String)({"", "ASC", "DESC"}),
                    .HeaderText = "sort",
                    .DataPropertyName = "sort",
                    .SortMode = DataGridViewColumnSortMode.NotSortable
                }
        Dim lookupTB As New DataGridViewTextBoxColumn With {
                    .Name = "lookup",
                    .HeaderText = "lookup",
                    .DataPropertyName = "lookup",
                    .SortMode = DataGridViewColumnSortMode.NotSortable
                }
        DBSheetCols.AutoGenerateColumns = False
        DBSheetCols.Columns.AddRange(nameCB, ftableCB, fkeyCB, flookupCB, outerCB, primkeyCB, typeTB, sortCB, lookupTB)
        finalizeSetup()
    End Sub

    ''' <summary>factored out parts of setup for reuse in Environment_SelectedIndexChanged (after the environment is changed, need to reconnect and load all schema information)</summary>
    Private Sub finalizeSetup()
        ' get settings for DBSheet definition editing
        myDBConnHelper = New DBConnHelper(env())
        tblPlaceHolder = fetchSetting("tblPlaceHolder" + env(), "!T!")
        specialNonNullableChar = fetchSetting("specialNonNullableChar" + env(), "*")
        FormDisabled = True
        Me.Environment.DataSource = environdefs
        Me.Environment.Text = environdefs(env(True))
        FormDisabled = False
        DBSheetColsEditable(False)
        EnvironEditable(True)

        ' if we have a Password to enter (dbPwdSpec contained in dbsheetConnString and no password entered yet), just display explanation text in title bar and let user enter password... 
        If InStr(myDBConnHelper.dbsheetConnString, myDBConnHelper.dbPwdSpec) > 0 And myDBConnHelper.dbPwdSpec <> "" And existingPwd = "" Then
            Me.Text = "DB Sheet creation: Please enter required Password into Field Pwd to access schema information"
            resetDBSheetCreateForm()
        Else ' otherwise jump in immediately
            assignDBSheet.Enabled = False
            ' password-less connection string, reset password and disable...
            If InStr(myDBConnHelper.dbsheetConnString, myDBConnHelper.dbPwdSpec) = 0 Or myDBConnHelper.dbPwdSpec = "" Then
                If InStr(myDBConnHelper.dbsheetConnString, myDBConnHelper.dbPwdSpec) = 0 Then UserMsg("The DB specific password string (" + myDBConnHelper.dbPwdSpec + ") is not contained in connection string:" + myDBConnHelper.dbsheetConnString + ", therefore no password entry is possible")
                Password.Enabled = False
                existingPwd = ""
            Else ' set to stored existing password
                Password.Text = existingPwd
            End If
            fillDatabasesAndSetDropDown()
            ' initialize with empty DBSheet definitions is done by above call, changing Database.SelectedIndex (Database_SelectedIndexChanged)
        End If
    End Sub

    ''' <summary>reset the DBSheet Create form after 1) an error (change environment to get out) or 2) to let user enter password (Password_Leave with setPasswordAndInit afterwards)</summary>
    Private Sub resetDBSheetCreateForm()
        TableEditable(False)
        saveEnabled(False)
        DBSheetColsEditable(False)
        assignDBSheet.Enabled = False
        Try : myDBConnHelper.dbshcnn.Close() : Catch ex As Exception : End Try
        myDBConnHelper.dbshcnn = Nothing
        ' if called by error in openConnection, reset existing password to allow for refreshing...
        existingPwd = ""
    End Sub

    ''' <summary>enter pressed in Password text-box triggering initialization</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Password_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Password.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) And FormLocalPwd <> Password.Text Then setPasswordAndInit()
    End Sub

    ''' <summary>temporary storage for password to check if changed</summary>
    Private FormLocalPwd As String = ""

    ''' <summary>entering Password box to remember local changed password</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Password_Enter(sender As Object, e As EventArgs) Handles Password.Enter
        FormLocalPwd = Password.Text
    End Sub

    ''' <summary>leaving Password text-box triggering initialization</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Password_Leave(sender As Object, e As EventArgs) Handles Password.Leave
        If FormLocalPwd <> Password.Text Then setPasswordAndInit()
    End Sub

    ''' <summary>called after new password has been entered, reset stored password and fill database dropdown</summary>
    Private Sub setPasswordAndInit()
        existingPwd = Password.Text
        FormLocalPwd = Password.Text
        Try : myDBConnHelper.dbshcnn.Close() : Catch ex As Exception : End Try
        myDBConnHelper.dbshcnn = Nothing
        fillDatabasesAndSetDropDown()
    End Sub

    ''' <summary>fill the Database dropdown and set dropdown to database set in connection string</summary>
    Private Sub fillDatabasesAndSetDropDown()
        Try
            fillDatabases()
        Catch ex As System.Exception
            resetDBSheetCreateForm()
            UserMsg(ex.Message)
            Exit Sub
        End Try
        Me.Text = "DB Sheet creation: Select Database and Table to start building a DBSheet Definition"
        Database.SelectedIndex = Database.Items.IndexOf(fetchSubstr(myDBConnHelper.dbsheetConnString, myDBConnHelper.dbidentifier, ";"))
        'initialization of everything else is triggered by above change and caught by Database_SelectedIndexChanged
    End Sub

    ''' <summary>fills all possible databases of current connection using db proprietary code in dbGetAllStr, data coming from field DBGetAllFieldName</summary>
    Private Sub fillDatabases()
        Dim dbs As OdbcDataReader

        ' do not catch exception here, as it should be handled by fillDatabasesAndSetDropDown
        myDBConnHelper.openConnection(usedForDBSheetCreate:=True)
        FormDisabled = True
        Database.Items.Clear()
        Dim sqlCommand As New OdbcCommand(myDBConnHelper.dbGetAllStr, myDBConnHelper.dbshcnn)
        Try
            dbs = sqlCommand.ExecuteReader()
        Catch ex As OdbcException
            FormDisabled = False
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
                FormDisabled = False
            Catch ex As System.Exception
                FormDisabled = False
                Throw New Exception("Exception when filling DatabaseComboBox: " + ex.Message)
            End Try
        Else
            FormDisabled = False
            Throw New Exception("Could not retrieve any databases with: " + myDBConnHelper.dbGetAllStr + "!")
        End If
    End Sub

    ''' <summary>database changed, initialize everything else (Tables, DBSheetCols definition) from scratch</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Database_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Database.SelectedIndexChanged
        ' add database information to signal a change in connection string to selected database !
        Try : myDBConnHelper.openConnection(Database.Text, usedForDBSheetCreate:=True) : Catch ex As Exception : UserMsg(ex.Message) : resetDBSheetCreateForm() : Exit Sub : End Try
        Try
            fillTables()
            ' start with empty columns list
            TableEditable(True)
            saveEnabled(False)
            FormDisabled = True
            Query.Text = ""
            WhereClause.Text = ""
            DBSheetCols.DataSource = Nothing
            DBSheetCols.Rows.Clear()
            FormDisabled = False
        Catch ex As System.Exception
            UserMsg("Exception in Database_SelectedIndexChanged: " + ex.Message)
        End Try
    End Sub

    ''' <summary>selecting the Table triggers enabling the DBSheetCols definition (fills columns/fields of that table and resetting DBSheetCols definition)</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Table_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles Table.SelectedIndexChanged
        If FormDisabled Then Exit Sub
        Try
            FormDisabled = True
            If Table.SelectedIndex >= 0 Then DBSheetColsEditable(True)
            ' just in case this wasn't cleared before...
            DBSheetCols.DataSource = Nothing
            DBSheetCols.Rows.Clear()
            DBSheetCols.DataSource = New DBSheetDefTable()
            Query.Text = ""
            DirectCast(DBSheetCols.Columns("name"), DataGridViewComboBoxColumn).DataSource = getColumns()
            DirectCast(DBSheetCols.Columns("ftable"), DataGridViewComboBoxColumn).DataSource = getforeignTables()
            FormDisabled = False
        Catch ex As System.Exception
            UserMsg("Exception in Table_SelectedIndexChanged: " + ex.Message)
        End Try
        Me.Text = "DB Sheet creation: Select one or more columns (fields) adding possible foreign key lookup information in foreign tables, finally click create query to finish DBSheet definition"
    End Sub


#End Region

#Region "DBSheetCols Grid-view"
    ''' <summary>handles the various changes in the DBSheetCols grid-view</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DBSheetCols_CellValueChanged(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles DBSheetCols.CellValueChanged
        If FormDisabled Then Exit Sub
        FormDisabled = True
        ' lock environment, table and database choice to not let user accidentally clear DBSheet definitions
        TableEditable(False)
        EnvironEditable(False)
        Dim selIndex As Integer = e.RowIndex
        If e.ColumnIndex = 0 Then     ' field name column 
            ' fill empty lists for fkey and flookup combo-boxes (resetting) ..
            DirectCast(DBSheetCols.Rows(selIndex).Cells("fkey"), DataGridViewComboBoxCell).DataSource = New List(Of String)({""})
            DirectCast(DBSheetCols.Rows(selIndex).Cells("flookup"), DataGridViewComboBoxCell).DataSource = New List(Of String)({""})
            ' if first column then always set primary key!
            If selIndex = 0 Then DBSheetCols.Rows(selIndex).Cells("primkey").Value = True
            ' fill field type for current column
            DBSheetCols.Rows(selIndex).Cells("type").Value = TableDataTypes(DBSheetCols.Rows(selIndex).Cells("name").Value)
            ' fill default values ..
            DBSheetCols.Rows(selIndex).Cells("ftable").Value = ""
            DBSheetCols.Rows(selIndex).Cells("fkey").Value = ""
            DBSheetCols.Rows(selIndex).Cells("flookup").Value = ""
            DBSheetCols.Rows(selIndex).Cells("sort").Value = ""
            DBSheetCols.Rows(selIndex).Cells("lookup").Value = ""
        ElseIf e.ColumnIndex = 1 Then  ' ftable column 
            If DBSheetCols.Rows(selIndex).Cells("ftable").Value.ToString() = "" Then ' reset fkey and flookup dropdown
                DirectCast(DBSheetCols.Rows(selIndex).Cells("fkey"), DataGridViewComboBoxCell).DataSource = New List(Of String)({""})
                DirectCast(DBSheetCols.Rows(selIndex).Cells("flookup"), DataGridViewComboBoxCell).DataSource = New List(Of String)({""})
            Else ' fill fkey and flookup with fields of ftable
                Dim forColsList As List(Of String) = getforeignTables(DBSheetCols.Rows(selIndex).Cells("ftable").Value.ToString())
                DirectCast(DBSheetCols.Rows(selIndex).Cells("fkey"), DataGridViewComboBoxCell).DataSource = forColsList
                DirectCast(DBSheetCols.Rows(selIndex).Cells("flookup"), DataGridViewComboBoxCell).DataSource = forColsList
            End If
            DBSheetCols.Rows(selIndex).Cells("fkey").Value = ""
            DBSheetCols.Rows(selIndex).Cells("flookup").Value = ""
        ElseIf e.ColumnIndex = 2 Then ' fkey column
            If DBSheetCols.Rows(selIndex).Cells("fkey").Value.ToString() <> "" AndAlso DBSheetCols.Rows(selIndex).Cells("flookup").Value.ToString() <> "" AndAlso DBSheetCols.Rows(selIndex).Cells("ftable").Value.ToString() <> "" Then
                regenLookupForRow(selIndex, True)
            End If
        ElseIf e.ColumnIndex = 3 Then ' flookup column
            If DBSheetCols.Rows(selIndex).Cells("fkey").Value.ToString() <> "" AndAlso DBSheetCols.Rows(selIndex).Cells("flookup").Value.ToString() <> "" AndAlso DBSheetCols.Rows(selIndex).Cells("ftable").Value.ToString() <> "" Then
                regenLookupForRow(selIndex, True)
            End If
        ElseIf e.ColumnIndex = 5 Then ' primkey column
            Try
                ' not first row selected: check for previous row (field) if also primary column..
                If Not selIndex = 0 Then
                    If Not DBSheetCols.Rows(selIndex - 1).Cells("primkey").Value And DBSheetCols.Rows(selIndex).Cells("primkey").Value Then
                        UserMsg("All primary keys have to be first and there is at least one non-primary key column before that one !", "DBSheet Definition Error")
                        DBSheetCols.Rows(selIndex).Cells("primkey").Value = False
                    End If
                    ' check if next row (field) is primary key column (only for non-last rows)
                    If selIndex <> DBSheetCols.Rows.Count - 2 Then
                        If DBSheetCols.Rows(selIndex + 1).Cells("primkey").Value And Not DBSheetCols.Rows(selIndex).Cells("primkey").Value Then
                            UserMsg("All primary keys have to be first and there is at least one primary key column after that one !", "DBSheet Definition Error")
                            DBSheetCols.Rows(selIndex).Cells("primkey").Value = True
                        End If
                    End If
                ElseIf Not DBSheetCols.Rows(selIndex).Cells("primkey").Value Then
                    UserMsg("first column always has to be primary key", "DBSheet Definition Error")
                    DBSheetCols.Rows(selIndex).Cells("primkey").Value = True
                End If
            Catch ex As System.Exception
                UserMsg("Exception in DBSheetCols_CellValueChanged: " + ex.Message)
            End Try
        End If
        assignDBSheet.Enabled = False
        DBSheetCols.AutoResizeColumns()
        FormDisabled = False
    End Sub

    ''' <summary>reset ContextMenuStrip if outside cells</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DBSheetCols_MouseDown(sender As Object, e As MouseEventArgs) Handles DBSheetCols.MouseDown
        DBSheetCols.ContextMenuStrip = Nothing
    End Sub

    ''' <summary>store row index on key press/CellMouseDown to share with displayContextMenus, DBSheetColsForDatabases_ItemClicked, moveRow and RegenerateThisLookupQuery_Click</summary>
    Private selRowIndex As Integer
    ''' <summary>store column index on key press/CellMouseDown to share with displayContextMenus</summary>
    Private selColIndex As Integer

    ''' <summary>catch key presses: Shift F10 or menu key to get the context menu, Ctrl-C/Ctrl-V for copy/pasting foreign lookup info, DEL for clearing cells</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DBSheetCols_KeyDown(sender As Object, e As KeyEventArgs) Handles DBSheetCols.KeyDown
        selRowIndex = DBSheetCols.CurrentCell.RowIndex
        selColIndex = DBSheetCols.CurrentCell.ColumnIndex
        ' menu via menu key 
        ' only for whole rows ...
        If e.KeyCode = Keys.Apps Or (e.KeyCode = Keys.F10 And e.Modifiers = Keys.Shift) Then
            If DBSheetCols.SelectedRows.Count > 0 Then
                ' whole row selection -> move up/down menu...
                selColIndex = -1
                displayContextMenus()
                ' need show here as the context menu is not displayed otherwise...
                DBSheetCols.ContextMenuStrip.Show()
                ' ... or lookup column
            ElseIf selColIndex = 8 And selRowIndex >= 0 Then
                displayContextMenus()
                DBSheetCols.ContextMenuStrip.Show()
            End If
            ' Ctrl-C only when rows are available to copy
        ElseIf e.KeyCode = Keys.C And e.Modifiers = Keys.Control And DBSheetCols.DataSource.Rows.Count > 0 Then
            clipboardDataRow = DBSheetCols.DataSource.GetNewRow()
            clipboardDataRow.ItemArray = DBSheetCols.DataSource.Rows(selRowIndex).ItemArray.Clone()
            ' Ctrl-V only when clipboardDataRow has been copied
        ElseIf e.KeyCode = Keys.V And e.Modifiers = Keys.Control And clipboardDataRow IsNot Nothing Then
            FormDisabled = True
            DBSheetCols.Rows(selRowIndex).Cells("ftable").Value = clipboardDataRow.ftable
            DBSheetCols.Rows(selRowIndex).Cells("fkey").Value = clipboardDataRow.fkey
            DBSheetCols.Rows(selRowIndex).Cells("flookup").Value = clipboardDataRow.flookup
            DBSheetCols.Rows(selRowIndex).Cells("lookup").Value = clipboardDataRow.lookup
            DBSheetCols.Rows(selRowIndex).Cells("outer").Value = clipboardDataRow.outer
            DBSheetCols.Rows(selRowIndex).Cells("primkey").Value = clipboardDataRow.primkey
            ' fill in the dropdown values
            If DBSheetCols.Rows(selRowIndex).Cells("ftable").Value.ToString() = "" Then
                DirectCast(DBSheetCols.Rows(selRowIndex).Cells("fkey"), DataGridViewComboBoxCell).DataSource = New List(Of String)({""})
                DirectCast(DBSheetCols.Rows(selRowIndex).Cells("flookup"), DataGridViewComboBoxCell).DataSource = New List(Of String)({""})
            Else
                Dim forColsList As List(Of String) = getforeignTables(DBSheetCols.Rows(selRowIndex).Cells("ftable").Value.ToString())
                DirectCast(DBSheetCols.Rows(selRowIndex).Cells("fkey"), DataGridViewComboBoxCell).DataSource = forColsList
                DirectCast(DBSheetCols.Rows(selRowIndex).Cells("flookup"), DataGridViewComboBoxCell).DataSource = forColsList
            End If
            assignDBSheet.Enabled = False
            FormDisabled = False
            ' Delete key sets column values to empty
        ElseIf e.KeyCode = Keys.Delete Then
            If selRowIndex >= 0 Then
                ' avoid setting tick-boxes and type column to empty...
                If selColIndex <> 0 And Not (selColIndex >= 4 And selColIndex <= 6) Then DBSheetCols.Rows(selRowIndex).Cells().Item(selColIndex).Value = ""
                assignDBSheet.Enabled = False
            End If
            ' shortcut for move up
        ElseIf e.KeyCode = Keys.Up And DBSheetCols.SelectedRows.Count > 0 Then
            moveRow(-1)
            e.Handled = True
            ' shortcut for move down
        ElseIf e.KeyCode = Keys.Down And DBSheetCols.SelectedRows.Count > 0 Then
            moveRow(1)
            e.Handled = True
        End If
    End Sub

    ''' <summary>prepare context menus to be displayed after right mouse click</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DBSheetCols_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DBSheetCols.CellMouseDown
        selRowIndex = e.RowIndex
        selColIndex = e.ColumnIndex
        ' in case this was left somewhere, reset it here...
        FormDisabled = False
        If e.Button = Windows.Forms.MouseButtons.Right Then displayContextMenus()
    End Sub

    ''' <summary>display context menus depending on cell selected</summary>
    Private Sub displayContextMenus()
        If selColIndex = 8 And selRowIndex >= 0 Then
            DBSheetCols.ContextMenuStrip = DBSheetColsLookupMenu
        ElseIf selColIndex = 1 Then
            DBSheetColsForDatabases.Items.Clear()
            For Each entry As String In Database.Items
                DBSheetColsForDatabases.Items.Add(entry)
            Next
            DBSheetCols.ContextMenuStrip = DBSheetColsForDatabases
        ElseIf selColIndex = -1 And selRowIndex >= 0 AndAlso DBSheetCols.SelectedRows.Count > 0 AndAlso DBSheetCols.SelectedRows(0).Index = selRowIndex Then
            DBSheetCols.ContextMenuStrip = DBSheetColsMoveMenu
        Else
            DBSheetCols.ContextMenuStrip = Nothing
        End If
    End Sub

    ''' <summary>connect to the selected foreign database and get the tables into the ftable cell (!), the rest of the column still has the foreign tables of the main database.</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DBSheetColsForDatabases_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles DBSheetColsForDatabases.ItemClicked
        Try : myDBConnHelper.openConnection(e.ClickedItem.Text, usedForDBSheetCreate:=True) : Catch ex As Exception : UserMsg(ex.Message) : Exit Sub : End Try
        DirectCast(DBSheetCols.Rows(selRowIndex).Cells("ftable"), DataGridViewComboBoxCell).DataSource = getforeignTables()
        ' revert back to main database
        Try : myDBConnHelper.openConnection(Database.Text, usedForDBSheetCreate:=True) : Catch ex As Exception : UserMsg(ex.Message) : Exit Sub : End Try
    End Sub

    ''' <summary>move (shift) row up</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub MoveRowUp_Click(sender As Object, e As EventArgs) Handles MoveRowUp.Click
        moveRow(-1)
    End Sub

    ''' <summary>move (shift) row down</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub MoveRowDown_Click(sender As Object, e As EventArgs) Handles MoveRowDown.Click
        moveRow(1)
    End Sub

    ''' <summary>move row in direction given in param direction (-1: up, 1: down)</summary>
    ''' <param name="direction"></param>
    Private Sub moveRow(direction As Integer)
        Try
            ' avoid moving down of last row, DBSeqenceDataGrid.Rows.Count is 1 more than the actual inserted rows because of the "new" row, selIndex is 0 based. Also avoid moving up of first row
            If (selRowIndex = DBSheetCols.Rows.Count - 2 And direction = 1) Or (selRowIndex = 0 And direction = -1) Then Exit Sub
            If Not DBSheetCols.Rows(selRowIndex + 1).Cells("primkey").Value And DBSheetCols.Rows(selRowIndex).Cells("primkey").Value Then
                UserMsg("All primary keys have to be first and there is a NON-primary key column that would be shifted above this primary one !", "DBSheet Definition Error")
                Exit Sub
            End If
            If direction = -1 Then
                If (DBSheetCols.DataSource.Rows.Count - 1 < selRowIndex) Then
                    UserMsg("Editing not finished in selected row (values not committed), cannot move up!", "DBSheet Definition Error")
                    Exit Sub
                End If
                If DBSheetCols.Rows(selRowIndex - 1).Cells("primkey").Value And Not DBSheetCols.Rows(selRowIndex).Cells("primkey").Value Then
                    UserMsg("All primary keys have to be first and there is a primary key column that would be shifted below this NON-primary one !", "DBSheet Definition Error")
                    Exit Sub
                End If
            End If
            Dim rw As DBSheetDefRow = DBSheetCols.DataSource.GetNewRow()
            rw.ItemArray = DBSheetCols.DataSource.Rows(selRowIndex).ItemArray.Clone()
            Dim colnameList As List(Of String) = DirectCast(DBSheetCols.Rows(selRowIndex).Cells("fkey"), DataGridViewComboBoxCell).DataSource
            FormDisabled = True
            DBSheetCols.DataSource.Rows.RemoveAt(selRowIndex)
            DBSheetCols.DataSource.Rows.InsertAt(rw, selRowIndex + direction)
            DirectCast(DBSheetCols.Rows(selRowIndex + direction).Cells("flookup"), DataGridViewComboBoxCell).DataSource = colnameList
            DirectCast(DBSheetCols.Rows(selRowIndex + direction).Cells("fkey"), DataGridViewComboBoxCell).DataSource = colnameList
            DBSheetCols.CurrentCell = DBSheetCols.Rows(selRowIndex + direction).Cells(0)
            DBSheetCols.Rows(selRowIndex + direction).Selected = True
        Catch ex As System.Exception
            UserMsg("Exception in moveRow: " + ex.Message)
        End Try
        assignDBSheet.Enabled = False
        FormDisabled = False
    End Sub

    ''' <summary>(re)generates the lookup query for active row/cell</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub RegenerateThisLookupQuery_Click(sender As Object, e As EventArgs) Handles RegenerateThisLookupQuery.Click
        regenLookupForRow(selRowIndex)
        assignDBSheet.Enabled = False
        DBSheetCols.AutoResizeColumns()
    End Sub

    ''' <summary>(re)generate lookup for row in rowIndex</summary>
    ''' <param name="rowIndex"></param>
    Private Sub regenLookupForRow(rowIndex As Integer, Optional askForRegenerate As Boolean = False)
        If askForRegenerate Then
            Dim retval As MsgBoxResult = QuestionMsg("(re)generate foreign lookup, overwriting customizations?",, "DBSheet Definition")
            If retval = MsgBoxResult.Cancel Then Exit Sub
        End If
        If DBSheetCols.Rows(rowIndex).Cells("ftable").Value.ToString() <> "" And DBSheetCols.Rows(rowIndex).Cells("fkey").Value.ToString() <> "" And DBSheetCols.Rows(rowIndex).Cells("flookup").Value.ToString() <> "" Then
            DBSheetCols.Rows(rowIndex).Cells("lookup").Value = "SELECT " + tblPlaceHolder + "." + DBSheetCols.Rows(rowIndex).Cells("flookup").Value.ToString() + " " + correctNonNull(DBSheetCols.Rows(rowIndex).Cells("name").Value.ToString()) + "," + tblPlaceHolder + "." + DBSheetCols.Rows(rowIndex).Cells("fkey").Value.ToString() + " FROM " + DBSheetCols.Rows(rowIndex).Cells("ftable").Value.ToString() + " " + tblPlaceHolder + " ORDER BY " + DBSheetCols.Rows(rowIndex).Cells("flookup").Value.ToString()
            If askForRegenerate Then DBSheetCols.AutoResizeColumns()
        ElseIf DBSheetCols.Rows(rowIndex).Cells("ftable").Value.ToString() = "" And DBSheetCols.Rows(rowIndex).Cells("fkey").Value.ToString() = "" And DBSheetCols.Rows(rowIndex).Cells("flookup").Value.ToString() = "" Then
            Dim retval As MsgBoxResult = QuestionMsg("clear foreign lookup ?", MsgBoxStyle.OkCancel, "DBSheet Definition")
            If retval = MsgBoxResult.Ok Then DBSheetCols.Rows(rowIndex).Cells("lookup").Value = ""
        Else
            UserMsg("lookup query cannot be (re)generated as foreign table, key and lookup is not (fully) defined for field " + DBSheetCols.Rows(rowIndex).Cells("name").Value, "DBSheet Definition Error")
        End If
    End Sub

    ''' <summary>(re)generates ALL lookup queries</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub RegenerateAllLookupQueries_Click(sender As Object, e As EventArgs) Handles RegenerateAllLookupQueries.Click
        Dim retval As MsgBoxResult = QuestionMsg("regenerate foreign lookups completely, overwriting all customizations there: yes" + vbCrLf + "generate only new for empty lookups: no", MsgBoxStyle.YesNoCancel, "DBSheet Definition")
        If retval = MsgBoxResult.Cancel Then
            FormDisabled = False
            Exit Sub
        End If
        For i As Integer = 0 To DBSheetCols.Rows.Count - 2
            'only overwrite if forced regenerate or empty restriction def...
            If DBSheetCols.Rows(i).Cells("ftable").Value.ToString() <> "" Then
                If (retval = MsgBoxResult.Yes Or DBSheetCols.Rows(i).Cells("lookup").Value.ToString() = "") Then regenLookupForRow(i)
                ' remove lookup if forced regenerate and empty ftable...
            Else
                If retval = MsgBoxResult.Yes Then DBSheetCols.Rows(i).Cells("lookup").Value = ""
            End If
        Next
        assignDBSheet.Enabled = False
        DBSheetCols.AutoResizeColumns()
    End Sub

    ''' <summary>test the (generated or manually edited) lookup query in currently selected row</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TestLookupQuery_Click(sender As Object, e As EventArgs) Handles TestLookupQuery.Click
        If Strings.Len(DBSheetCols.Rows(selRowIndex).Cells("lookup").Value) > 0 Then
            testTheQuery(DBSheetCols.Rows(selRowIndex).Cells("lookup").Value, True)
        Else
            UserMsg("No restriction query created to test !!!", "DBSheet Definition Error")
        End If
    End Sub

    ''' <summary>removes the lookup query test currently open</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub RemoveLookupQueryTest_Click(sender As Object, e As EventArgs) Handles RemoveLookupQueryTest.Click
        If ExcelDnaUtil.Application.ActiveSheet.Name <> "TESTSHEET" Then
            UserMsg("Active sheet doesn't seem to be a query test sheet !!!", "DBSheet Definition Error")
        Else
            ExcelDnaUtil.Application.ActiveWorkbook.Close(False)
        End If
    End Sub

    ''' <summary>ignore data errors when loading data into grid-view</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DBSheetCols_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DBSheetCols.DataError
    End Sub

    ''' <summary>check for first Row/field if primary column field set</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DBSheetCols_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles DBSheetCols.RowsRemoved
        If FormDisabled Then Exit Sub
        If Not DBSheetCols.Rows(0).Cells("primkey").Value Then DBSheetCols.Rows(0).Cells("primkey").Value = True
    End Sub

#End Region

#Region "GUI element filling procedures"
    ''' <summary>add all fields of currently selected Table to DBSheetCols definitions</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub addAllFields_Click(ByVal sender As Object, ByVal e As EventArgs) Handles addAllFields.Click
        If DBSheetCols.Rows.Count > 1 Then
            Dim answr As MsgBoxResult = QuestionMsg("adding all fields resets current definitions, continue?",, "DBSheet Definition")
            If answr = MsgBoxResult.Cancel Then Exit Sub
        End If

        Try
            FormDisabled = True
            Dim firstRow As Boolean = True
            DBSheetCols.DataSource = Nothing
            DBSheetCols.Rows.Clear()
            Dim tableSchemaReader As OdbcDataReader
            Dim sqlCommand As New OdbcCommand() With {
                .CommandText = "SELECT TOP 1 * FROM " + Table.Text,
                .CommandType = CommandType.Text,
                .Connection = myDBConnHelper.dbshcnn
            }
            sqlCommand.Prepare()
            tableSchemaReader = sqlCommand.ExecuteReader(CommandBehavior.KeyInfo)
            Try
                Dim schemaInfo As DataTable = tableSchemaReader.GetSchemaTable()
                Dim theDBSheetDefTable = New DBSheetDefTable
                For Each schemaRow As DataRow In schemaInfo.Rows
                    Dim newRow As DBSheetDefRow = theDBSheetDefTable.GetNewRow()
                    newRow.name = If(schemaRow("AllowDBNull"), "", specialNonNullableChar) + schemaRow("ColumnName")
                    ' first field is always primary col by default, otherwise use IsKey property:
                    If firstRow Then
                        newRow.primkey = True
                        firstRow = False
                    Else
                        newRow.primkey = schemaRow("IsKey")
                    End If
                    ' having the specialNonNullableChar prepended to name is no problem as TableDataTypes respects this already!
                    newRow.type = TableDataTypes(newRow.name)
                    theDBSheetDefTable.Add(newRow)
                Next
                DBSheetCols.DataSource = theDBSheetDefTable
                DBSheetCols.AutoResizeColumns()
            Catch ex As Exception
                UserMsg("Could not get schema information for fields of table: '" + Table.Text + "', error: " + ex.Message, "DBSheet Definition Error")
            End Try
            tableSchemaReader.Close()
            FormDisabled = False
            ' after changing the columns no more change to table allowed !!
            TableEditable(False)
            EnvironEditable(False)
        Catch ex As System.Exception
            UserMsg("Exception in addAllFields_Click: " + ex.Message)
        End Try
    End Sub

    ''' <summary>clears the defined columns and resets the selection fields (Table, ForTable) and the Query</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub clearAllFields_Click(ByVal sender As Object, ByVal e As EventArgs) Handles clearAllFields.Click
        FormDisabled = True
        DBSheetCols.DataSource = Nothing
        DBSheetCols.Rows.Clear()
        TableEditable(True)
        EnvironEditable(True)
        FormDisabled = True
        Table.SelectedIndex = -1
        Query.Text = ""
        WhereClause.Text = ""
        ' reset the current filename
        currentFilepath = ""
        CurrentFileLinkLabel.Text = ""
        saveEnabled(False)
        DBSheetColsEditable(False)
        assignDBSheet.Enabled = False
        FormDisabled = False
    End Sub

    ''' <summary>mapping of column names (including null-able flag before) to their types (including size and precision)</summary>
    Private TableDataTypes As Dictionary(Of String, String)

    ''' <summary>gets the types of currently selected table including size, precision and scale into DataTypes</summary>
    Private Sub getTableDataTypes()
        TableDataTypes = New Dictionary(Of String, String)
        Dim rstSchema As OdbcDataReader
        Dim selectStmt As String = "SELECT TOP 1 * FROM " + Table.Text
        Dim sqlCommand As New OdbcCommand(selectStmt, myDBConnHelper.dbshcnn)
        rstSchema = sqlCommand.ExecuteReader()
        Try
            Dim schemaInfo As DataTable = rstSchema.GetSchemaTable()
            For Each schemaRow As DataRow In schemaInfo.Rows
                Dim appendInfo As String = If(schemaRow("AllowDBNull"), "", specialNonNullableChar)
                Dim precInfo As String = ""
                If schemaRow("DataType").Name <> "String" And schemaRow("DataType").Name <> "Boolean" Then
                    precInfo = "/" + schemaRow("NumericPrecision").ToString() + "/" + schemaRow("NumericScale").ToString()
                End If
                TableDataTypes(appendInfo + schemaRow("ColumnName")) = schemaRow("DataType").Name + "(" + schemaRow("ColumnSize").ToString() + precInfo + ")"
            Next
        Catch ex As Exception
            UserMsg("Could not get type information for table fields with query: '" + selectStmt + "', error: " + ex.Message, "DBSheet Definition Error")
        End Try
        rstSchema.Close()
    End Sub

    ''' <summary>gets the possible foreign tables in the instance (over all databases)</summary>
    ''' <param name="foreignTable"></param>
    ''' <returns>List of columns</returns>
    Private Function getforeignTables(foreignTable As String) As List(Of String)
        getforeignTables = New List(Of String)({""})
        Dim rstSchema As OdbcDataReader
        Dim selectStmt As String = "SELECT TOP 1 * FROM " + foreignTable
        Dim sqlCommand As New OdbcCommand(selectStmt, myDBConnHelper.dbshcnn)
        rstSchema = sqlCommand.ExecuteReader()
        Try
            Dim schemaInfo As DataTable = rstSchema.GetSchemaTable()
            For Each schemaRow As DataRow In schemaInfo.Rows
                getforeignTables.Add(schemaRow("ColumnName"))
            Next
        Catch ex As Exception
            UserMsg("Could not get type information for table fields with query: '" + selectStmt + "', error: " + ex.Message, "DBSheet Definition Error")
        End Try
        rstSchema.Close()
    End Function

    ''' <summary>fill all possible columns of currently selected table</summary>
    Private Function getColumns() As List(Of String)
        ' first get column/type Dictionary TableDataTypes
        getTableDataTypes()
        getColumns = New List(Of String)({""})
        For Each colname As String In TableDataTypes.Keys
            getColumns.Add(colname)
        Next
    End Function

    ''' <summary>fill all possible tables of currently selected Database</summary>
    Private Sub fillTables()
        Dim schemaTable As DataTable
        Dim tableTemp As String
        Try
            schemaTable = myDBConnHelper.dbshcnn.GetSchema("Tables")
            If schemaTable.Rows.Count = 0 Then Throw New Exception("No Tables could be fetched from Schema")
        Catch ex As Exception
            Throw New Exception("Error getting schema information for tables in connection strings database ' " + Database.Text + "'." + ",error: " + ex.Message)
        End Try
        FormDisabled = True
        tableTemp = Table.Text
        Table.Items.Clear()
        Try
            For Each iteration_row As DataRow In schemaTable.Rows
                If iteration_row("TABLE_CAT") = Database.Text Then Table.Items.Add(Database.Text + "." + iteration_row("TABLE_SCHEM") + "." + iteration_row("TABLE_NAME"))
            Next iteration_row
            If Strings.Len(tableTemp) > 0 Then Table.SelectedIndex = Table.Items.IndexOf(tableTemp)
            FormDisabled = False
        Catch ex As System.Exception
            FormDisabled = False
            Throw New Exception("Exception in fillTables: " + ex.Message)
        End Try
    End Sub

    ''' <summary>fill foreign tables into list of strings (is called by and filled in DBSheetCols.CellValueChanged)</summary>
    Private Function getforeignTables() As List(Of String)
        getforeignTables = New List(Of String)({""})
        Dim schemaTable As DataTable
        Try
            schemaTable = myDBConnHelper.dbshcnn.GetSchema("Tables")
            If schemaTable.Rows.Count = 0 Then Throw New Exception("No Tables could be fetched from Schema")
        Catch ex As Exception
            Throw New Exception("Error getting schema information for tables in connection strings database ' " + Database.Text + "'." + ",error: " + ex.Message)
        End Try
        Try
            For Each iteration_row As DataRow In schemaTable.Rows
                getforeignTables.Add(iteration_row("TABLE_CAT") + "." + iteration_row("TABLE_SCHEM") + "." + iteration_row("TABLE_NAME"))
            Next iteration_row
        Catch ex As System.Exception
            Throw New Exception("Exception in fillForTables: " + ex.Message)
        End Try
    End Function
#End Region

#Region "Query Creation and Testing"
    ''' <summary>create the final DBSheet Main Query</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub createQuery_Click(ByVal sender As Object, ByVal e As EventArgs) Handles createQuery.Click
        testQuery.Text = "&test DBSheet Query"
        If DBSheetCols.Rows.Count < 2 Then
            UserMsg("No columns defined yet, can't create query !", "DBSheet Definition Error")
            Exit Sub
        End If
        Dim retval As DialogResult = QuestionMsg("regenerate DBSheet query, overwriting all customizations there?",, "DBSheet Definition")
        If retval = MsgBoxResult.Cancel Then Exit Sub
        Dim queryStr As String = "", selectStr As String = "", orderByStr As String = ""
        Try
            Dim fromStr As String = "FROM " + Table.Text + " T1"
            Dim tableCounter As Integer = 1
            Dim selectPart As String = ""
            For i As Integer = 0 To DBSheetCols.Rows.Count - 2
                ' plain table field
                Dim usedColumn As String = correctNonNull(DBSheetCols.Rows(i).Cells("name").Value.ToString())
                ' used for foreign table lookups
                tableCounter += 1
                If DBSheetCols.Rows(i).Cells("sort").Value.ToString() <> "" Then
                    orderByStr = If(orderByStr = "", "", orderByStr + ", ") + (i + 1).ToString() + " " + DBSheetCols.Rows(i).Cells("sort").Value.ToString()
                End If
                Dim ftableStr As String = DBSheetCols.Rows(i).Cells("ftable").Value.ToString()
                If ftableStr = "" Then
                    selectStr += "T1." + usedColumn + ", "
                    ' create (inner or outer) joins for foreign key lookup id
                Else
                    Dim lookupStr As String = DBSheetCols.Rows(i).Cells("lookup").Value.ToString()
                    If lookupStr = "" Then
                        DBSheetCols.Rows(i).Selected = True
                        UserMsg("No lookup query created for field " + DBSheetCols.Rows(i).Cells("name").Value + ", can't proceed !", "DBSheet Definition Error")
                        Exit Sub
                    End If
                    Dim theTable As String = "T" + tableCounter.ToString()
                    ' either we go for the whole part after the last join
                    Dim completeJoin As String = fetchSubstr(lookupStr, "JOIN ", "")
                    ' or we have a simple WHERE and just "AND" it to the created join
                    Dim addRestrict As String = quotedReplace(fetchSubstr(lookupStr, "WHERE ", ""), "T" + tableCounter.ToString())

                    ' remove any ORDER BY clause from additional restrict...
                    Dim restrPos As Integer = addRestrict.ToUpper().LastIndexOf(" ORDER") + 1
                    If restrPos > 0 Then addRestrict = addRestrict.Substring(0, Math.Min(restrPos - 1, addRestrict.Length))
                    If completeJoin <> "" Then
                        ' when having the complete join, use additional restriction not for main subtable
                        addRestrict = ""
                        ' instead make it an additional condition for the join and replace placeholder with table-alias
                        completeJoin = quotedReplace(ciReplace(completeJoin, "WHERE", "AND"), "T" + tableCounter.ToString())
                    End If
                    Dim fkeyStr As String = DBSheetCols.Rows(i).Cells("fkey").Value.ToString()
                    If DBSheetCols.Rows(i).Cells("outer").Value Then
                        fromStr += " LEFT JOIN " + vbCrLf + ftableStr + " " + theTable +
                                       " ON " + "T1." + usedColumn + " = " + theTable + "." + fkeyStr + If(addRestrict <> "", " AND " + addRestrict, "")
                    Else
                        fromStr += " INNER JOIN " + vbCrLf + ftableStr + " " + theTable +
                                       " ON " + "T1." + usedColumn + " = " + theTable + "." + fkeyStr + If(addRestrict <> "", " AND " + addRestrict, "")
                    End If
                    ' we have additionally joined (an)other table(s) for the lookup display...
                    If completeJoin <> "" Then
                        ' remove any ORDER BY clause from completeJoin...
                        restrPos = completeJoin.ToUpper().LastIndexOf(" ORDER") + 1
                        If restrPos > 0 Then completeJoin = completeJoin.Substring(0, Math.Min(restrPos - 1, completeJoin.Length))
                        ' ..and add join of additional subtable(s) to the query
                        fromStr += " LEFT JOIN " + vbCrLf + completeJoin
                    End If

                    selectPart = fetchSubstr(lookupStr, "SELECT ", " FROM ").Trim()
                    ' remove second field in lookup query's select clause
                    restrPos = selectPart.LastIndexOf(",") + 1
                    selectPart = selectPart.Substring(0, Math.Min(restrPos - 1, selectPart.Length))
                    Dim aliasName As String = Strings.Mid(selectPart, InStrRev(selectPart, " ") + 1)
                    If aliasName <> usedColumn Then
                        UserMsg("Alias of lookup field '" + aliasName + "' is not consistent with field name '" + usedColumn + "', please change lookup definition !", "DBSheet Definition Error")
                        Exit Sub
                    End If
                    Dim flookupStr As String = DBSheetCols.Rows(i).Cells("flookup").Value.ToString()
                    ' customized select statement, take directly from lookup query..
                    If selectPart <> flookupStr Then
                        selectStr += quotedReplace(selectPart, "T" + tableCounter.ToString()) + ", "
                    Else
                        ' simple select statement (only the lookup field and id), put together...
                        selectStr += theTable + "." + flookupStr + " " + usedColumn + ", "
                    End If
                End If
            Next
            Dim wherePart As String = WhereClause.Text.Replace(vbCrLf, "")
            selectStr = "SELECT " + selectStr.Substring(0, Math.Min(Strings.Len(selectStr) - 2, selectStr.Length))
            queryStr = selectStr + vbCrLf + fromStr.ToString() + vbCrLf +
                     If(wherePart <> "", "WHERE " + wherePart + vbCrLf, "") +
                     If(orderByStr <> "", "ORDER BY " + orderByStr, "")
            saveEnabled(True)
        Catch ex As System.Exception
            UserMsg("Exception in createQuery_Click: " + ex.Message)
        End Try
        If queryStr <> "" Then Query.Text = queryStr
    End Sub

    ''' <summary>test the final DBSheet Main query or remove the test query sheet/workbook</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub testQuery_Click(ByVal sender As Object, ByVal e As EventArgs) Handles testQuery.Click
        Try
            If Query.Text <> "" Then
                If testQuery.Text = "&test DBSheet Query" Then
                    testTheQuery(Query.Text)
                ElseIf testQuery.Text = "&remove Test-sheet" Then
                    If ExcelDnaUtil.Application.ActiveSheet.Name <> "TESTSHEETQ" Then
                        UserMsg("Active sheet doesn't seem to be a query test sheet !", "DBSheet test-query Remove Warning", MessageBoxIcon.Exclamation)
                    Else
                        ExcelDnaUtil.Application.ActiveWorkbook.Close(False)
                    End If
                    testQuery.Text = "&test DBSheet Query"
                End If
            Else
                UserMsg("No Query created to test !!!", "DBSheet Query Test Warning", MessageBoxIcon.Exclamation)
            End If
        Catch ex As System.Exception
            UserMsg("Exception in testQuery_Click: " + ex.Message)
        End Try
    End Sub

    ''' <summary>for testing either the main query or the selected lookup query being given in theQueryText</summary>
    ''' <param name="theQueryText"></param>
    ''' <param name="isLookupQuery"></param>
    Private Sub testTheQuery(ByVal theQueryText As String, Optional ByRef isLookupQuery As Boolean = False)
        theQueryText = theQueryText.Replace(vbCrLf, " ").Replace(vbLf, " ")
        If isLookupQuery Then
            ' replace tblPlaceHolder with a more SQL compatible name..
            theQueryText = quotedReplace(theQueryText, "FT")
        ElseIf WhereClause.Text <> "" Then
            ' quoted replace of "?" with parameter values
            ' needs splitting of WhereClause by quotes !
            ' only for main query !!
            Dim whereClauseText As String = WhereClause.Text.Replace(vbCrLf, " ").Replace(vbLf, " ")
            Dim replacedStr As String = whereClauseText
            Dim j As Integer = 1
            While InStr(1, replacedStr, "?") > 0
                Dim questionMarkLoc As Integer = InStr(1, replacedStr, "?")
                Dim paramVal As String = InputBox("Value for parameter " + j.ToString() + " ?", "Enter parameter values..")
                If Len(paramVal) = 0 Then Exit Sub
                replacedStr = Strings.Mid(replacedStr, 1, questionMarkLoc - 1) + paramVal + Strings.Mid(replacedStr, questionMarkLoc + 1)
                j += 1
            End While
            If InStr(theQueryText, whereClauseText) = 0 Then
                UserMsg("Didn't find where clause " + whereClauseText + " in theQueryText: " + theQueryText + vbCrLf + "maybe creating DBSheet query again helps..", "DBSheet Definition Error")
                Exit Sub
            Else
                theQueryText = Replace(theQueryText, whereClauseText, replacedStr)
            End If
        End If
        Try
            ExcelDnaUtil.Application.SheetsInNewWorkbook = 1
            Dim newWB As Excel.Workbook = ExcelDnaUtil.Application.Workbooks.Add
            Dim Preview As Excel.Worksheet = newWB.Sheets(1)
            Preview.Cells(1, 2).Value = theQueryText
            Preview.Cells(1, 2).WrapText = False
            ' create a DBListFetch with the query 
            createFunctionsInCells(Preview.Cells(1, 1), {"RC", "=DBListFetch(RC[1], """", R[1]C,,,True)"})
            newWB.Saved = True
            If isLookupQuery Then
                Preview.Name = "TESTSHEET"
            Else
                Preview.Name = "TESTSHEETQ"
                testQuery.Text = "&remove Test-sheet"
            End If
            Exit Sub
        Catch ex As System.Exception
            UserMsg("Exception In testTheQuery: " + ex.Message)
        End Try
    End Sub
#End Region

#Region "DBSheet Definition files Handling"

    ''' <summary>direct assignment from Create Form...</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub assignDBSheet_Click(sender As Object, e As EventArgs) Handles assignDBSheet.Click
        DBSheetConfig.createDBSheet(currentFilepath)
    End Sub

    ''' <summary>loads the DBSheet definitions from a file (xml format)</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub loadDefs_Click(ByVal sender As Object, ByVal e As EventArgs) Handles loadDefs.Click
        Try
            Dim openFileDialog1 = New OpenFileDialog With {
                .InitialDirectory = fetchSetting("lastDBsheetCreatePath", fetchSetting("DBSheetDefinitions" + myDBConnHelper.DBenv, "")),
                .Filter = "XML files (*.xml)|*.xml",
                .RestoreDirectory = True
            }
            Dim loadOK As Boolean = True
            Dim result As DialogResult = openFileDialog1.ShowDialog()
            If result = Windows.Forms.DialogResult.OK Then
                ' remember path for possible storing in DBSheetParams
                currentFilepath = openFileDialog1.FileName
                If currentFilepath <> "" Then setUserSetting("lastDBsheetCreatePath", Strings.Left(currentFilepath, InStrRev(currentFilepath, "\") - 1))
                Dim DBSheetParams As String = File.ReadAllText(currentFilepath, System.Text.Encoding.Default)
                ' fetch parameters into form from sheet or file
                FormDisabled = True
                ' get Database from (legacy) connID (legacy connID was prefixed with connIDPrefixDBtype)
                Dim configDatabase As String = Replace(DBSheetConfig.getEntry("connID", DBSheetParams), fetchSetting("connIDPrefixDBtype", "MSSQL"), "")
                Try : myDBConnHelper.openConnection(configDatabase, usedForDBSheetCreate:=True) : Catch ex As Exception : UserMsg(ex.Message) : Exit Sub : End Try
                fillDatabases()
                Database.SelectedIndex = Database.Items.IndexOf(configDatabase)
                fillTables()
                DirectCast(DBSheetCols.Columns("ftable"), DataGridViewComboBoxColumn).DataSource = getforeignTables()
                FormDisabled = True
                Dim theTable As String = If(InStr(DBSheetConfig.getEntry("table", DBSheetParams), Database.Text + ".") > 0, DBSheetConfig.getEntry("table", DBSheetParams), Database.Text + fetchSetting("ownerQualifier" + env.ToString(), "") + DBSheetConfig.getEntry("table", DBSheetParams))
                Table.SelectedIndex = Table.Items.IndexOf(theTable)
                If Table.SelectedIndex = -1 Then
                    UserMsg("couldn't find table " + theTable + " defined in definitions file in database " + Database.Text + " in current environment (" + myDBConnHelper.DBenv + ") !", "DBSheet Definition Error")
                    FormDisabled = False
                    Exit Sub
                End If
                DirectCast(DBSheetCols.Columns("name"), DataGridViewComboBoxColumn).DataSource = getColumns()
                Dim columnslist As Object = DBSheetConfig.getEntryList("columns", "field", "", DBSheetParams)
                Dim theDBSheetDefTable = New DBSheetDefTable
                Dim fieldInfoMsg As String = ""
                For Each DBSheetColumnDef As String In columnslist
                    Dim newRow As DBSheetDefRow = theDBSheetDefTable.GetNewRow()
                    newRow.name = DBSheetConfig.getEntry("name", DBSheetColumnDef)
                    newRow.ftable = DBSheetConfig.getEntry("ftable", DBSheetColumnDef)
                    newRow.fkey = DBSheetConfig.getEntry("fkey", DBSheetColumnDef)
                    newRow.flookup = DBSheetConfig.getEntry("flookup", DBSheetColumnDef)
                    newRow.outer = DBSheetConfig.getEntry("outer", DBSheetColumnDef) <> ""
                    newRow.primkey = DBSheetConfig.getEntry("primkey", DBSheetColumnDef) <> ""
                    Dim noInfoForField As Boolean = False
                    If Not TableDataTypes.ContainsKey(newRow.name) Then
                        ' try to back up from changed non null-able info
                        If newRow.name.Substring(0, 1) = specialNonNullableChar Then
                            newRow.name = newRow.name.Substring(1)
                            If Not TableDataTypes.ContainsKey(newRow.name) Then
                                noInfoForField = True
                            Else
                                fieldInfoMsg += "Field " + specialNonNullableChar + newRow.name + " was wrong being non-null-able." + vbCrLf
                            End If
                        Else
                            newRow.name = specialNonNullableChar + newRow.name
                            If Not TableDataTypes.ContainsKey(newRow.name) Then
                                noInfoForField = True
                            Else
                                fieldInfoMsg += "Field " + newRow.name.Substring(1) + " was wrong being null-able." + vbCrLf
                            End If
                        End If
                    End If
                    If noInfoForField Then
                        UserMsg("couldn't retrieve information for field " + newRow.name + " in database, this field should be removed !", "DBSheet Definition Error")
                        loadOK = False
                        newRow.RowError = "couldn't retrieve information for field " + newRow.name + " in database, this field should be removed !"
                        newRow.type = ""
                    Else
                        newRow.type = TableDataTypes(newRow.name)
                        If newRow.type = "" Then
                            UserMsg("empty type information for field " + newRow.name + " in database !", "DBSheet Definition Error")
                            newRow.RowError = "empty type information for field " + newRow.name + " in database !"
                            loadOK = False
                        End If
                    End If
                    Dim sortMode As String = DBSheetConfig.getEntry("sort", DBSheetColumnDef)
                    ' legacy naming: Ascending/Descending
                    newRow.sort = If(sortMode = "Ascending", "ASC", If(sortMode = "Descending", "DESC", sortMode))
                    newRow.lookup = DBSheetConfig.getEntry("lookup", DBSheetColumnDef)
                    theDBSheetDefTable.Add(newRow)
                Next
                DBSheetCols.DataSource = theDBSheetDefTable
                ' re-add the fkey and flookup combo-box values...
                For i As Integer = 0 To DBSheetCols.Rows.Count - 2
                    If DBSheetCols.Rows(i).Cells("ftable").Value.ToString() <> "" Then
                        Dim colnameList As List(Of String) = getforeignTables(DBSheetCols.Rows(i).Cells("ftable").Value.ToString())
                        DirectCast(DBSheetCols.Rows(i).Cells("fkey"), DataGridViewComboBoxCell).DataSource = colnameList
                        DirectCast(DBSheetCols.Rows(i).Cells("flookup"), DataGridViewComboBoxCell).DataSource = colnameList
                    End If
                Next
                DBSheetCols.AutoResizeColumns()
                Query.Text = DBSheetConfig.getEntry("query", DBSheetParams)
                WhereClause.Text = DBSheetConfig.getEntry("whereClause", DBSheetParams)
                TableEditable(False)
                EnvironEditable(False)
                DBSheetColsEditable(True)
                saveEnabled(True)
                setLinkLabel(currentFilepath)
                ' pass info about wrong fields
                If fieldInfoMsg <> "" Then UserMsg(fieldInfoMsg, "DBSheet Definition Problem", MsgBoxStyle.Exclamation)
                If loadOK Then assignDBSheet.Enabled = True
            End If
        Catch ex As System.Exception
            UserMsg("Exception in loadDefs_Click: " + ex.Message)
        End Try
    End Sub

    ''' <summary>save definitions button</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub saveDefs_Click(ByVal sender As Object, ByVal e As EventArgs) Handles saveDefs.Click
        saveDefinitionsToFile(False)
    End Sub

    ''' <summary>save definitions as button</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub saveDefsAs_Click(ByVal sender As Object, ByVal e As EventArgs) Handles saveDefsAs.Click
        saveDefinitionsToFile(True)
    End Sub

    ''' <summary>the currently set file path of the DBSheet definition file</summary>
    Private currentFilepath As String = ""

    ''' <summary>saves the definitions currently stored in theDBSheetCreateForm to newly selected file (saveAs = True) or to the file already stored in setting "dsdPath"</summary>
    ''' <param name="saveAs"></param>
    Private Sub saveDefinitionsToFile(ByRef saveAs As Boolean)
        Try
            If saveAs Or currentFilepath = "" Then
                Dim invalidPath As Boolean = True
                ' loop until a valid path with write access is chosen (or cancel exits sub)
                While invalidPath
                    Dim tableName As String = If(InStrRev(Table.Text, ".") > 0, Strings.Mid(Table.Text, InStrRev(Table.Text, ".") + 1), Table.Text)
                    Dim saveFileDialog1 As New SaveFileDialog With {
                        .Title = "Save DBSheet Definition",
                        .FileName = tableName + ".xml",
                        .InitialDirectory = fetchSetting("lastDBsheetCreatePath", fetchSetting("DBSheetDefinitions" + myDBConnHelper.DBenv, "")),
                        .Filter = "XML files (*.xml)|*.xml",
                        .RestoreDirectory = True
                    }
                    Dim result As DialogResult = saveFileDialog1.ShowDialog()
                    If result = Windows.Forms.DialogResult.OK Then
                        currentFilepath = saveFileDialog1.FileName
                        Try
                            FileSystem.FileOpen(1, currentFilepath, OpenMode.Output)
                            FileSystem.PrintLine(1, xmlDbsheetConfig())
                        Catch ex As Exception
                            UserMsg("Can't save to folder " + currentFilepath + ", please choose another!")
                            Continue While
                        End Try
                        FileSystem.FileClose(1)
                        setUserSetting("lastDBsheetCreatePath", Strings.Left(currentFilepath, InStrRev(currentFilepath, "\") - 1))
                        setLinkLabel(currentFilepath)
                        invalidPath = False
                    Else
                        Exit Sub
                    End If
                End While
            Else
                FileSystem.FileOpen(1, currentFilepath, OpenMode.Output)
                FileSystem.PrintLine(1, xmlDbsheetConfig())
                FileSystem.FileClose(1)
            End If
            assignDBSheet.Enabled = True
        Catch ex As System.Exception
            UserMsg("Exception in saveDefinitionsToFile: " + ex.Message)
        End Try
    End Sub

    ''' <summary>used to display the full path of the DBSheet definition filename</summary>
    Private linklabelToolTip As System.Windows.Forms.ToolTip

    ''' <summary>sets current definition file path hyperlink label. Displayed is only the filename, full path is stored in tag and visible in tool-tip</summary>
    ''' <param name="filepath">definition file path</param>
    Private Sub setLinkLabel(filepath As String)
        CurrentFileLinkLabel.Text = Strings.Mid(filepath, InStrRev(filepath, "\") + 1)
        CurrentFileLinkLabel.Tag = filepath
        If IsNothing(linklabelToolTip) Then linklabelToolTip = New System.Windows.Forms.ToolTip()
        linklabelToolTip.SetToolTip(CurrentFileLinkLabel, filepath)
    End Sub

    ''' <summary>creates xml DBsheet parameter string from the data entered in theDBSheetCreateForm</summary>
    ''' <returns>the xml DBsheet parameter string</returns>
    Private Function xmlDbsheetConfig() As String
        Dim namedParams As String = "", columnsDef As String = ""
        Try
            ' first create the columns list
            Dim primKeyCount As Integer = 0
            ' collect lookups
            For i As Integer = 0 To DBSheetCols.RowCount - 2 ' respect the insert row !!!
                Dim columnLine As String = "<field>"
                For j As Integer = 0 To DBSheetCols.ColumnCount - 1
                    If DBSheetCols.Rows(i).Cells(j).Value.ToString() <> "" Then
                        ' store everything false values and type (is always inferred from Database)
                        If Not (DBSheetCols.Columns(j).Name = "type" OrElse
                            (TypeName(DBSheetCols.Rows(i).Cells(j).Value) = "Boolean" AndAlso Not DBSheetCols.Rows(i).Cells(j).Value)) Then
                            columnLine += DBSheetConfig.setEntry(DBSheetCols.Columns(j).Name, CStr(DBSheetCols.Rows(i).Cells(j).Value.ToString()))
                        End If
                    End If
                Next
                columnsDef += vbCrLf + columnLine + "</field>"
                If DBSheetCols.Rows(i).Cells("primkey").Value Then primKeyCount += 1
            Next
            ' then create the parameters stored in named cells
            namedParams += DBSheetConfig.setEntry("connID", Database.Text) + vbCrLf
            namedParams += DBSheetConfig.setEntry("table", Table.Text) + vbCrLf
            namedParams += DBSheetConfig.setEntry("query", Query.Text) + vbCrLf
            namedParams += DBSheetConfig.setEntry("whereClause", WhereClause.Text) + vbCrLf
            namedParams += DBSheetConfig.setEntry("primcols", primKeyCount.ToString())
            ' finally put everything together:
            Return "<DBsheetConfig>" + vbCrLf + namedParams + vbCrLf + "<columns>" + columnsDef + vbCrLf + "</columns>" + vbCrLf + "</DBsheetConfig>"
        Catch ex As System.Exception
            UserMsg("Exception in xmlDbsheetConfig: " + ex.Message)
            Return ""
        End Try
    End Function

    ''' <summary>current file link clicked: open file using default editor</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CurrentFileLinkLabel_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles CurrentFileLinkLabel.LinkClicked
        ' CurrentFileLinkLabel.Tag contains the path to the current file
        Diagnostics.Process.Start(CurrentFileLinkLabel.Tag)
    End Sub

#End Region

#Region "GUI Helper functions"
    ''' <summary>Table/Database/Password change possible</summary>
    ''' <param name="choice">possible True, not possible False</param>
    Private Sub TableEditable(choice As Boolean)
        Database.Enabled = choice
        LDatabase.Enabled = choice
        Table.Enabled = choice
        LTable.Enabled = choice
    End Sub

    ''' <summary>Environment change possible</summary>
    ''' <param name="choice">possible True, not possible False</param>
    Private Sub EnvironEditable(choice As Boolean)
        Environment.Enabled = choice
        Lenvironment.Enabled = choice
    End Sub
    ''' <summary>if DBSheetCols Definitions should be editable, enable relevant controls</summary>
    ''' <param name="choice">editable True, not editable False</param>
    Private Sub DBSheetColsEditable(choice As Boolean)
        ' block password change if DBSheetDef is editable or no password needed
        If InStr(1, myDBConnHelper.dbsheetConnString, myDBConnHelper.dbPwdSpec) > 0 And myDBConnHelper.dbPwdSpec <> "" Then
            Password.Enabled = Not choice
            LPwd.Enabled = Not choice
        Else
            Password.Enabled = False
            LPwd.Enabled = False
        End If
        addAllFields.Enabled = choice
        clearAllFields.Enabled = choice
        DBSheetCols.Enabled = choice
        createQuery.Enabled = choice
        testQuery.Enabled = choice
        Query.Enabled = choice
        WhereClause.Enabled = choice
    End Sub

    ''' <summary>toggle saveEnabled behaviour</summary>
    ''' <param name="choice"></param>
    Private Sub saveEnabled(ByRef choice As Boolean)
        saveDefs.Enabled = choice
        saveDefsAs.Enabled = choice
    End Sub
#End Region

#Region "Various Helper functions"
    ''' <summary>corrects field names of non null-able fields prepended with specialNonNullableChar (e.g. "*") back to the real name</summary>
    ''' <param name="name"></param>
    ''' <returns>the corrected string</returns>
    Public Function correctNonNull(name As String) As String
        correctNonNull = If(Strings.Left(name, 1) = specialNonNullableChar, Strings.Right(name, Len(name) - 1), name)
    End Function

    ''' <summary>replaces keystr with changed in theString, case insensitive !!</summary>
    ''' <param name="theString"></param>
    ''' <param name="keystr"></param>
    ''' <param name="changed"></param>
    ''' <returns>modified String</returns>
    Private Function ciReplace(ByVal theString As String, ByVal keystr As String, ByVal changed As String) As String
        Replace(theString, keystr, changed)
        Dim replaceBeg As Integer = InStr(1, theString.ToUpper(), keystr.ToUpper())
        If replaceBeg = 0 Then
            Return theString
        End If
        Return Strings.Left(theString, replaceBeg - 1) + changed + Strings.Right(theString, Len(theString) - replaceBeg - Len(keystr) + 1)
    End Function

    ''' <summary>replaces tblPlaceHolder with changed in theString, quote aware (keystr is not replaced within quotes) !!</summary>
    ''' <param name="theString"></param>
    ''' <param name="changed"></param>
    ''' <returns>the replaced string</returns>
    Private Function quotedReplace(ByVal theString As String, ByVal changed As String) As String
        Dim subresult As String
        quotedReplace = ""
        Dim teststr As String() = Split(theString, "'")
        ' walk through quote1 splitted parts and replace keystr in even ones
        For i As Integer = 0 To UBound(teststr)
            If i Mod 2 = 0 Then
                subresult = Replace(teststr(i), tblPlaceHolder, changed)
            Else
                subresult = teststr(i).ToString()
            End If
            quotedReplace += subresult + IIf(i < UBound(teststr), "'", "")
        Next
    End Function

    ''' <summary>block assignment possibility of DBSheet after query has been changed (needs to be saved first)</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Query_TextChanged(sender As Object, e As EventArgs) Handles Query.TextChanged
        assignDBSheet.Enabled = False
    End Sub

    ''' <summary>block assignment possibility of DBSheet after query has been changed (needs to be saved first)</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub WhereClause_TextChanged(sender As Object, e As EventArgs) Handles WhereClause.TextChanged
        assignDBSheet.Enabled = False
    End Sub

    ''' <summary>set selected environment (global) to set environment, reflect in ribbon and "restart" DBSheetCreateForm to get necessary information</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Environment_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Environment.SelectedIndexChanged
        If FormDisabled Then Exit Sub
        SettingsTools.selectedEnvironment = Me.Environment.SelectedIndex
        theRibbon.InvalidateControl("envDropDown")
        finalizeSetup()
    End Sub

#End Region
End Class

#Region "DBSheetCols Grid-view Helper Classes"
''' <summary>DataTable Class for filling DBSheetCols DataGridView</summary>
Public Class DBSheetDefTable : Inherits DataTable

    ''' <summary>returns one selected item from the DBSheetDefRow item collection</summary>
    ''' <param name="idx">selected index</param>
    ''' <returns>the selected DBSheetDefRow</returns>
    Default Public ReadOnly Property Item(ByVal idx As Integer) As DBSheetDefRow
        Get
            Return CType(Rows(idx), DBSheetDefRow)
        End Get
    End Property

    Public Sub New()
        Columns.Add(New DataColumn("name", GetType(String)))
        Columns.Add(New DataColumn("ftable", GetType(String)))
        Columns.Add(New DataColumn("fkey", GetType(String)))
        Columns.Add(New DataColumn("flookup", GetType(String)))
        Columns.Add(New DataColumn("outer", GetType(Boolean)))
        Columns.Add(New DataColumn("primkey", GetType(Boolean)))
        Columns.Add(New DataColumn("type", GetType(String)))
        Columns.Add(New DataColumn("sort", GetType(String)))
        Columns.Add(New DataColumn("lookup", GetType(String)))
    End Sub

    ''' <summary>stub for adding a row</summary>
    ''' <param name="row"></param>
    Public Sub Add(ByVal row As DBSheetDefRow)
        Rows.Add(row)
    End Sub

    ''' <summary>stub for removing a row</summary>
    ''' <param name="row"></param>
    Public Sub Remove(ByVal row As DBSheetDefRow)
        Rows.Remove(row)
    End Sub

    ''' <summary>get a new row for DBSheetDefRow type</summary>
    ''' <returns>the new row</returns>
    Public Function GetNewRow() As DBSheetDefRow
        Dim row As DBSheetDefRow = CType(NewRow(), DBSheetDefRow)
        Return row
    End Function

    ''' <summary>get the allowed type for a row (= DBSheetDefRow)</summary>
    ''' <returns></returns>
    Protected Overrides Function GetRowType() As Type
        Return GetType(DBSheetDefRow)
    End Function

    ''' <summary>override to get a new row from the DataRowBuilder</summary>
    ''' <param name="builder"></param>
    ''' <returns>the new row</returns>
    Protected Overrides Function NewRowFromBuilder(ByVal builder As DataRowBuilder) As DataRow
        Return New DBSheetDefRow(builder)
    End Function

End Class

''' <summary>DataRow Class for DBSheetDefTable</summary>
Public Class DBSheetDefRow : Inherits DataRow
    ''' <summary>accessor for the name property</summary>
    ''' <returns></returns>
    Public Property name As String
        Get
            Return CStr(MyBase.Item("name"))
        End Get
        Set(ByVal value As String)
            MyBase.Item("name") = value
        End Set
    End Property
    ''' <summary>accessor for the ftable property</summary>
    ''' <returns></returns>
    Public Property ftable As String
        Get
            Return CStr(MyBase.Item("ftable"))
        End Get
        Set(ByVal value As String)
            MyBase.Item("ftable") = value
        End Set
    End Property
    ''' <summary>accessor for the name property</summary>
    ''' <returns></returns>
    Public Property fkey As String
        Get
            Return CStr(MyBase.Item("fkey"))
        End Get
        Set(ByVal value As String)
            MyBase.Item("fkey") = value
        End Set
    End Property
    ''' <summary>accessor for the flookup property</summary>
    ''' <returns></returns>
    Public Property flookup As String
        Get
            Return CStr(MyBase.Item("flookup"))
        End Get
        Set(ByVal value As String)
            MyBase.Item("flookup") = value
        End Set
    End Property
    ''' <summary>accessor for the outer property</summary>
    ''' <returns></returns>
    Public Property outer As Boolean
        Get
            Return CBool(MyBase.Item("outer"))
        End Get
        Set(ByVal value As Boolean)
            MyBase.Item("outer") = value
        End Set
    End Property
    ''' <summary>accessor for the primkey property</summary>
    ''' <returns></returns>
    Public Property primkey As Boolean
        Get
            Return CBool(MyBase.Item("primkey"))
        End Get
        Set(ByVal value As Boolean)
            MyBase.Item("primkey") = value
        End Set
    End Property
    ''' <summary>accessor for the type property</summary>
    ''' <returns></returns>
    Public Property type As String
        Get
            Return CStr(MyBase.Item("type"))
        End Get
        Set(ByVal value As String)
            MyBase.Item("type") = value
        End Set
    End Property
    ''' <summary>accessor for the sort property</summary>
    ''' <returns></returns>
    Public Property sort As String
        Get
            Return CStr(MyBase.Item("sort"))
        End Get
        Set(ByVal value As String)
            MyBase.Item("sort") = value
        End Set
    End Property
    ''' <summary>accessor for the lookup property</summary>
    ''' <returns></returns>
    Public Property lookup As String
        Get
            Return CStr(MyBase.Item("lookup"))
        End Get
        Set(ByVal value As String)
            MyBase.Item("lookup") = value
        End Set
    End Property
    ''' <summary>constructor for a DBSheetDefRow</summary>
    ''' <param name="builder"></param>
    Friend Sub New(ByVal builder As DataRowBuilder)
        MyBase.New(builder)
        name = ""
        ftable = ""
        fkey = ""
        flookup = ""
        outer = False
        primkey = False
        type = ""
        sort = ""
        lookup = ""
    End Sub
End Class
#End Region