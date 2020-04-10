Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.Odbc
Imports System.IO
Imports System.Windows.Forms

''' <summary>Form for defining/creating DBSheets</summary>
Public Class DBSheetCreateForm
    Inherits System.Windows.Forms.Form
    ''' <summary>whether the form fields should react to changes (set if making changes within code)....</summary>
    Private FormDisabled As Boolean
    ''' <summary>the connection string for dbsheet definitions, different from the normal one (extended rights for schema viewing required)</summary>
    Private dbsheetConnString As String
    ''' <summary>identifier needed to fetch database from connection string (eg Database=)</summary>
    Private dbidentifier As String
    ''' <summary>statement/procedure to get all databases in a DB instance</summary>
    Private dbGetAllStr As String
    ''' <summary>fieldname where databases are returned by dbGetAllStr</summary>
    Private DBGetAllFieldName As String
    ''' <summary>the DB connection for the dbsheet definition activities</summary>
    Private dbshcnn As OdbcConnection
    ''' <summary>identifier needed to put password into connection string (eg PWD=)</summary>
    Private dbPwdSpec As String
    Private tblPlaceHolder As String = "!T!"
    Private specialNonNullableChar As String = "*"

#Region "Initialization of DBSheetDefs"
    ''' <summary>entry point of form, invoked by clicking "create/edit DBSheet definition"</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DBSheetCreateForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        ' get settings for DBSheet definition editing
        Dim env As Integer = Globals.selectedEnvironment + 1
        dbGetAllStr = fetchSetting("dbGetAll" + env.ToString, "NONEXISTENT")
        If dbGetAllStr = "NONEXISTENT" Then
            ErrorMsg("No dbGetAllStr given for environment: " + env.ToString + ", please correct and rerun.", "createDefinitions Error")
            Exit Sub
        End If
        DBGetAllFieldName = fetchSetting("dbGetAllFieldName" + env.ToString, "NONEXISTENT")
        If DBGetAllFieldName = "NONEXISTENT" Then
            ErrorMsg("No DBGetAllFieldName given for environment: " + env.ToString + ", please correct and rerun.", "createDefinitions Error")
            Exit Sub
        End If
        dbsheetConnString = fetchSetting("DBSheetConnString" + env.ToString, "NONEXISTENT")
        If dbsheetConnString = "NONEXISTENT" Then
            ErrorMsg("No Connectionstring given for environment: " + env.ToString + ", please correct and rerun.", "createDefinitions Error")
            Exit Sub
        End If
        dbidentifier = fetchSetting("DBidentifierCCS" + env.ToString, "NONEXISTENT")
        If dbidentifier = "NONEXISTENT" Then
            ErrorMsg("No DB identifier given for environment: " + env.ToString + ", please correct and rerun.", "createDefinitions Error")
            Exit Sub
        End If
        dbPwdSpec = fetchSetting("dbPwdSpec" + env.ToString, "")
        tblPlaceHolder = fetchSetting("tblPlaceHolder" + env.ToString, "!T!")
        specialNonNullableChar = fetchSetting("specialNonNullableChar" + env.ToString, "*")

        ' set up columns for DBSheetCols gridview
        Dim nameCB As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn With {
                    .Name = "name",
                    .DataSource = New List(Of String),
                    .HeaderText = "name",
                    .DataPropertyName = "name",
                    .SortMode = DataGridViewColumnSortMode.NotSortable
                }
        Dim ftableCB As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn With {
                    .Name = "ftable",
                    .DataSource = New List(Of String),
                    .HeaderText = "ftable",
                    .DataPropertyName = "ftable",
                    .SortMode = DataGridViewColumnSortMode.NotSortable
                }
        Dim fkeyCB As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn With {
                    .Name = "fkey",
                    .DataSource = New List(Of String),
                    .HeaderText = "fkey",
                    .DataPropertyName = "fkey",
                    .SortMode = DataGridViewColumnSortMode.NotSortable
                }
        Dim flookupCB As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn With {
                    .Name = "flookup",
                    .DataSource = New List(Of String),
                    .HeaderText = "flookup",
                    .DataPropertyName = "flookup",
                    .SortMode = DataGridViewColumnSortMode.NotSortable
                }
        Dim outerCB As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn With {
                    .Name = "outer",
                    .HeaderText = "outer",
                    .DataPropertyName = "outer",
                    .SortMode = DataGridViewColumnSortMode.NotSortable
                }
        Dim primkeyCB As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn With {
                    .Name = "primkey",
                    .HeaderText = "primkey",
                    .DataPropertyName = "primkey",
                    .SortMode = DataGridViewColumnSortMode.NotSortable
                }
        Dim typeTB As DataGridViewTextBoxColumn = New DataGridViewTextBoxColumn With {
                    .Name = "type",
                    .HeaderText = "type",
                    .DataPropertyName = "type",
                    .[ReadOnly] = True,
                    .SortMode = DataGridViewColumnSortMode.NotSortable
                }
        Dim sortCB As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn With {
                    .Name = "sort",
                    .DataSource = New List(Of String)({"", "ASC", "DESC"}),
                    .HeaderText = "sort",
                    .DataPropertyName = "sort",
                    .SortMode = DataGridViewColumnSortMode.NotSortable
                }
        Dim lookupTB As DataGridViewTextBoxColumn = New DataGridViewTextBoxColumn With {
                    .Name = "lookup",
                    .HeaderText = "lookup",
                    .DataPropertyName = "lookup",
                    .SortMode = DataGridViewColumnSortMode.NotSortable
                }
        DBSheetCols.AutoGenerateColumns = False
        DBSheetCols.Columns.AddRange(nameCB, ftableCB, fkeyCB, flookupCB, outerCB, primkeyCB, typeTB, sortCB, lookupTB)
        DBSheetColsEditable(False)

        ' if we have a Password to enter, just display explanation text in title bar and let user enter password... 
        If dbPwdSpec <> "" And existingPwd = "" Then
            Me.Text = "DB Sheet creation: Please enter required Password into Pwd to access schema information"
            TableEditable(False)
            saveEnabled(False)
        Else ' otherwise jump in immediately
            Password.Text = existingPwd
            fillDatabasesAndSetDropDown()
            ' initialize with empty DBSheet definitions is done by above call, changing Database.SelectedIndex (Database_SelectedIndexChanged)
        End If
    End Sub

    ''' <summary>enter pressed in Password textbox triggering initialisation</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Password_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Password.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then setPasswordAndInit()
    End Sub

    ''' <summary>leaving Password textbox triggering initialisation</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Password_Leave(sender As Object, e As EventArgs) Handles Password.Leave
        If existingPwd <> Password.Text Then setPasswordAndInit()
    End Sub

    ''' <summary>called after new password has been entered, reset stored password and fill database dropdown</summary>
    Private Sub setPasswordAndInit()
        existingPwd = Password.Text
        Try : dbshcnn.Close() : Catch ex As Exception : End Try
        dbshcnn = Nothing
        fillDatabasesAndSetDropDown()
    End Sub

    ''' <summary>fill the Database dropdown and set dropdown to database set in connection string</summary>
    Private Sub fillDatabasesAndSetDropDown()
        Try
            fillDatabases(Database)
        Catch ex As System.Exception
            TableEditable(False)
            saveEnabled(False)
            DBSheetColsEditable(False)
            ErrorMsg(ex.Message)
            Exit Sub
        End Try
        Me.Text = "DB Sheet creation: Select Database and Table to start building a DBSheet Definition"
        Database.SelectedIndex = Database.Items.IndexOf(fetch(dbsheetConnString, dbidentifier, ";"))
        'initialization of everything else is triggered by above change and caught by Database_SelectedIndexChanged
    End Sub

    ''' <summary>fills all possible databases of current connection using db proprietary code in dbGetAllStr, data coming from field DBGetAllFieldName</summary>
    Private Sub fillDatabases(DatabaseComboBox As ComboBox)
        Dim addVal As String
        Dim dbs As OdbcDataReader

        ' do not catch exception here, as it should be handled by fillDatabasesAndSetDropDown
        openConnection()
        FormDisabled = True
        DatabaseComboBox.Items.Clear()
        Dim sqlCommand As OdbcCommand = New OdbcCommand(dbGetAllStr, dbshcnn)
        Try
            dbs = sqlCommand.ExecuteReader()
        Catch ex As OdbcException
            FormDisabled = False
            Throw New Exception("Could not retrieve schema information for databases in connection string: '" + dbsheetConnString + "',error: " + ex.Message)
        End Try
        If dbs.HasRows Then
            Try
                Do
                    If Strings.Len(DBGetAllFieldName) = 0 Then
                        addVal = dbs(0)
                    Else
                        addVal = dbs(DBGetAllFieldName)
                    End If
                    DatabaseComboBox.Items.Add(addVal)
                Loop While dbs.Read()
                dbs.Close()
                FormDisabled = False
            Catch ex As System.Exception
                FormDisabled = False
                Throw New Exception("Exception when filling DatabaseComboBox: " + ex.Message)
            End Try
        Else
            FormDisabled = False
            Throw New Exception("Could not retrieve any databases with: " + dbGetAllStr + "!")
        End If
    End Sub

    ''' <summary>database changed, initialize everything else (Tables, DBSheetCols definition) from scratch</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Database_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Database.SelectedIndexChanged
        Try : openConnection() : Catch ex As Exception : ErrorMsg(ex.Message) : Exit Sub : End Try
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
            ErrorMsg("Exception in Database_SelectedIndexChanged: " + ex.Message)
        End Try
    End Sub

    ''' <summary>selecting the Table triggers enabling the DBSheetCols definition (fils columns/fields of that table and resetting DBSheetCols definition)</summary>
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
            ErrorMsg("Exception in Table_SelectedIndexChanged: " + ex.Message)
        End Try
        Me.Text = "DB Sheet creation: Select one or more columns (fields) adding possible foreign key lookup information in foreign tables, finally click create query to finish DBSheet definition"
    End Sub
#End Region

#Region "DBSheetCols Gridview"
    ''' <summary>handles the various changes in the DBSheetCols gridview</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DBSheetCols_CellValueChanged(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles DBSheetCols.CellValueChanged
        If FormDisabled Then Exit Sub
        FormDisabled = True
        Dim selIndex As Integer = e.RowIndex
        If e.ColumnIndex = 0 Then     ' field name column 
            ' lock table and database choice to not let user accidentally clear DBSheet defs
            TableEditable(False)
            ' fill empty lists for fkey and flookup comboboxes (resetting) ..
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
            If DBSheetCols.Rows(selIndex).Cells("ftable").Value.ToString = "" Then ' reset fkey and flookup dropdown
                DirectCast(DBSheetCols.Rows(selIndex).Cells("fkey"), DataGridViewComboBoxCell).DataSource = New List(Of String)({""})
                DirectCast(DBSheetCols.Rows(selIndex).Cells("flookup"), DataGridViewComboBoxCell).DataSource = New List(Of String)({""})
            Else ' fill fkey and flookup with fields of ftable
                Dim forColsList As List(Of String) = getforeignTables(DBSheetCols.Rows(selIndex).Cells("ftable").Value.ToString)
                DirectCast(DBSheetCols.Rows(selIndex).Cells("fkey"), DataGridViewComboBoxCell).DataSource = forColsList
                DirectCast(DBSheetCols.Rows(selIndex).Cells("flookup"), DataGridViewComboBoxCell).DataSource = forColsList
            End If
            ' reset fkey and flookup
            DBSheetCols.Rows(selIndex).Cells("fkey").Value = ""
            DBSheetCols.Rows(selIndex).Cells("flookup").Value = ""
        ElseIf e.ColumnIndex = 3 Then  ' flookup column -> ask for regeneration of flookup
            Dim retval As MsgBoxResult = QuestionMsg("regenerate foreign lookup (overwriting all customizations there)?",, "DBSheet Definition")
            If retval <> MsgBoxResult.Cancel Then regenLookupForRow(selIndex)
            DBSheetCols.AutoResizeColumns()
        ElseIf e.ColumnIndex = 5 Then ' primkey column
        Try
                ' not first row selected: check for previous row (field) if also primary column..
                If Not selIndex = 0 Then
                    If Not DBSheetCols.Rows(selIndex - 1).Cells("primkey").Value And DBSheetCols.Rows(selIndex).Cells("primkey").Value Then
                        ErrorMsg("All primary keys have to be first and there is at least one non-primary key column before that one !", "DBSheet Definition Warning")
                        DBSheetCols.Rows(selIndex).Cells("primkey").Value = False
                    End If
                    ' check if next row (field) is primary key column (only for non-last rows)
                    If selIndex <> DBSheetCols.Rows.Count - 2 Then
                        If DBSheetCols.Rows(selIndex + 1).Cells("primkey").Value And Not DBSheetCols.Rows(selIndex).Cells("primkey").Value Then
                            ErrorMsg("All primary keys have to be first and there is at least one primary key column after that one !", "DBSheet Definition Warning")
                            DBSheetCols.Rows(selIndex).Cells("primkey").Value = True
                        End If
                    End If
                ElseIf Not DBSheetCols.Rows(selIndex).Cells("primkey").Value Then
                    ErrorMsg("first column always has to be primary key", "DBSheet Definition Warning")
                    DBSheetCols.Rows(selIndex).Cells("primkey").Value = True
                End If
            Catch ex As System.Exception
                ErrorMsg("Exception in DBSheetCols_CellValueChanged: " + ex.Message)
            End Try
        End If
        DBSheetCols.AutoResizeColumns()
        FormDisabled = False
    End Sub

    ''' <summary>reset ContextMenuStrip if outside cells</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DBSheetCols_MouseDown(sender As Object, e As MouseEventArgs) Handles DBSheetCols.MouseDown
        DBSheetCols.ContextMenuStrip = Nothing
    End Sub

    Private selRowIndex As Integer
    Private selColIndex As Integer

    ''' <summary>catch key presses: Shift F10 or menu key to get the context menu, Ctrl-C/Ctrl-V for copy/pasting foreign lookup info</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DBSheetCols_KeyDown(sender As Object, e As KeyEventArgs) Handles DBSheetCols.KeyDown
        selRowIndex = DBSheetCols.CurrentCell.RowIndex
        selColIndex = DBSheetCols.CurrentCell.ColumnIndex
        If e.KeyCode = Keys.Apps Or (e.KeyCode = Keys.F10 And e.Modifiers = Keys.Shift) Or (e.KeyCode = Keys.Down And e.Modifiers = Keys.Alt) Then
            If DBSheetCols.SelectedRows.Count > 0 Then
                ' whole row selection -> move up/down menu...
                selColIndex = -1
                displayContextMenus()
                ' need show here as the context menu is not displayed otherwise...
                DBSheetCols.ContextMenuStrip.Show()
            ElseIf selColIndex = 8 And selRowIndex >= 0 Then
                displayContextMenus()
                DBSheetCols.ContextMenuStrip.Show()
            End If
            ' set to handled to avoid moving down cell selection (Keys.Down)
            e.Handled = True
        ElseIf e.KeyCode = Keys.C And e.Modifiers = Keys.Control Then
            clipboardDataRow = DBSheetCols.DataSource.GetNewRow()
            clipboardDataRow.ItemArray = DBSheetCols.DataSource.Rows(selRowIndex).ItemArray.Clone()
        ElseIf e.KeyCode = Keys.V And e.Modifiers = Keys.Control Then
            DBSheetCols.Rows(selRowIndex).Cells("ftable").Value = clipboardDataRow.ftable
            DBSheetCols.Rows(selRowIndex).Cells("fkey").Value = clipboardDataRow.fkey
            DBSheetCols.Rows(selRowIndex).Cells("flookup").Value = clipboardDataRow.flookup
            DBSheetCols.Rows(selRowIndex).Cells("lookup").Value = clipboardDataRow.lookup
            DBSheetCols.Rows(selRowIndex).Cells("outer").Value = clipboardDataRow.outer
            DBSheetCols.Rows(selRowIndex).Cells("primkey").Value = clipboardDataRow.primkey
            ' fill in the dropdown values
            If DBSheetCols.Rows(selRowIndex).Cells("ftable").Value.ToString = "" Then
                DirectCast(DBSheetCols.Rows(selRowIndex).Cells("fkey"), DataGridViewComboBoxCell).DataSource = New List(Of String)({""})
                DirectCast(DBSheetCols.Rows(selRowIndex).Cells("flookup"), DataGridViewComboBoxCell).DataSource = New List(Of String)({""})
            Else
                Dim forColsList As List(Of String) = getforeignTables(DBSheetCols.Rows(selRowIndex).Cells("ftable").Value.ToString)
                DirectCast(DBSheetCols.Rows(selRowIndex).Cells("fkey"), DataGridViewComboBoxCell).DataSource = forColsList
                DirectCast(DBSheetCols.Rows(selRowIndex).Cells("flookup"), DataGridViewComboBoxCell).DataSource = forColsList
            End If
        End If
    End Sub

    ''' <summary>display context menus depending on cell selected</summary>
    Private Sub displayContextMenus()
        If selColIndex = 8 And selRowIndex >= 0 Then
            DBSheetCols.ContextMenuStrip = DBSheetColsLookupMenu
        ElseIf selColIndex = -1 And selRowIndex >= 0 AndAlso DBSheetCols.SelectedRows.Count > 0 AndAlso DBSheetCols.SelectedRows(0).Index = selRowIndex Then
            DBSheetCols.ContextMenuStrip = DBSheetColsMoveMenu
        Else
            DBSheetCols.ContextMenuStrip = Nothing
        End If
    End Sub

    ''' <summary>prepare context menus to be displayed after right mouse click</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DBSheetCols_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DBSheetCols.CellMouseDown
        selRowIndex = e.RowIndex
        selColIndex = e.ColumnIndex
        If e.Button = Windows.Forms.MouseButtons.Right Then displayContextMenus()
    End Sub

    ''' <summary>move (shift) row up</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub MoveRowUp_Click(sender As Object, e As EventArgs) Handles MoveRowUp.Click
        Try
            ' avoid moving up of first row
            If selRowIndex = 0 Then Return
            If (DBSheetCols.DataSource.Rows.Count - 1 < selRowIndex) Then
                ErrorMsg("Editing not finished in selected row (values not committed), cannot move up!")
                Exit Sub
            End If
            If DBSheetCols.Rows(selRowIndex - 1).Cells("primkey").Value And Not DBSheetCols.Rows(selRowIndex).Cells("primkey").Value Then
                ErrorMsg("All primary keys have to be first and there is a primary key column that would be shifted below this non-primary one !", "DBSheet Definition Warning")
                Exit Sub
            End If
            Dim rw As DBSheetDefRow = DBSheetCols.DataSource.GetNewRow()
            rw.ItemArray = DBSheetCols.DataSource.Rows(selRowIndex).ItemArray.Clone()
            Dim colnameList As List(Of String) = DirectCast(DBSheetCols.Rows(selRowIndex).Cells("fkey"), DataGridViewComboBoxCell).DataSource
            FormDisabled = True
            DBSheetCols.DataSource.Rows.RemoveAt(selRowIndex)
            DBSheetCols.DataSource.Rows.InsertAt(rw, selRowIndex - 1)
            DirectCast(DBSheetCols.Rows(selRowIndex - 1).Cells("flookup"), DataGridViewComboBoxCell).DataSource = colnameList
            DirectCast(DBSheetCols.Rows(selRowIndex - 1).Cells("fkey"), DataGridViewComboBoxCell).DataSource = colnameList
            DBSheetCols.CurrentCell = DBSheetCols.Rows(selRowIndex - 1).Cells(0)
            DBSheetCols.Rows(selRowIndex - 1).Selected = True
        Catch ex As System.Exception
            ErrorMsg("Exception in MoveRowUpToolStripMenuItem_Click: " + ex.Message)
        End Try
        FormDisabled = False
    End Sub

    ''' <summary>move (shift) row down</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub MoveRowDown_Click(sender As Object, e As EventArgs) Handles MoveRowDown.Click
        Try
            ' avoid moving down of last row, DBSeqenceDataGrid.Rows.Count is 1 more than the actual inserted rows because of the "new" row, selIndex is 0 based
            If selRowIndex = DBSheetCols.Rows.Count - 2 Then Exit Sub
            If Not DBSheetCols.Rows(selRowIndex + 1).Cells("primkey").Value And DBSheetCols.Rows(selRowIndex).Cells("primkey").Value Then
                ErrorMsg("All primary keys have to be first and there is a non primary key column that would be shifted above this primary one !", "DBSheet Definition Warning")
                Exit Sub
            End If
            Dim rw As DBSheetDefRow = DBSheetCols.DataSource.GetNewRow()
            rw.ItemArray = DBSheetCols.DataSource.Rows(selRowIndex).ItemArray.Clone()
            Dim colnameList As List(Of String) = DirectCast(DBSheetCols.Rows(selRowIndex).Cells("fkey"), DataGridViewComboBoxCell).DataSource
            FormDisabled = True
            DBSheetCols.DataSource.Rows.RemoveAt(selRowIndex)
            DBSheetCols.DataSource.Rows.InsertAt(rw, selRowIndex + 1)
            DirectCast(DBSheetCols.Rows(selRowIndex + 1).Cells("flookup"), DataGridViewComboBoxCell).DataSource = colnameList
            DirectCast(DBSheetCols.Rows(selRowIndex + 1).Cells("fkey"), DataGridViewComboBoxCell).DataSource = colnameList
            DBSheetCols.CurrentCell = DBSheetCols.Rows(selRowIndex + 1).Cells(0)
            DBSheetCols.Rows(selRowIndex + 1).Selected = True
        Catch ex As System.Exception
            ErrorMsg("Exception in MoveRowDownToolStripMenuItem_Click: " + ex.Message)
        End Try
        FormDisabled = False
    End Sub

    ''' <summary>(re)generates the lookup query for active row/cell</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub RegenerateThisLookupQuery_Click(sender As Object, e As EventArgs) Handles RegenerateThisLookupQuery.Click
        regenLookupForRow(selRowIndex)
        DBSheetCols.AutoResizeColumns()
    End Sub

    ''' <summary>regenerate lookup for row in rowIndex</summary>
    ''' <param name="rowIndex"></param>
    Private Sub regenLookupForRow(rowIndex As Integer)
        If DBSheetCols.Rows(rowIndex).Cells("ftable").Value.ToString <> "" And DBSheetCols.Rows(rowIndex).Cells("fkey").Value.ToString <> "" And DBSheetCols.Rows(rowIndex).Cells("flookup").Value.ToString <> "" Then
            DBSheetCols.Rows(rowIndex).Cells("lookup").Value = "SELECT " + tblPlaceHolder + "." + DBSheetCols.Rows(rowIndex).Cells("flookup").Value.ToString + " " + correctNonNull(DBSheetCols.Rows(rowIndex).Cells("name").Value.ToString) + "," + tblPlaceHolder + "." + DBSheetCols.Rows(rowIndex).Cells("fkey").Value.ToString + " FROM " + DBSheetCols.Rows(rowIndex).Cells("ftable").Value.ToString + " " + tblPlaceHolder + " ORDER BY " + DBSheetCols.Rows(rowIndex).Cells("flookup").Value.ToString
        Else
            ErrorMsg("No lookup query to regenerate as foreign keys are not (fully) defined for field " + DBSheetCols.Rows(rowIndex).Cells("name").Value)
        End If
    End Sub

    ''' <summary>(re)generates ALL lookup queries</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub RegenerateAllLookupQueries_Click(sender As Object, e As EventArgs) Handles RegenerateAllLookupQueries.Click
        Dim retval As MsgBoxResult = QuestionMsg("regenerate foreign lookups completely, overwriting all customizations there: yes" + vbCrLf + "generate only new: no", MsgBoxStyle.YesNoCancel, "DBSheet Definition")
        If retval = MsgBoxResult.Cancel Then
            FormDisabled = False
            Exit Sub
        End If
        For i As Integer = 0 To DBSheetCols.Rows.Count - 2
            'only overwrite if forced regenerate or empty restriction def...
            If (retval = MsgBoxResult.Yes Or DBSheetCols.Rows(i).Cells("lookup").Value.ToString = "") Then regenLookupForRow(i)
        Next
        DBSheetCols.AutoResizeColumns()
    End Sub

    ''' <summary>test the (generated or manually edited) lookup query in currently selected row</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TestLookupQuery_Click(sender As Object, e As EventArgs) Handles TestLookupQuery.Click
        If Strings.Len(DBSheetCols.Rows(selRowIndex).Cells("lookup").Value) > 0 Then
            testTheQuery(DBSheetCols.Rows(selRowIndex).Cells("lookup").Value, True)
        Else
            ErrorMsg("No restriction query created to test !!!", "DBSheet Query Test Warning")
        End If
    End Sub

    ''' <summary>removes the lookup query test currently open</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub RemoveLookupQueryTest_Click(sender As Object, e As EventArgs) Handles RemoveLookupQueryTest.Click
        If ExcelDnaUtil.Application.ActiveSheet.Name <> "TESTSHEET" Then
            ErrorMsg("Active sheet doesn't seem to be a query test sheet !!!", "DBSheet Testsheet Remove Warning")
        Else
            ExcelDnaUtil.Application.ActiveWorkbook.Close(False)
        End If
    End Sub

    ''' <summary>ignore data errors when loading data into gridview</summary>
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
            Try : openConnection() : Catch ex As Exception : ErrorMsg(ex.Message) : Exit Sub : End Try
            Dim rstSchema As OdbcDataReader
            Dim selectStmt As String = "SELECT TOP 1 * FROM " + Table.Text
            Dim sqlCommand As OdbcCommand = New OdbcCommand(selectStmt, dbshcnn)
            rstSchema = sqlCommand.ExecuteReader()
            Try
                Dim schemaInfo As DataTable = rstSchema.GetSchemaTable()
                Dim theDBSheetDefTable = New DBSheetDefTable
                For Each schemaRow As DataRow In schemaInfo.Rows
                    Dim appendInfo As String = If(schemaRow("AllowDBNull"), "", specialNonNullableChar)
                    Dim newRow As DBSheetDefRow = theDBSheetDefTable.GetNewRow()
                    newRow.name = appendInfo + schemaRow("ColumnName")
                    'fist field is always primary col by default:
                    If firstRow Then newRow.primkey = True
                    firstRow = False
                    newRow.type = TableDataTypes(newRow.name)
                    theDBSheetDefTable.Add(newRow)
                Next
                DBSheetCols.DataSource = theDBSheetDefTable
                DBSheetCols.AutoResizeColumns()
            Catch ex As Exception
                ErrorMsg("Could not get schema information for table fields with query: '" + selectStmt + "', error: " + ex.Message)
            End Try
            rstSchema.Close()
            FormDisabled = False
            ' after changing the column no more change to table allowed !!
            TableEditable(False)
        Catch ex As System.Exception
            ErrorMsg("Exception in addAllFields_Click: " + ex.Message)
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
        FormDisabled = True
        Table.SelectedIndex = -1
        Query.Text = ""
        WhereClause.Text = ""
        FormDisabled = False
        ' reset the current filename
        currentFilepath = ""
        CurrentFileLinkLabel.Text = ""
        saveEnabled(False)
        DBSheetColsEditable(False)
    End Sub

    Private TableDataTypes As Dictionary(Of String, String)

    ''' <summary>gets the types of currently selected table including size, precision and scale into DataTypes</summary>
    Private Sub getTableDataTypes()
        Try : openConnection() : Catch ex As Exception : ErrorMsg(ex.Message) : Exit Sub : End Try
        TableDataTypes = New Dictionary(Of String, String)
        Dim rstSchema As OdbcDataReader
        Dim selectStmt As String = "SELECT TOP 1 * FROM " + Table.Text
        Dim sqlCommand As OdbcCommand = New OdbcCommand(selectStmt, dbshcnn)
        rstSchema = sqlCommand.ExecuteReader()
        Try
            Dim schemaInfo As DataTable = rstSchema.GetSchemaTable()
            For Each schemaRow As DataRow In schemaInfo.Rows
                Dim appendInfo As String = If(schemaRow("AllowDBNull"), "", specialNonNullableChar)
                Dim precInfo As String = ""
                If schemaRow("DataType").Name <> "String" And schemaRow("DataType").Name <> "Boolean" Then
                    precInfo = "/" + schemaRow("NumericPrecision").ToString + "/" + schemaRow("NumericScale").ToString
                End If
                TableDataTypes(appendInfo + schemaRow("ColumnName")) = schemaRow("DataType").Name + "(" + schemaRow("ColumnSize").ToString + precInfo + ")"
            Next
        Catch ex As Exception
            ErrorMsg("Could not get type information for table fields with query: '" + selectStmt + "', error: " + ex.Message)
        End Try
        rstSchema.Close()
    End Sub

    ''' <summary>gets the possible foreign tables in the instance (over all databases)</summary>
    ''' <param name="foreignTable"></param>
    ''' <returns>List of columns</returns>
    Private Function getforeignTables(foreignTable As String) As List(Of String)
        getforeignTables = New List(Of String)({""})
        Try : openConnection() : Catch ex As Exception : ErrorMsg(ex.Message) : Exit Function : End Try
        Dim rstSchema As OdbcDataReader
        Dim selectStmt As String = "SELECT TOP 1 * FROM " + foreignTable
        Dim sqlCommand As OdbcCommand = New OdbcCommand(selectStmt, dbshcnn)
        rstSchema = sqlCommand.ExecuteReader()
        Try
            Dim schemaInfo As DataTable = rstSchema.GetSchemaTable()
            For Each schemaRow As DataRow In schemaInfo.Rows
                getforeignTables.Add(schemaRow("ColumnName"))
            Next
        Catch ex As Exception
            ErrorMsg("Could not get type information for table fields with query: '" + selectStmt + "', error: " + ex.Message)
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

    ''' <summary>fill all possible tables of configDatabase</summary>
    Private Sub fillTables()
        Dim schemaTable As DataTable
        Dim tableTemp As String
        Try : openConnection() : Catch ex As Exception : ErrorMsg(ex.Message) : Exit Sub : End Try
        Try
            schemaTable = dbshcnn.GetSchema("Tables")
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

    'TODO: check bug when multiple DBSHeet dialogs are open and lookup/key fields are not filled...
    ''' <summary>fill foreign tables into list of strings (is called by and filled in DBSheetCols.CellValueChanged)</summary>
    Private Function getforeignTables() As List(Of String)
        getforeignTables = New List(Of String)({""})
        Dim schemaTable As DataTable
        Try
            schemaTable = dbshcnn.GetSchema("Tables")
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

    'TODO: Complex select columns (anything that has more than just the table field) must have an alias associated, which has to be named like the foreign table key. If that is not the case, DBAddin wont be able to associate the foreign column in the main table with the lookup id, and thus displays following error message
#Region "Query Creation and Testing"
    ''' <summary>create the final DBSheet Main Query</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub createQuery_Click(ByVal sender As Object, ByVal e As EventArgs) Handles createQuery.Click
        testQuery.Text = "&test DBSheet Query"
        If DBSheetCols.Rows.Count < 2 Then
            ErrorMsg("No columns defined yet, can't create query !", "DBSheet Definition Error")
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
                Dim usedColumn As String = correctNonNull(DBSheetCols.Rows(i).Cells("name").Value.ToString)
                ' used for foreign table lookups
                tableCounter += 1
                If DBSheetCols.Rows(i).Cells("sort").Value.ToString <> "" Then
                    orderByStr = If(orderByStr = "", "", orderByStr + ", ") + (i + 1).ToString + " " + DBSheetCols.Rows(i).Cells("sort").Value.ToString
                End If
                Dim ftableStr As String = DBSheetCols.Rows(i).Cells("ftable").Value.ToString
                If ftableStr = "" Then
                    selectStr += "T1." + usedColumn + ", "
                    ' create (inner or outer) joins for foreign key lookup id
                Else
                    Dim lookupStr As String = DBSheetCols.Rows(i).Cells("lookup").Value.ToString
                    If lookupStr = "" Then
                        DBSheetCols.Rows(i).Selected = True
                        ErrorMsg("No lookup query created for field " + DBSheetCols.Rows(i).Cells("name").Value + ", can't proceed !")
                        Exit Sub
                    End If
                    Dim theTable As String = "T" + tableCounter.ToString
                    ' either we go for the whole part after the last join
                    Dim completeJoin As String = fetch(lookupStr, "JOIN ", "")
                    ' or we have a simple WHERE and just "AND" it to the created join
                    Dim addRestrict As String = quotedReplace(fetch(lookupStr, "WHERE ", ""), "T" + tableCounter.ToString)

                    ' remove any ORDER BY clause from additional restrict...
                    Dim restrPos As Integer = addRestrict.ToUpper().LastIndexOf(" ORDER") + 1
                    If restrPos > 0 Then addRestrict = addRestrict.Substring(0, Math.Min(restrPos - 1, addRestrict.Length))
                    If completeJoin <> "" Then
                        ' when having the complete join, use additional restriction not for main subtable
                        addRestrict = ""
                        ' instead make it an additional condition for the join and replace placeholder with tablealias
                        completeJoin = quotedReplace(ciReplace(completeJoin, "WHERE", "AND"), "T" + tableCounter.ToString)
                    End If
                    Dim fkeyStr As String = DBSheetCols.Rows(i).Cells("fkey").Value.ToString
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

                    selectPart = fetch(lookupStr, "SELECT ", " FROM ").Trim()
                    ' remove second field in lookup query's select clause
                    restrPos = selectPart.LastIndexOf(",") + 1
                    selectPart = selectPart.Substring(0, Math.Min(restrPos - 1, selectPart.Length))
                    ' complex select statement, take directly from lookup query..
                    Dim flookupStr As String = DBSheetCols.Rows(i).Cells("flookup").Value.ToString
                    If selectPart <> flookupStr Then
                        selectStr += quotedReplace(selectPart, "T" + tableCounter.ToString) + ", "
                    Else
                        ' simple select statement (only the lookup field and id), put together...
                        selectStr += theTable + "." + flookupStr + " AS " + usedColumn + ", "
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
            ErrorMsg("Exception in createQuery_Click: " + ex.Message)
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
                ElseIf testQuery.Text = "&remove Testsheet" Then
                    If ExcelDnaUtil.Application.ActiveSheet.Name <> "TESTSHEETQ" Then
                        ErrorMsg("Active sheet doesn't seem to be a query test sheet !", "DBSheet Testsheet Remove Warning", MessageBoxIcon.Exclamation)
                    Else
                        ExcelDnaUtil.Application.ActiveWorkbook.Close(False)
                    End If
                    testQuery.Text = "&test DBSheet Query"
                End If
            Else
                ErrorMsg("No Query created to test !!!", "DBSheet Query Test Warning", MessageBoxIcon.Exclamation)
            End If
        Catch ex As System.Exception
            ErrorMsg("Exception in testQuery_Click: " + ex.Message)
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
                Dim paramVal As String = InputBox("Value for parameter " + j.ToString + " ?", "Enter parameter values..")
                If Len(paramVal) = 0 Then Exit Sub
                replacedStr = Strings.Mid(replacedStr, 1, questionMarkLoc - 1) + paramVal + Strings.Mid(replacedStr, questionMarkLoc + 1)
                j += 1
            End While
            If InStr(theQueryText, whereClauseText) = 0 Then
                ErrorMsg("Didn't find where clause " + whereClauseText + " in theQueryText: " + theQueryText + vbCrLf + "maybe creating DBSheet query again helps..")
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
            ConfigFiles.createFunctionsInCells(Preview.Cells(1, 1), {"RC", "=DBListFetch(RC[1], """", R[1]C,,,True)"})
            newWB.Saved = True
            If isLookupQuery Then
                Preview.Name = "TESTSHEET"
            Else
                Preview.Name = "TESTSHEETQ"
                testQuery.Text = "&remove Testsheet"
            End If
            Exit Sub
        Catch ex As System.Exception
            ErrorMsg("Exception In testTheQuery: " + ex.Message)
        End Try
    End Sub
#End Region

#Region "DBSheet Definitionfiles Handling"

    ''' <summary>loads the DBSHeet definitions from a file (xml format)</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub loadDefs_Click(ByVal sender As Object, ByVal e As EventArgs) Handles loadDefs.Click
        Try
            Dim openFileDialog1 = New OpenFileDialog With {
                .InitialDirectory = fetchSetting("DBSheetDefinitions" + Globals.env, ""),
                .Filter = "XML files (*.xml)|*.xml",
                .RestoreDirectory = True
            }
            Dim result As DialogResult = openFileDialog1.ShowDialog()
            If result = Windows.Forms.DialogResult.OK Then
                ' remember path for possible storing in DBSheetParams
                currentFilepath = openFileDialog1.FileName
                CurrentFileLinkLabel.Text = currentFilepath
                Dim DBSheetParams As String = File.ReadAllText(currentFilepath, System.Text.Encoding.Default)
                ' fetch params into form from sheet or file
                FormDisabled = True
                ' get Database from (legacy) connID (legacy connID was prefixed with connIDPrefixDBtype)
                Dim configDatabase As String = Replace(DBSheetConfig.getEntry("connID", DBSheetParams), fetchSetting("connIDPrefixDBtype", "MSSQL"), "")
                Database.SelectedIndex = Database.Items.IndexOf(configDatabase)
                Try : openConnection() : Catch ex As Exception : ErrorMsg(ex.Message) : Exit Sub : End Try
                fillTables()
                DirectCast(DBSheetCols.Columns("ftable"), DataGridViewComboBoxColumn).DataSource = getforeignTables()
                FormDisabled = True
                Dim theTable As String = If(InStr(DBSheetConfig.getEntry("table", DBSheetParams), Database.Text + ".") > 0, DBSheetConfig.getEntry("table", DBSheetParams), Database.Text + fetchSetting("ownerQualifier" + env.ToString, "") + DBSheetConfig.getEntry("table", DBSheetParams))
                Table.SelectedIndex = Table.Items.IndexOf(theTable)
                If Table.SelectedIndex = -1 Then
                    ErrorMsg("couldn't find table " + theTable + " defined in definitions file in database " + Database.Text + "!")
                    FormDisabled = False
                    Exit Sub
                End If
                DirectCast(DBSheetCols.Columns("name"), DataGridViewComboBoxColumn).DataSource = getColumns()
                Dim columnslist As Object = DBSheetConfig.getEntryList("columns", "field", "", DBSheetParams)
                Dim theDBSheetDefTable = New DBSheetDefTable
                For Each DBSheetColumnDef As String In columnslist
                    Dim newRow As DBSheetDefRow = theDBSheetDefTable.GetNewRow()
                    newRow.name = DBSheetConfig.getEntry("name", DBSheetColumnDef)
                    newRow.ftable = DBSheetConfig.getEntry("ftable", DBSheetColumnDef)
                    newRow.fkey = DBSheetConfig.getEntry("fkey", DBSheetColumnDef)
                    newRow.flookup = DBSheetConfig.getEntry("flookup", DBSheetColumnDef)
                    newRow.outer = If(DBSheetConfig.getEntry("outer", DBSheetColumnDef) <> "", True, False)
                    newRow.primkey = If(DBSheetConfig.getEntry("primkey", DBSheetColumnDef) <> "", True, False)
                    newRow.type = TableDataTypes(newRow.name)
                    If newRow.type = "" Then Exit Sub
                    Dim sortMode As String = DBSheetConfig.getEntry("sort", DBSheetColumnDef)
                    newRow.sort = If(sortMode = "Ascending", "ASC", If(sortMode = "Descending", "DESC", ""))
                    newRow.lookup = DBSheetConfig.getEntry("lookup", DBSheetColumnDef)
                    theDBSheetDefTable.Add(newRow)
                Next
                DBSheetCols.DataSource = theDBSheetDefTable
                ' re-add the fkey and flookup combobox values...
                For i As Integer = 0 To DBSheetCols.Rows.Count - 2
                    If DBSheetCols.Rows(i).Cells("ftable").Value.ToString <> "" Then
                        Dim colnameList As List(Of String) = getforeignTables(DBSheetCols.Rows(i).Cells("ftable").Value.ToString)
                        DirectCast(DBSheetCols.Rows(i).Cells("fkey"), DataGridViewComboBoxCell).DataSource = colnameList
                        DirectCast(DBSheetCols.Rows(i).Cells("flookup"), DataGridViewComboBoxCell).DataSource = colnameList
                    End If
                Next
                DBSheetCols.AutoResizeColumns()
                Query.Text = DBSheetConfig.getEntry("query", DBSheetParams)
                WhereClause.Text = DBSheetConfig.getEntry("whereClause", DBSheetParams)
                TableEditable(False)
                DBSheetColsEditable(True)
                saveEnabled(True)
            End If
        Catch ex As System.Exception
            ErrorMsg("Exception in loadDefs_Click: " + ex.Message)
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

    Private currentFilepath As String

    ''' <summary>saves the definitions currently stored in theDBSheetCreateForm to newly selected file (saveAs = True) or to the file already stored in setting "dsdPath"</summary>
    ''' <param name="saveAs"></param>
    Private Sub saveDefinitionsToFile(ByRef saveAs As Boolean)
        Try
            If saveAs Or currentFilepath = "" Then
                Dim tableName As String = If(InStrRev(Table.Text, ".") > 0, Strings.Mid(Table.Text, InStrRev(Table.Text, ".") + 1), Table.Text)
                Dim saveFileDialog1 As SaveFileDialog = New SaveFileDialog With {
                    .Title = "Save DBSheet Definition",
                    .FileName = tableName + ".xml",
                    .InitialDirectory = fetchSetting("DBSheetDefinitions" + Globals.env, ""),
                    .Filter = "XML files (*.xml)|*.xml",
                    .RestoreDirectory = True
                }
                Dim result As DialogResult = saveFileDialog1.ShowDialog()
                If result = Windows.Forms.DialogResult.OK Then
                    currentFilepath = saveFileDialog1.FileName
                    CurrentFileLinkLabel.Text = currentFilepath
                Else
                    Exit Sub
                End If
            End If
            FileSystem.FileOpen(1, currentFilepath, OpenMode.Output)
            FileSystem.PrintLine(1, xmlDbsheetConfig())
            FileSystem.FileClose(1)
        Catch ex As System.Exception
            ErrorMsg("Exception in saveDefinitionsToFile: " + ex.Message)
        End Try
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
                    If DBSheetCols.Rows(i).Cells(j).Value.ToString <> "" Then
                        ' store everything false values and type (is always inferred from Database)
                        If Not (DBSheetCols.Columns(j).Name = "type" OrElse
                            (TypeName(DBSheetCols.Rows(i).Cells(j).Value) = "Boolean" AndAlso Not DBSheetCols.Rows(i).Cells(j).Value)) Then
                            columnLine += DBSheetConfig.setEntry(DBSheetCols.Columns(j).Name, CStr(DBSheetCols.Rows(i).Cells(j).Value.ToString))
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
            namedParams += DBSheetConfig.setEntry("primcols", primKeyCount.ToString)
            ' finally put everything together:
            Return "<DBsheetConfig>" + vbCrLf + namedParams + vbCrLf + "<columns>" + columnsDef + vbCrLf + "</columns>" + vbCrLf + "</DBsheetConfig>"
        Catch ex As System.Exception
            ErrorMsg("Exception in xmlDbsheetConfig: " + ex.Message)
            Return ""
        End Try
    End Function

    ''' <summary>current file link clicked: open possibility</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CurrentFileLinkLabel_Click(sender As Object, e As EventArgs) Handles CurrentFileLinkLabel.Click
        Diagnostics.Process.Start(CurrentFileLinkLabel.Text)
    End Sub
#End Region

    ''' <summary>opens a database connection with active connstring</summary>
    Sub openConnection()
        ' connections are pooled by ADO depending on the connection string:
        If InStr(1, dbsheetConnString, dbPwdSpec) > 0 And Strings.Len(existingPwd) = 0 Then
            Throw New Exception("Password is required by connection string: " + dbsheetConnString)
        End If
        If Strings.Len(existingPwd) > 0 Then
            If InStr(1, dbsheetConnString, dbPwdSpec) > 0 Then
                dbsheetConnString = Change(dbsheetConnString, dbPwdSpec, existingPwd, ";")
            Else
                dbsheetConnString = dbsheetConnString + ";" + dbPwdSpec + existingPwd
            End If
        End If
        Try
            dbshcnn = New OdbcConnection With {
                .ConnectionString = dbsheetConnString,
                .ConnectionTimeout = Globals.CnnTimeout
            }
            dbshcnn.Open()
        Catch ex As Exception
            dbsheetConnString = Replace(dbsheetConnString, dbPwdSpec + existingPwd, dbPwdSpec + "*******")
            dbshcnn = Nothing
            Throw New Exception("Error connecting to DB: " + ex.Message + ", connection string: " + dbsheetConnString)
        End Try
    End Sub

#Region "GUI Helper functions"
    ''' <summary>Table/Database/Password change possible</summary>
    ''' <param name="choice">possible True, not possible False</param>
    Private Sub TableEditable(choice As Boolean)
        Database.Enabled = choice
        LDatabase.Enabled = choice
        Table.Enabled = choice
        LTable.Enabled = choice
    End Sub

    ''' <summary>if DBSheetCols Definitions should be editable, enable relevant cotrols</summary>
    ''' <param name="choice">editable True, not editable False</param>
    Private Sub DBSheetColsEditable(choice As Boolean)
        ' block password change if DBSheetDef is editable
        Password.Enabled = Not choice
        LPwd.Enabled = Not choice
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
    ''' <summary>corrects field names of nonnullable fields prepended with specialNonNullableChar (e.g. "*") back to the real name</summary>
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
                subresult = teststr(i).ToString
            End If
            quotedReplace += subresult + IIf(i < UBound(teststr), "'", "")
        Next
    End Function

#End Region
End Class

#Region "DBSheetCols Gridview Helper Classes"
''' <summary>DataTable Class for filling DBSheetCols DataGridView</summary>
Public Class DBSheetDefTable : Inherits DataTable

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

    Public Sub Add(ByVal row As DBSheetDefRow)
        Rows.Add(row)
    End Sub

    Public Sub Remove(ByVal row As DBSheetDefRow)
        Rows.Remove(row)
    End Sub

    Public Function GetNewRow() As DBSheetDefRow
        Dim row As DBSheetDefRow = CType(NewRow(), DBSheetDefRow)
        Return row
    End Function

    Protected Overrides Function GetRowType() As Type
        Return GetType(DBSheetDefRow)
    End Function

    Protected Overrides Function NewRowFromBuilder(ByVal builder As DataRowBuilder) As DataRow
        Return New DBSheetDefRow(builder)
    End Function

End Class

''' <summary>DataRow Class for DBSheetDefTable</summary>
Public Class DBSheetDefRow : Inherits DataRow
    Public Property name As String
        Get
            Return CStr(MyBase.Item("name"))
        End Get
        Set(ByVal value As String)
            MyBase.Item("name") = value
        End Set
    End Property
    Public Property ftable As String
        Get
            Return CStr(MyBase.Item("ftable"))
        End Get
        Set(ByVal value As String)
            MyBase.Item("ftable") = value
        End Set
    End Property
    Public Property fkey As String
        Get
            Return CStr(MyBase.Item("fkey"))
        End Get
        Set(ByVal value As String)
            MyBase.Item("fkey") = value
        End Set
    End Property
    Public Property flookup As String
        Get
            Return CStr(MyBase.Item("flookup"))
        End Get
        Set(ByVal value As String)
            MyBase.Item("flookup") = value
        End Set
    End Property
    Public Property outer As Boolean
        Get
            Return CBool(MyBase.Item("outer"))
        End Get
        Set(ByVal value As Boolean)
            MyBase.Item("outer") = value
        End Set
    End Property
    Public Property primkey As Boolean
        Get
            Return CBool(MyBase.Item("primkey"))
        End Get
        Set(ByVal value As Boolean)
            MyBase.Item("primkey") = value
        End Set
    End Property
    Public Property type As String
        Get
            Return CStr(MyBase.Item("type"))
        End Get
        Set(ByVal value As String)
            MyBase.Item("type") = value
        End Set
    End Property
    Public Property sort As String
        Get
            Return CStr(MyBase.Item("sort"))
        End Get
        Set(ByVal value As String)
            MyBase.Item("sort") = value
        End Set
    End Property
    Public Property lookup As String
        Get
            Return CStr(MyBase.Item("lookup"))
        End Get
        Set(ByVal value As String)
            MyBase.Item("lookup") = value
        End Set
    End Property
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