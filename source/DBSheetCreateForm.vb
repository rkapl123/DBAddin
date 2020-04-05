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
    ''' <summary>specifiying the owner (schema) of a table</summary>
    Private ownerQualifier As String
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
        ownerQualifier = fetchSetting("ownerQualifier" + env.ToString, "NONEXISTENT")
        If ownerQualifier = "NONEXISTENT" Then
            ErrorMsg("No ownerQualifier given for environment: " + env.ToString + ", please correct and rerun.", "createDefinitions Error")
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
                    .DataPropertyName = "name"
                }
        Dim ftableCB As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn With {
                    .Name = "ftable",
                    .DataSource = New List(Of String),
                    .HeaderText = "ftable",
                    .DataPropertyName = "ftable"
                }
        Dim fkeyCB As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn With {
                    .Name = "fkey",
                    .DataSource = New List(Of String),
                    .HeaderText = "fkey",
                    .DataPropertyName = "fkey"
                }
        Dim flookupCB As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn With {
                    .Name = "flookup",
                    .DataSource = New List(Of String),
                    .HeaderText = "flookup",
                    .DataPropertyName = "flookup"
                }
        Dim outerCB As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn With {
                    .Name = "outer",
                    .HeaderText = "outer",
                    .DataPropertyName = "outer"
                }
        Dim primkeyCB As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn With {
                    .Name = "primkey",
                    .HeaderText = "primkey",
                    .DataPropertyName = "primkey"
                }
        Dim ColTypeTB As DataGridViewTextBoxColumn = New DataGridViewTextBoxColumn With {
                    .Name = "ColType",
                    .HeaderText = "ColType",
                    .DataPropertyName = "ColType",
                    .[ReadOnly] = True
                }
        Dim sortCB As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn With {
                    .Name = "sort",
                    .DataSource = New List(Of String)({"None", "Ascending", "Descending"}),
                    .HeaderText = "sort",
                    .DataPropertyName = "sort"
                }
        Dim lookupTB As DataGridViewTextBoxColumn = New DataGridViewTextBoxColumn With {
                    .Name = "lookup",
                    .HeaderText = "lookup",
                    .DataPropertyName = "lookup"
                }
        DBSheetCols.AutoGenerateColumns = False
        DBSheetCols.Columns.AddRange(nameCB, ftableCB, fkeyCB, flookupCB, outerCB, primkeyCB, ColTypeTB, sortCB, lookupTB)
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

    Private Sub setPasswordAndInit()
        existingPwd = Password.Text
        Try : dbshcnn.Close() : Catch ex As Exception : End Try
        dbshcnn = Nothing
        fillDatabasesAndSetDropDown()
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

    ''' <summary>fill the Database dropdown and set dropdown to database set in connection string</summary>
    Private Sub fillDatabasesAndSetDropDown()
        Try
            fillDatabases(Database)
        Catch ex As System.Exception
            ErrorMsg("Error: " + ex.Message)
            Exit Sub
        End Try
        Me.Text = "DB Sheet creation: Select Database and Table to start building a DBSheet Definition"
        Database.SelectedIndex = Database.Items.IndexOf(fetch(dbsheetConnString, dbidentifier, ";"))
        'initialization of everything else is triggered by above change and caught by Database_SelectedIndexChanged
    End Sub

    ''' <summary>database changed, initialize everything (Tables, DBSheetCols definition) from scratch</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Database_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Database.SelectedIndexChanged
        Try
            If Not openConnection() Then
                ErrorMsg("Couldn't open connection to database " + Database.Text)
                Exit Sub
            End If
            fillTables(Database.Text)
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
            ErrorMsg("Error: " + ex.Message, "Database_SelectedIndexChanged")
        End Try
    End Sub

    ''' <summary>selecting the Table triggers enabling the DBSheetCols definition (fils columns/fields of that table and resetting DBSheetCols definition)</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    Private Sub Table_SelectedIndexChanged(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles Table.SelectedIndexChanged
        If FormDisabled Then Exit Sub
        Try
            FormDisabled = True
            If Table.SelectedIndex >= 0 Then DBSheetColsEditable(True)
            ' just in case this wasn't cleared before...
            DBSheetCols.DataSource = Nothing
            DBSheetCols.Rows.Clear()
            Query.Text = ""
            fillColumns()
            fillForTables()
            FormDisabled = False
        Catch ex As System.Exception
            ErrorMsg("Error: " + ex.Message)
        End Try
        Me.Text = "DB Sheet creation: Select one or more columns (fields) adding possible foreign key lookup information in foreign tables, finally click create query to finish DBSheet definition"
    End Sub

    ''' <summary>handles the various changes in the DBSheetCols gridview</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    Private Sub DBSheetCols_CellValueChanged(ByVal eventSender As Object, ByVal eventArgs As DataGridViewCellEventArgs) Handles DBSheetCols.CellValueChanged
        If FormDisabled Then Exit Sub
        Dim selIndex As Integer = eventArgs.RowIndex
        If eventArgs.ColumnIndex = 0 Then  ' field name column 
            ' fill foreign tables ..
            fillForTables()
            ' ..and column type
            DBSheetCols.Rows(selIndex).Cells("ColType").Value = TableDataTypes(DBSheetCols.Rows(selIndex).Cells("name").Value)
            ' if first column then always primary key!
            If selIndex = 0 Then DBSheetCols.Rows(selIndex).Cells("primkey").Value = True
        ElseIf eventArgs.ColumnIndex = 1 Then         ' ftable column -> fill fkey and flookup
            Dim colnameList As List(Of String) = getforeignTableColumns(DBSheetCols.Rows(selIndex).Cells("ftable").Value)
            FormDisabled = True
            DirectCast(DBSheetCols.Rows(selIndex).Cells("fkey"), DataGridViewComboBoxCell).DataSource = colnameList
            DirectCast(DBSheetCols.Rows(selIndex).Cells("flookup"), DataGridViewComboBoxCell).DataSource = colnameList
            FormDisabled = False
        ElseIf eventArgs.ColumnIndex = 5 Then ' primkey column
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
                ErrorMsg("Error: " + ex.Message)
            End Try
        End If
    End Sub

    Private selRowIndex As Integer
    Private selColIndex As Integer
    Private Sub DBSheetCols_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DBSheetCols.CellMouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then
            selRowIndex = e.RowIndex
            selColIndex = e.ColumnIndex
            If selColIndex = 8 And selRowIndex >= 0 Then
                DBSheetCols.ContextMenuStrip = DBSheetColsLookupMenu
                DBSheetColsLookupMenu.Items(0).Text = "regenerate lookup query"
                DBSheetColsLookupMenu.Items(1).Visible = True
                DBSheetColsLookupMenu.Items(2).Visible = True
            ElseIf selColIndex = -1 And selRowIndex >= 0 AndAlso DBSheetCols.SelectedRows.Count > 0 AndAlso DBSheetCols.SelectedRows(0).Index = selRowIndex Then
                DBSheetCols.ContextMenuStrip = DBSheetColsMoveMenu
            ElseIf selColIndex = 8 And selRowIndex = -1 Then
                DBSheetCols.ContextMenuStrip = DBSheetColsLookupMenu
                DBSheetColsLookupMenu.Items(0).Text = "regenerate all lookup queries"
                DBSheetColsLookupMenu.Items(1).Visible = False
                DBSheetColsLookupMenu.Items(2).Visible = False
            Else
                DBSheetCols.ContextMenuStrip = Nothing
            End If
        End If
    End Sub

    Private Sub MoveRowUpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MoveRowUpToolStripMenuItem.Click
        Try
            ' avoid moving up of first row
            If selRowIndex = 0 Then Return
            If DBSheetCols.Rows(selRowIndex + 1).Cells("primkey").Value And Not DBSheetCols.Rows(selRowIndex).Cells("primkey").Value Then
                ErrorMsg("All primary keys have to be first and there is a primary key column that would be shifted below this non-primary one !", "DBSheet Definition Warning")
                Exit Sub
            End If
            Dim dsdt As DBSheetDefTable = DBSheetCols.DataSource
            Dim rw As DBSheetDefRow = dsdt.GetNewRow()
            rw.ItemArray = dsdt.Rows(selRowIndex).ItemArray.Clone()
            Dim colnameList As List(Of String) = DirectCast(DBSheetCols.Rows(selRowIndex).Cells("fkey"), DataGridViewComboBoxCell).DataSource
            FormDisabled = True
            dsdt.Rows.RemoveAt(selRowIndex)
            dsdt.Rows.InsertAt(rw, selRowIndex - 1)
            DirectCast(DBSheetCols.Rows(selRowIndex - 1).Cells("flookup"), DataGridViewComboBoxCell).DataSource = colnameList
            DirectCast(DBSheetCols.Rows(selRowIndex - 1).Cells("fkey"), DataGridViewComboBoxCell).DataSource = colnameList
            DBSheetCols.CurrentCell = DBSheetCols.Rows(selRowIndex - 1).Cells(0)
            DBSheetCols.Rows(selRowIndex - 1).Selected = True
        Catch ex As System.Exception
            ErrorMsg("Error: " + ex.Message)
        End Try
        FormDisabled = False
    End Sub

    Private Sub MoveRowDownToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MoveRowDownToolStripMenuItem.Click
        Try
            ' avoid moving down of last row, DBSeqenceDataGrid.Rows.Count is 1 more than the actual inserted rows because of the "new" row, selIndex is 0 based
            If selRowIndex = DBSheetCols.Rows.Count - 2 Then Exit Sub
            If Not DBSheetCols.Rows(selRowIndex + 1).Cells("primkey").Value And DBSheetCols.Rows(selRowIndex).Cells("primkey").Value Then
                ErrorMsg("All primary keys have to be first and there is a non primary key column that would be shifted above this primary one !", "DBSheet Definition Warning")
                Exit Sub
            End If
            Dim dsdt As DBSheetDefTable = DBSheetCols.DataSource
            Dim rw As DBSheetDefRow = dsdt.GetNewRow()
            rw.ItemArray = dsdt.Rows(selRowIndex).ItemArray.Clone()
            Dim colnameList As List(Of String) = DirectCast(DBSheetCols.Rows(selRowIndex).Cells("fkey"), DataGridViewComboBoxCell).DataSource
            FormDisabled = True
            dsdt.Rows.RemoveAt(selRowIndex)
            dsdt.Rows.InsertAt(rw, selRowIndex + 1)
            DirectCast(DBSheetCols.Rows(selRowIndex + 1).Cells("flookup"), DataGridViewComboBoxCell).DataSource = colnameList
            DirectCast(DBSheetCols.Rows(selRowIndex + 1).Cells("fkey"), DataGridViewComboBoxCell).DataSource = colnameList
            DBSheetCols.CurrentCell = DBSheetCols.Rows(selRowIndex + 1).Cells(0)
            DBSheetCols.Rows(selRowIndex + 1).Selected = True
        Catch ex As System.Exception
            ErrorMsg("Error: " + ex.Message)
        End Try
        FormDisabled = False
    End Sub

    ''' <summary>(re)generates the lookup query for active cell or all cells..</summary>
    Private Sub RegenerateLookupQueryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        If selColIndex = 8 And selRowIndex = -1 Then
            Dim retval As MsgBoxResult = QuestionMsg("regenerating Foreign Lookups completely (overwriting all customizations there): yes or generate only new: no !", MsgBoxStyle.YesNoCancel, "DBSheet Definition")
            If retval = MsgBoxResult.Cancel Then
                FormDisabled = False
                Exit Sub
            End If
            For i As Integer = 0 To DBSheetCols.Rows.Count - 2
                'only overwrite if forced regenerate or empty restriction def...
                If (retval = MsgBoxResult.Yes Or DBSheetCols.Rows(i).Cells("lookup").Value.ToString = "") And (DBSheetCols.Rows(i).Cells("ftable").Value.ToString <> "" And DBSheetCols.Rows(i).Cells("fkey").Value.ToString <> "" And DBSheetCols.Rows(i).Cells("flookup").Value.ToString <> "") Then
                    DBSheetCols.Rows(i).Cells("lookup").Value = "SELECT " + DBSheetCols.Rows(i).Cells("flookup").Value.ToString + "," + DBSheetCols.Rows(i).Cells("fkey").Value.ToString + " FROM " + DBSheetCols.Rows(i).Cells("ftable").Value.ToString + " ORDER BY " + DBSheetCols.Rows(i).Cells("flookup").Value.ToString
                End If
            Next
        Else
            If DBSheetCols.Rows(selRowIndex).Cells("ftable").Value.ToString <> "" And DBSheetCols.Rows(selRowIndex).Cells("fkey").Value.ToString <> "" And DBSheetCols.Rows(selRowIndex).Cells("flookup").Value.ToString <> "" Then
                DBSheetCols.Rows(selRowIndex).Cells("lookup").Value = "SELECT " + DBSheetCols.Rows(selRowIndex).Cells("flookup").Value.ToString + "," + DBSheetCols.Rows(selRowIndex).Cells("fkey").Value.ToString + " FROM " + DBSheetCols.Rows(selRowIndex).Cells("ftable").Value.ToString + " ORDER BY " + DBSheetCols.Rows(selRowIndex).Cells("flookup").Value.ToString
            Else
                ErrorMsg("No lookup query to regenerate as foreign keys are not (fully) defined for field " + DBSheetCols.Rows(selRowIndex).Cells("name").Value)
            End If
        End If
    End Sub

    ''' <summary>test the (generated or manually edited) lookup query in currently selected row</summary>
    Private Sub TestLookupQueryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        If Strings.Len(DBSheetCols.Rows(selRowIndex).Cells("lookup").Value) > 0 Then
            testTheQuery(DBSheetCols.Rows(selRowIndex).Cells("lookup").Value, True)
        Else
            ErrorMsg("No restriction query created to test !!!", "DBSheet Query Test Warning")
        End If
    End Sub

    ''' <summary>removes the lookup query test currently open</summary>
    Private Sub RemoveLookupQueryTestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem5.Click
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

    ''' <summary>add all fields of currently selected Table to DBSheetCols definitions</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    Private Sub addAllFields_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles addAllFields.Click
        Try
            FormDisabled = True
            Dim firstRow As Boolean = True
            DBSheetCols.DataSource = Nothing
            DBSheetCols.Rows.Clear()
            Dim rstSchema As OdbcDataReader
            If Not openConnection() Then Exit Sub
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
                    newRow.ColType = TableDataTypes(newRow.name)
                    theDBSheetDefTable.Add(newRow)
                Next
                DBSheetCols.DataSource = theDBSheetDefTable
            Catch ex As Exception
                ErrorMsg("Could not get schema information for table fields with query: '" + selectStmt + "', error: " + ex.Message)
            End Try
            rstSchema.Close()
            FormDisabled = False
            ' after changing the column no more change to table allowed !!
            TableEditable(False)
        Catch ex As System.Exception
            ErrorMsg("Error: " + ex.Message)
        End Try
    End Sub

    ''' <summary>clears the defined columns and resets the selection fields (Table, ForTable) and the Query</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    Private Sub clearAllFields_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles clearAllFields.Click
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
        saveEnabled(False)
        DBSheetColsEditable(False)
    End Sub

    ''' <summary>moves selected row up</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    Private Sub moveUp_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs)

    End Sub

    ''' <summary>moves selected row down</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    Private Sub moveDown_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs)
        Try
            Dim selIndex As Integer = DBSheetCols.CurrentRow.Index
            ' avoid moving down of last row, DBSeqenceDataGrid.Rows.Count is 1 more than the actual inserted rows because of the "new" row, selIndex is 0 based
            If Not DBSheetCols.CurrentRow.Selected Or selIndex = DBSheetCols.Rows.Count - 2 Then Exit Sub
            If Not DBSheetCols.Rows(selIndex + 1).Cells("primkey").Value And DBSheetCols.Rows(selIndex).Cells("primkey").Value Then
                ErrorMsg("All primary keys have to be first and there is a non primary key column that would be shifted above this primary one !", "DBSheet Definition Warning")
                Exit Sub
            End If
            If IsNothing(DBSheetCols.CurrentRow) Then Return
            Dim rw As DataGridViewRow = DBSheetCols.CurrentRow
            DBSheetCols.Rows.RemoveAt(selIndex)
            DBSheetCols.Rows.Insert(selIndex + 1, rw)
            DBSheetCols.Rows(selIndex + 1).Cells(0).Selected = True
        Catch ex As System.Exception
            ErrorMsg("Error: " + ex.Message)
        End Try
    End Sub

    Private TableDataTypes As Dictionary(Of String, String)

    ''' <summary>gets the types of currently selected table including size, precision and scale into DataTypes</summary>
    Private Sub getTableDataTypes()
        TableDataTypes = New Dictionary(Of String, String)
        Dim rstSchema As OdbcDataReader
        If Not openConnection() Then Exit Sub
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

    ''' <summary>gets the columns of the foreignTable</summary>
    ''' <param name="foreignTable"></param>
    ''' <returns>List of columns</returns>
    Private Function getforeignTableColumns(foreignTable As String) As List(Of String)
        getforeignTableColumns = New List(Of String)({""})
        Dim rstSchema As OdbcDataReader
        If Not openConnection() Then Exit Function
        Dim selectStmt As String = "SELECT TOP 1 * FROM " + foreignTable
        Dim sqlCommand As OdbcCommand = New OdbcCommand(selectStmt, dbshcnn)
        rstSchema = sqlCommand.ExecuteReader()
        Try
            Dim schemaInfo As DataTable = rstSchema.GetSchemaTable()
            For Each schemaRow As DataRow In schemaInfo.Rows
                getforeignTableColumns.Add(schemaRow("ColumnName"))
            Next
        Catch ex As Exception
            ErrorMsg("Could not get type information for table fields with query: '" + selectStmt + "', error: " + ex.Message)
        End Try
        rstSchema.Close()
    End Function

    ''' <summary>fill all possible columns of currently selected table</summary>
    Private Sub fillColumns()
        getTableDataTypes()
        Dim colnameList As List(Of String) = New List(Of String)({""})
        Try
            For Each colname As String In TableDataTypes.Keys
                colnameList.Add(colname)
            Next
            FormDisabled = True
            DirectCast(DBSheetCols.Columns("name"), DataGridViewComboBoxColumn).DataSource = colnameList
            FormDisabled = False
        Catch ex As System.Exception
            Throw New Exception("Exception in fillColumns: " + ex.Message)
        End Try
    End Sub

    ''' <summary>fill all possible tables of configDatabase</summary>
    Private Sub fillTables(configDatabase As String)
        Dim schemaTable As DataTable
        Dim tableTemp As String

        If Not openConnection() Then
            Throw New Exception("could not open connection for database '" + Database.Text + "' in connection string '" + dbsheetConnString + "'.")
        End If
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
                If iteration_row("TABLE_CAT") = Database.Text Or iteration_row("TABLE_SCHEM") = Database.Text Then Table.Items.Add(iteration_row("TABLE_NAME"))
            Next iteration_row
            If Strings.Len(tableTemp) > 0 Then Table.SelectedIndex = Table.Items.IndexOf(tableTemp)
            FormDisabled = False
        Catch ex As System.Exception
            FormDisabled = False
            Throw New Exception("Exception in fillTables: " + ex.Message)
        End Try
    End Sub

    ''' <summary>fill foreign tables, see above</summary>
    Private Sub fillForTables()
        Dim schemaTable As DataTable
        Try
            schemaTable = dbshcnn.GetSchema("Tables")
            If schemaTable.Rows.Count = 0 Then Throw New Exception("No Tables could be fetched from Schema")
        Catch ex As Exception
            Throw New Exception("Error getting schema information for tables in connection strings database ' " + Database.Text + "'." + ",error: " + ex.Message)
        End Try
        Try
            Dim forTableList As List(Of String) = New List(Of String)({""})
            For Each iteration_row As DataRow In schemaTable.Rows
                forTableList.Add(iteration_row("TABLE_CAT") + "." + iteration_row("TABLE_SCHEM") + "." + iteration_row("TABLE_NAME"))
            Next iteration_row
            FormDisabled = True
            DirectCast(DBSheetCols.Columns("ftable"), DataGridViewComboBoxColumn).DataSource = forTableList
            FormDisabled = False
        Catch ex As System.Exception
            Throw New Exception("Exception in fillForTables: " + ex.Message)
        End Try
    End Sub

    ''' <summary>fills all possible databases of current connection using db proprietary code in dbGetAllStr, data coming from field DBGetAllFieldName</summary>
    Private Sub fillDatabases(DatabaseComboBox As ComboBox)
        Dim addVal As String
        Dim dbs As OdbcDataReader

        If Not openConnection() Then Exit Sub
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
                Throw New Exception("Exception: " + ex.Message)
            End Try
        Else
            FormDisabled = False
            Throw New Exception("Could not retrieve any databases with: " + dbGetAllStr + "!")
        End If
    End Sub

    ''' <summary>create the final DBSheet Main Query</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    Private Sub createQuery_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles createQuery.Click
        If DBSheetCols.Rows.Count = 0 Then
            ErrorMsg("No columns defined yet, can't create Query !", "DBSheet Definition Error")
            Exit Sub
        End If
        Dim retval As DialogResult = QuestionMsg("regenerating DBSheet Query, overwriting all customizations there !", MessageBoxButtons.OKCancel, "DBSheet Definition Warning", MessageBoxIcon.Exclamation)
        If retval = MsgBoxResult.Cancel Then Exit Sub
        Dim queryStr As String = createTheQuery()
        If Strings.Len(queryStr) > 0 Then Query.Text = queryStr
    End Sub

    ''' <summary>test the final DBSheet Main query</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    Private Sub testQuery_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles testQuery.Click
        Try
            If Strings.Len(Query.Text) > 0 Then
                If testQuery.Text = "test DBSheet Query" Then
                    testTheQuery(Query.Text)
                ElseIf testQuery.Text = "remove Testsheet" Then
                    If ExcelDnaUtil.Application.ActiveSheet.Name <> "TESTSHEETQ" Then
                        MessageBox.Show("Active sheet doesn't seem to be a query test sheet !!!", "DBAddin: DBSheet Testsheet Remove Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Else
                        ExcelDnaUtil.Application.ActiveWorkbook.Close(False)
                    End If
                    testQuery.Text = "test DBSheet Query"
                End If
            Else
                MessageBox.Show("No Query created to test !!!", "DBAddin: DBSheet Query Test Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        Catch ex As System.Exception
            ErrorMsg("Error: " + ex.Message)
        End Try
    End Sub

    ''' <summary>for testing either the main query or the selected lookup query being given in theQueryText</summary>
    ''' <param name="theQueryText"></param>
    ''' <param name="isLookupQuery"></param>
    Private Sub testTheQuery(ByVal theQueryText As String, Optional ByRef isLookupQuery As Boolean = False)
        theQueryText = theQueryText.Replace(vbCrLf, " ").Replace(vbLf, " ")
        If isLookupQuery Then theQueryText = quotedReplace(theQueryText, "FT")
        ' quoted replace of "?" with parameter values
        ' needs splitting of WhereClause by quotes !
        ' only for main query !!
        If Not isLookupQuery Then
            Dim teststr() As String = Split(WhereClause.Text, "'")
            Dim whereStr As String = vbNullString
            Dim j, i As Integer
            Dim subresult As String
            j = 1
            For i = 0 To UBound(teststr)
                If i Mod 2 = 0 Then
                    Dim replacedStr As String = teststr(i)
                    While InStr(1, replacedStr, "?")
                        Dim paramVal As String = InputBox("Value for parameter " + j.ToString + " ?", "Enter parameter values..")
                        If Len(paramVal) = 0 Then Exit Sub
                        Dim questionMarkLoc As Integer
                        questionMarkLoc = InStr(1, replacedStr, "?")
                        replacedStr = Strings.Mid(replacedStr, 1, questionMarkLoc - 1) + paramVal + Strings.Mid(replacedStr, questionMarkLoc + 1)
                        j += 1
                    End While
                    subresult = replacedStr
                Else
                    subresult = teststr(i)
                End If
                whereStr = whereStr + subresult + IIf(i < UBound(teststr), "'", "")
            Next
        End If
        Dim rst As DataSet = New DataSet()
        Try
            Dim adap As OdbcDataAdapter = New OdbcDataAdapter(theQueryText, dbshcnn)
            rst.Tables.Clear()
            adap.Fill(rst)
        Catch ex As Exception
            LogWarn("Error in query: " + theQueryText + vbCrLf + ex.Message)
            Exit Sub
        End Try
        Try
            ExcelDnaUtil.Application.SheetsInNewWorkbook = 1
            Dim newWB As Excel.Workbook = ExcelDnaUtil.Application.Workbooks.Add
            Dim Preview As Excel.Worksheet = newWB.Sheets(1)
            Preview.Select()
            With Preview.QueryTables.Add(rst, Preview.Range("A1"))
                .FieldNames = True
                .AdjustColumnWidth = True
                .RefreshStyle = Excel.XlCellInsertionMode.xlOverwriteCells
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .Refresh()
                .Delete()
            End With
            rst = Nothing
            newWB.Saved = True
            Preview.Select()
            If isLookupQuery Then
                Preview.Name = "TESTSHEET"
            Else
                Preview.Name = "TESTSHEETQ"
                testQuery.Text = "remove Testsheet"
            End If
            Exit Sub
        Catch ex As System.Exception
            ErrorMsg("Error: " + ex.Message)
        End Try
    End Sub

    ''' <summary>creates the main query from the column definitions found in DBsheetCols</summary>
    ''' <returns>the generated query</returns>
    Private Function createTheQuery() As String
        Dim selectStr As String = "", orderByStr As String = ""
        Try
            ' always take primary table database from connection definition !
            Dim fromStr As String = "FROM " + ownerQualifier + Table.Text + " T1"
            Dim tableCounter As Integer = 1
            Dim selectPart As String = ""
            For i As Integer = 0 To DBSheetCols.Rows.Count - 2
                ' plain table field
                Dim usedColumn As String = correctNonNull(DBSheetCols.Rows(i).Cells("name").Value.ToString)
                tableCounter += 1
                Dim sortChoice As String = DBSheetCols.Rows(i).Cells("sort").Value.ToString
                Select Case sortChoice
                    Case "Ascending" : orderByStr = If(orderByStr = "", "", orderByStr + ", ") + (i + 1).ToString + " ASC"
                    Case "Descending" : orderByStr = If(orderByStr = "", "", orderByStr + ", ") + (i + 1).ToString + " DESC"
                End Select
                Dim ftableStr As String = DBSheetCols.Rows(i).Cells("ftable").Value.ToString
                If ftableStr = "" Then
                    selectStr += "T1." + usedColumn + ", "
                    ' create (inner or outer) joins for foreign key lookup id
                Else
                    Dim lookupStr As String = DBSheetCols.Rows(i).Cells("lookup").Value.ToString
                    If lookupStr = "" Then
                        DBSheetCols.Rows(i).Selected = True
                        ErrorMsg("No Lookup Query created for field " + DBSheetCols.Rows(i).Cells("name").Value + ", can't proceed !")
                        Return ""
                    End If
                    Dim theTable As String = "T" + tableCounter.ToString
                    ' either we go for the whole part after the last join
                    Dim completeJoin As String = fetch(lookupStr, "JOIN ", "")
                    ' or we have a simple WHERE and just "AND" it to the created join
                    Dim addRestrict As String = quotedReplace(fetch(lookupStr, "WHERE ", ""), "T" + tableCounter.ToString)

                    ' remove any ORDER BY clause from additional restrict...
                    Dim restrPos As Integer = addRestrict.ToUpper().LastIndexOf(" ORDER") + 1
                    If restrPos > 0 Then addRestrict = addRestrict.Substring(0, Math.Min(restrPos - 1, addRestrict.Length))
                    If Strings.Len(completeJoin) > 0 Then
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
                    If Strings.Len(completeJoin) > 0 Then
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
            createTheQuery = selectStr + vbCrLf + fromStr.ToString() + vbCrLf +
                     If(Strings.Len(wherePart) > 0, "WHERE " + wherePart + vbCrLf, "") +
                     If(Strings.Len(orderByStr) > 0, "ORDER BY " + orderByStr, "")
            saveEnabled(True)
        Catch ex As System.Exception
            ErrorMsg("Error: " + ex.Message)
            createTheQuery = ""
        End Try
    End Function

    ''' <summary>loads the DBSHeet definitions from a file (xml format)</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    Private Sub loadDefs_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles loadDefs.Click
        Try
            Dim openFileDialog1 = New OpenFileDialog With {
                .InitialDirectory = fetchSetting("DBSheetDefinitions", ""),
                .Filter = "XML files (*.xml)|*.xml",
                .RestoreDirectory = True
            }
            Dim result As DialogResult = openFileDialog1.ShowDialog()
            If result = Windows.Forms.DialogResult.OK Then
                Dim retval As String = openFileDialog1.FileName
                If Strings.Len(retval) = 0 Then Exit Sub
                ' remember path for possible storing in DBSheetParams
                currentFilepath = retval
                Dim DBSheetParams As String = File.ReadAllText(currentFilepath, System.Text.Encoding.Default)
                ' fetch params into form from sheet or file
                FormDisabled = True
                ' get Database from (legacy) connID (legacy connID was prefixed with connIDPrefixDBtype)
                Dim configDatabase As String = Replace(DBSheetConfig.getEntry("connID", DBSheetParams), fetchSetting("connIDPrefixDBtype", "MSSQL"), "")
                Database.SelectedIndex = Database.Items.IndexOf(configDatabase)
                If Not openConnection() Then
                    ErrorMsg("Couldn't open connection to database " + Database.Text)
                    Exit Sub
                End If
                fillTables(Database.Text)
                fillForTables()
                FormDisabled = True
                Table.SelectedIndex = Table.Items.IndexOf(DBSheetConfig.getEntry("table", DBSheetParams))
                If Table.SelectedIndex = -1 Then
                    ErrorMsg("couldn't find table " + DBSheetConfig.getEntry("table", DBSheetParams) + " defined in definitions file in database " + Database.Text + "!")
                    FormDisabled = False
                    Exit Sub
                End If
                fillColumns()
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
                    newRow.ColType = TableDataTypes(newRow.name)
                    If newRow.ColType = "" Then Exit Sub
                    newRow.sort = DBSheetConfig.getEntry("sort", DBSheetColumnDef)
                    newRow.lookup = DBSheetConfig.getEntry("lookup", DBSheetColumnDef)
                    theDBSheetDefTable.Add(newRow)
                Next
                DBSheetCols.DataSource = theDBSheetDefTable
                Query.Text = DBSheetConfig.getEntry("query", DBSheetParams)
                WhereClause.Text = DBSheetConfig.getEntry("whereClause", DBSheetParams)
                TableEditable(False)
                DBSheetColsEditable(True)
                saveEnabled(True)
            End If
        Catch ex As System.Exception
            ErrorMsg("Error: " + ex.Message)
        End Try
    End Sub

    ''' <summary>save definitions button</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    Private Sub saveDefs_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles saveDefs.Click
        saveDefinitionsToFile(False)
    End Sub

    ''' <summary>save definitions as button</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    Private Sub saveDefsAs_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles saveDefsAs.Click
        saveDefinitionsToFile(True)
    End Sub

    ''' <summary>toggle saveEnabled behaviour</summary>
    ''' <param name="choice"></param>
    Private Sub saveEnabled(ByRef choice As Boolean)
        saveDefs.Enabled = choice
        saveDefsAs.Enabled = choice
    End Sub

    Private currentFilepath As String

    ''' <summary>saves the definitions currently stored in theDBSheetCreateForm to newly selected file (saveAs = True) or to the file already stored in setting "dsdPath"</summary>
    ''' <param name="saveAs"></param>
    Private Sub saveDefinitionsToFile(ByRef saveAs As Boolean)

        Dim fileToStore As String = FileSystem.Dir(currentFilepath, FileAttribute.Normal)
        Try
            If Strings.Len(fileToStore) = 0 Or saveAs Or Strings.Len(currentFilepath) = 0 Then
                Dim saveFileDialog1 As SaveFileDialog = New SaveFileDialog With {
                    .Title = "Save DBSheet Definition",
                    .FileName = Table.Text + ".xml",
                    .InitialDirectory = fetchSetting("DBSheetDefinitions" + Globals.env, ""),
                    .Filter = "XML files (*.xml)|*.xml",
                    .RestoreDirectory = True
                }
                Dim result As DialogResult = saveFileDialog1.ShowDialog()
                If result = Windows.Forms.DialogResult.OK Then
                    fileToStore = saveFileDialog1.FileName
                Else
                    Exit Sub
                End If
                currentFilepath = fileToStore
            End If
            FileSystem.FileOpen(1, fileToStore, OpenMode.Output)
            FileSystem.PrintLine(1, xmlDbsheetConfig())
            FileSystem.FileClose(1)
        Catch ex As System.Exception
            ErrorMsg("Error: " + ex.Message)
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
                    If Not IsDBNull(DBSheetCols.Rows(i).Cells(j).Value) Then
                        ' store everything except "none" sorting, false values and ColType (is always inferred from Database)
                        If Not ((DBSheetCols.Columns(j).Name = "sort" AndAlso DBSheetCols.Rows(i).Cells(j).Value = "None") OrElse
                            DBSheetCols.Columns(j).Name = "ColType" OrElse (TypeName(DBSheetCols.Rows(i).Cells(j).Value) = "Boolean" And Not DBSheetCols.Rows(i).Cells(j).Value)) Then
                            columnLine += DBSheetConfig.setEntry(DBSheetCols.Columns(j).Name, CStr(DBSheetCols.Rows(i).Cells(j).Value))
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
            ErrorMsg("Error: " + ex.Message)
            Return ""
        End Try
    End Function

    ''' <summary>opens a database connection with active connstring</summary>
    ''' <returns>true on success</returns>
    Function openConnection() As Boolean
        openConnection = False
        ' connections are pooled by ADO depending on the connection string:
        If InStr(1, dbsheetConnString, dbPwdSpec) > 0 And Strings.Len(existingPwd) = 0 Then
            ErrorMsg("Password is required by connection string: " + dbsheetConnString, "Open Connection Error")
            Exit Function
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
            openConnection = True
        Catch ex As Exception
            dbsheetConnString = Replace(dbsheetConnString, dbPwdSpec + existingPwd, dbPwdSpec + "*******")
            ErrorMsg("Error connecting to DB: " + ex.Message + ", connection string: " + dbsheetConnString, "Open Connection Error")
            dbshcnn = Nothing
        End Try
    End Function

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

    ''' <summary>Table/Database choice change possible</summary>
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
        addAllFields.Enabled = choice
        clearAllFields.Enabled = choice
        DBSheetCols.Enabled = choice
        createQuery.Enabled = choice
        testQuery.Enabled = choice
        Query.Enabled = choice
        WhereClause.Enabled = choice
    End Sub

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

End Class

''' <summary>Helper Class for filling DBSheetCols DataGridView</summary>
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
        Columns.Add(New DataColumn("ColType", GetType(String)))
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

''' <summary>Row Class for DBSheetDefTable</summary>
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
    Public Property ColType As String
        Get
            Return CStr(MyBase.Item("ColType"))
        End Get
        Set(ByVal value As String)
            MyBase.Item("ColType") = value
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
        ColType = ""
        sort = ""
        lookup = ""
    End Sub
End Class