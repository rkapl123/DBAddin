Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.Odbc
Imports System.IO
Imports System.Windows.Forms

''' <summary>Form for defining/creating DBSheets</summary>
Partial Friend Class DBSheetCreateForm
    Inherits System.Windows.Forms.Form
    ''' <summary>whether the form fields should react to changes (set if making changes within code)....</summary>
    Private FormDisabled As Boolean
    ''' <summary>sometimes we make an exception to FormDisabled ...</summary>
    Private ForceFieldUpdate As Boolean
    ''' <summary>last selected column</summary>
    Private last As Integer
    Private dbsheetConnString As String
    Private CtrlPressed As Boolean
    Private maxColCount As Integer

    Private dbidentifier As String
    Private ownerQualifier As String

    Private dbGetAllStr As String
    Private DBGetAllFieldName As String
    Private dbshcnn As OdbcConnection
    Private dbPwdSpec As String
    Private tblPlaceHolder As String = "!T!"
    Private specialNonNullableChar As String = "*"

    ''' <summary>sets up DBSheetCreateForm for editing DBSHeet definitions</summary>
    ''' <param name="DBSheetParams"></param>
    ''' <remarks>Main entry point for DBSheetCreateForm, invoked by clicking "create/edit DBSheet definition" or loadDefs Button (loads stored connection definitions into Connection tab)</remarks>
    Public Sub createDefinitions(Optional ByVal DBSheetParams As String = "")
        Try
            maxColCount = 0

            ' if we have a valid dbsheet definition (either selected a valid dbsheeet or loaded from file)
            ' fetch params into form from sheet or file
            If DBSheetParams <> "" Then
                FormDisabled = True
                ' get Database from (legacy) connID (legacy connID prefixed with connIDPrefixDBtype)
                Dim configDatabase As String = Replace(DBSheetConfig.getEntry("connID", DBSheetParams), fetchSetting("connIDPrefixDBtype", "MSSQL"), "")
                Database.SelectedIndex = Database.Items.IndexOf(configDatabase)
                If Not openConnection(Database.Text) Then
                    ErrorMsg("Couldn't open connection to database " & Database.Text)
                    Exit Sub
                End If
                fillTables(Database.Text)
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
                    newRow.outer = If(DBSheetConfig.getEntry("outer", DBSheetColumnDef) = 1, True, False)
                    newRow.primkey = If(DBSheetConfig.getEntry("primkey", DBSheetColumnDef) = 1, True, False)
                    newRow.ColType = TableDataTypes(newRow.name)
                    If newRow.ColType = "" Then Exit Sub
                    newRow.sort = DBSheetConfig.getEntry("sort", DBSheetColumnDef)
                    newRow.lookup = DBSheetConfig.getEntry("lookup", DBSheetColumnDef)
                    theDBSheetDefTable.Add(newRow)
                Next
                DBSheetCols.DataSource = theDBSheetDefTable
                Query.Text = DBSheetConfig.getEntry("query", DBSheetParams)
                WhereClause.Text = DBSheetConfig.getEntry("whereClause", DBSheetParams)
                fillForTables()
                TableEditable(False)
                saveEnabled(True)
            Else
                If Not openConnection(Database.Text) Then
                    ErrorMsg("Couldn't open connection to database " & Database.Text)
                    Exit Sub
                End If
                fillTables(Database.Text)
                ' start with empty columns list
                TableEditable(True)
                saveEnabled(False)
                loadDefs.Enabled = True
                FormDisabled = True
                Query.Text = ""
                WhereClause.Text = ""
                DBSheetCols.DataSource = Nothing
                DBSheetCols.Rows.Clear()
                FormDisabled = False
            End If
        Catch ex As System.Exception
            ErrorMsg("Error: " & ex.Message)
        End Try
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
                    .FileName = Table.Text & ".xml",
                    .InitialDirectory = fetchSetting("DBSheetDefinitions" & Globals.env, ""),
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
            ErrorMsg("Error: " & ex.Message)
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
                            DBSheetCols.Columns(j).Name = "ColType" OrElse
                            (TypeName(DBSheetCols.Rows(i).Cells(j).Value) = "Boolean" AndAlso Not DBSheetCols.Rows(i).Cells(j).Value)) Then
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
            ErrorMsg("Error: " & ex.Message)
            Return ""
        End Try
    End Function

    Private Sub testLookupQuery_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles testLookupQuery.Click
        Try
            Dim testcheck As String = ""
            'TODO: change LookupQuery to Gridview Textbox
            'If Strings.Len(LookupQuery.Text) > 0 Then
            '    If testLookupQuery.Text = "test &Lookup Query" Then
            '        testTheQuery(LookupQuery.Text, True)
            '    ElseIf testLookupQuery.Text = "remove &Lookup Testsheet" Then
            '        ' TODO: check for lookup testsheet...
            '        If (testcheck.IndexOf("TESTSHEET") + 1) = 0 Then
            '            ErrorMsg("Active sheet doesn't seem to be a query test sheet !!!", "DBSheet Testsheet Remove Warning")
            '        Else
            '            ExcelDnaUtil.Application.ActiveWorkbook.Close(False)
            '        End If
            '        testLookupQuery.Text = "test &Lookup Query"
            '    End If
            'Else
            '    ErrorMsg("No restriction query created to test !!!", "DBSheet Query Test Warning")
            'End If
        Catch ex As System.Exception
            ErrorMsg("Error: " & ex.Message)
        End Try
    End Sub

    'TODO: change to gridview combobox foreign DB change
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    Private Sub ForDatabase_SelectedIndexChanged(ByVal eventSender As Object, ByVal eventArgs As EventArgs)
        If FormDisabled Then Exit Sub
        fillForTables()
    End Sub

    Private Sub Table_SelectedIndexChanged(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles Table.SelectedIndexChanged
        If FormDisabled Then Exit Sub
        Try
            FormDisabled = True
            If Table.SelectedIndex >= 0 Then
                addAllFields.Enabled = True
                DBSheetCols.Enabled = True
            End If
            ' just in case this wasn't cleared before...
            DBSheetCols.DataSource = Nothing
            DBSheetCols.Rows.Clear()
            Query.Text = ""
            fillColumns()
            columnEditMode(False)
            FormDisabled = False
        Catch ex As System.Exception
            ErrorMsg("Error: " & ex.Message)
        End Try
        Me.Text = "DB Sheet creation: Select on or more columns (fields) adding possible foreign key lookup information in foreign tables"
    End Sub

    'TODO: change to Gridview event checkedStateChanged
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    Private Sub isPrimary_CheckStateChanged(ByVal eventSender As Object, ByVal eventArgs As DataGridViewCellEventArgs) Handles DBSheetCols.CellValueChanged
        If FormDisabled Then Exit Sub
        ' primkey column
        If eventArgs.ColumnIndex = 5 Then
            Try
                Dim selIndex As Integer = eventArgs.RowIndex
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
                ErrorMsg("Error: " & ex.Message)
            End Try
        End If
    End Sub

    'TODO: change to gridview combobox
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    Private Sub Column_SelectedIndexChanged(ByVal eventSender As Object, ByVal eventArgs As EventArgs)
        If FormDisabled Then Exit Sub
        TableEditable(False)
        FormDisabled = True

        FormDisabled = False
    End Sub

    'TODO: change to gridview filling
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    Private Sub addAllFields_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles addAllFields.Click
        Dim rstSchema As DataSet

        Try
            FormDisabled = True
            rstSchema = dbshcnn.GetSchema().DataSet
            Dim firstRow As Boolean : firstRow = True
            DBSheetCols.DataSource = Nothing
            DBSheetCols.Rows.Clear()
            Dim newRow As Integer
            For Each iteration_row As DataRow In rstSchema.Tables(0).Rows
                If iteration_row("TABLE_CATALOG").ToUpper() = Database.Text.ToUpper() Or iteration_row("TABLE_SCHEMA").ToUpper() = Database.Text.ToUpper() Then
                    Dim attached As String = ""
                    If Not iteration_row("IS_NULLABLE") Then attached = specialNonNullableChar
                    newRow = DBSheetCols.Rows.Add(New DataGridViewRow())
                    'TODO: change theDBSheetColumnList
                    'theDBSheetColumnList.Value(newRow, 0) = attached & iteration_row("COLUMN_NAME")
                    ''fist field is always primary col by default:
                    'If firstRow Then theDBSheetColumnList.Value(newRow, 5) = 1
                    'firstRow = False
                    'theDBSheetColumnList.Value(newRow, 6) = getType_Renamed(iteration_row("COLUMN_NAME"))
                    'theDBSheetColumnList.Value(newRow, 7) = "None"
                End If
            Next iteration_row
            columnEditMode(False)
            FormDisabled = False
            ExcelDnaUtil.Application.EnableEvents = True
            ' after changing the column no more change to table allowed !!
            TableEditable(False)
        Catch ex As System.Exception
            ErrorMsg("Error: " & ex.Message)
        End Try
    End Sub

    'TODO: change to gridview combobox (adding a row)
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    Private Sub addToDBsheetCols_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs)
        'Try
        '    If maxColCount = 0 Then
        '        maxColCount = ExcelDnaUtil.ExcelLimits.MaxColumns
        '    End If
        '    If theDBSheetColumnList.RowCount = maxColCount Then
        '        ErrorMsg("Max. Columns allowed in DBSheet: " & maxColCount & " (last column reserved for data status)", "DBSheet Definition Warning")
        '        Exit Sub
        '    End If

        '    ExcelDnaUtil.Application.EnableEvents = False
        '    Dim newRow As Integer
        '    newRow = DBSheetCols.Rows.Add(New DataGridViewRow())
        '    FormDisabled = True


        '    ' Foreign Table information
        '    If Strings.Len(ForTable.Text) > 0 And Strings.Len(ForTableKey.Text) > 0 And Strings.Len(ForTableLookup.Text) > 0 Then
        '        'TODO: change theDBSheetColumnList
        '        'theDBSheetColumnList.Value(newRow, 1) = ForDatabase.Text & ownerQualifier & ForTable.Text
        '        'theDBSheetColumnList.Value(newRow, 2) = ForTableKey.Text
        '        'theDBSheetColumnList.Value(newRow, 3) = ForTableLookup.Text
        '        'If outerJoin.CheckState = CheckState.Checked Then theDBSheetColumnList.Value(newRow, 4) = 1
        '    ElseIf Strings.Len(ForTable.Text) > 0 Or Strings.Len(ForTableKey.Text) > 0 Or Strings.Len(ForTableLookup.Text) > 0 And Strings.Len(LookupQuery.Text) = 0 Then
        '        ErrorMsg("Please specify all 3 foreign column informations: ForeignTable, ForeignTableKey and ForeignTableLookup !", "DBSheet Definition Warning")
        '    End If

        '    ' Primary key
        '    If newRow = 0 Then ' always have first column as PK
        '        'TODO: change theDBSheetColumnList
        '        'theDBSheetColumnList.Value(newRow, 5) = 1
        '        IsPrimary.CheckState = CheckState.Checked
        '    End If
        '    ' check if primary keys are first
        '    Dim primaryAllowed As Boolean
        '    primaryAllowed = True
        '    For i As Integer = 0 To newRow
        '        'If Strings.Len(theDBSheetColumnList.Value(i, 5)) = 0 Then
        '        '    primaryAllowed = False
        '        '    Exit For
        '        'End If
        '    Next
        '    If IsPrimary.CheckState = CheckState.Checked Then
        '        If primaryAllowed Then
        '            'TODO: change theDBSheetColumnList
        '            'theDBSheetColumnList.Value(newRow, 5) = 1
        '        Else
        '            MessageBox.Show("Primary Keys must be first in a DBSheet (please place above)", "DBAddin: DBSheet Definition Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        '            IsPrimary.CheckState = CheckState.Unchecked
        '        End If
        '    End If

        '    columnEditMode(False)
        '    FormDisabled = False
        '    ExcelDnaUtil.Application.EnableEvents = True
        '    TableEditable(False) ' after changing the column no more change to table allowed !!
        'Catch ex As System.Exception
        '    ExcelDnaUtil.Application.EnableEvents = True
        '    ErrorMsg("Error: " & ex.Message)
        'End Try
    End Sub


    ''' <summary>clears the defined columns and resets the selection fields (Table, ForTable) and the Query</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    Private Sub clearAllFields_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles clearAllFields.Click
        FormDisabled = True
        DBSheetCols.DataSource = Nothing
        DBSheetCols.Rows.Clear()
        TableEditable(True)
        Table.SelectedIndex = -1
        Query.Text = ""
        WhereClause.Text = ""
        FormDisabled = False
        ' reset the current filename
        currentFilepath = ""
        saveEnabled(False)
        columnEditMode(False)
    End Sub


    'TODO: change to gridview
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    Private Sub regenLookupQueries_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles regenLookupQueries.Click
        'Try
        '    FormDisabled = True
        '    Dim retval As MsgBoxResult
        '    If regenLookupQueries.Text = "re&generate this lookup query" Then
        '        LookupQuery.Text = "SELECT " & ForTableLookup.Text & "," & ForTableKey.Text & " FROM " & ForDatabase.Text & ownerQualifier & ForTable.Text & " ORDER BY " & ForTableLookup.Text
        '    Else
        '        retval = QuestionMsg("regenerating Foreign Lookups completely (overwriting all customizations there): yes or generate only new: no !", MsgBoxStyle.YesNoCancel, "DBSheet Definition")
        '        If retval = MsgBoxResult.Cancel Then
        '            FormDisabled = False
        '            Exit Sub
        '        End If
        '        For i As Integer = 0 To theDBSheetColumnList.RowCount - 1
        '            If Strings.Len(theDBSheetColumnList.Value(i, 1)) > 0 Then
        '                'only overwrite if forced regenerate or empty restriction def...
        '                If retval = MsgBoxResult.Yes Or Strings.Len(theDBSheetColumnList.Value(i, 9)) = 0 Then
        '                    theDBSheetColumnList.Value(i, 9) = "SELECT " & theDBSheetColumnList.Value(i, 3) & "," & theDBSheetColumnList.Value(i, 2) & " FROM " & theDBSheetColumnList.Value(i, 1) & " ORDER BY " & theDBSheetColumnList.Value(i, 3)
        '                End If
        '            End If
        '        Next
        '    End If
        '    FormDisabled = False
        'Catch ex As System.Exception
        '    ErrorMsg("Error: " & ex.Message)
        'End Try
    End Sub

    ''' <summary>moves selected row up</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    Private Sub moveUp_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles moveUp.Click
        Try
            Dim selIndex As Integer = DBSheetCols.CurrentRow.Index
            If Not DBSheetCols.CurrentRow.Selected Or selIndex = 0 Then Exit Sub
            If DBSheetCols.CurrentCell.RowIndex = 0 And Not DBSheetCols.Rows(selIndex).Cells("primkey").Value Then
                ErrorMsg("first column always has to be primary key", "DBSheet Definition Warning")
                Exit Sub
            ElseIf DBSheetCols.Rows(selIndex + 1).Cells("primkey").Value And Not DBSheetCols.Rows(selIndex).Cells("primkey").Value Then
                ErrorMsg("All primary keys have to be first and there is a primary key column that would be shifted below this non-primary one !", "DBSheet Definition Warning")
                Exit Sub
            End If
            If IsNothing(DBSheetCols.CurrentRow) Then Return
            Dim rw As DataGridViewRow = DBSheetCols.CurrentRow
            ' avoid moving up of first row
            If selIndex = 0 Then Return
            DBSheetCols.Rows.RemoveAt(selIndex)
            DBSheetCols.Rows.Insert(selIndex - 1, rw)
            DBSheetCols.Rows(selIndex - 1).Cells(0).Selected = True
            last -= 1
            columnEditMode(True)
        Catch ex As System.Exception
            ErrorMsg("Error: " & ex.Message)
        End Try
    End Sub

    ''' <summary>moves selected row down</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    Private Sub moveDown_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles moveDown.Click
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
            last += 1
            columnEditMode(True)
        Catch ex As System.Exception
            ErrorMsg("Error: " & ex.Message)
        End Try
    End Sub


    ''' <summary>switch column edit mode: change regenerate this/all lookups, visible restrictions (moveUp/Down)</summary>
    ''' <param name="choice"></param>
    ''' <remarks>sets(choice=true) or resets(choice=false) column "edit" mode</remarks>
    Private Sub columnEditMode(ByRef choice As Boolean)
        FormDisabled = True
        moveDown.Visible = choice
        moveUp.Visible = choice
        If choice Then
            regenLookupQueries.Text = "re&generate this lookup query"
        Else
            DBSheetCols.ClearSelection()
            regenLookupQueries.Text = "re&generate all lookup queries"
        End If
        FormDisabled = False
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
                TableDataTypes(appendInfo + schemaRow("ColumnName")) = schemaRow("DataType").Name + "(" + schemaRow("ColumnSize").ToString + If(schemaRow("DataType").Name <> "String", "/" + schemaRow("NumericPrecision").ToString + "/" + schemaRow("NumericScale").ToString, "") + ")"
            Next
        Catch ex As Exception
            ErrorMsg("Could not get type information for table fields with query: '" & selectStmt & "', error: " & ex.Message)
        End Try
        rstSchema.Close()
    End Sub

    ''' <summary>fill all possible columns of currently selected table</summary>
    Private Sub fillColumns()
        getTableDataTypes()
        Dim colnameList As List(Of String) = New List(Of String)
        Try
            For Each colname As String In TableDataTypes.Keys
                colnameList.Add(colname)
            Next
            DirectCast(DBSheetCols.Columns("name"), DataGridViewComboBoxColumn).DataSource = colnameList
            FormDisabled = False
        Catch ex As System.Exception
            Throw New Exception("Exception in fillColumns: " & ex.Message)
        End Try
    End Sub

    ''' <summary>fill all possible tables of configDatabase</summary>
    Private Sub fillTables(configDatabase As String)
        Dim schemaTable As DataTable
        Dim tableTemp As String

        If Not openConnection(configDatabase) Then
            FormDisabled = False
            Throw New Exception("could not open connection for database '" & Database.Text & "' in connection string '" & dbsheetConnString & "'.")
        End If
        Try
            schemaTable = dbshcnn.GetSchema("Tables")
            If schemaTable.Rows.Count = 0 Then Throw New Exception("No Tables could be fetched from Schema")
        Catch ex As Exception
            FormDisabled = False
            Throw New Exception("Error getting schema information for tables in connection strings database ' " & Database.Text & "'." & ",error: " & ex.Message)
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
            Throw New Exception("Exception in fillTables: " & ex.Message)
        End Try
    End Sub

    ''' <summary>fill foreign tables, see above</summary>
    Private Sub fillForTables()
        Dim schemaTable As DataTable
        Dim tableTemp As String

        'TODO: change to gridview combobox
        'If Not openConnection(ForDatabase.Text) Then
        '    FormDisabled = False
        '    ForTable.Items.Clear() : ForTableKey.Items.Clear() : ForTableLookup.Items.Clear()
        '    Throw New Exception("could not open connection for foreign database '" & ForDatabase.Text & "' in connection string '" & dbsheetConnString & "'.")
        'End If

        'FormDisabled = True
        'tableTemp = ForTable.Text
        'ForTable.Items.Clear()
        'Try
        '    schemaTable = dbshcnn.GetSchema("Tables")
        '    If schemaTable.Rows.Count = 0 Then Throw New Exception("No Tables could be fetched from Schema")
        'Catch ex As Exception
        '    FormDisabled = False
        '    Throw New Exception("Error getting schema information for tables in connection strings database ' " & Database.Text & "'." & ",error: " & ex.Message)
        'End Try
        'Try
        '    For Each iteration_row As DataRow In schemaTable.Rows
        '        If iteration_row("TABLE_CAT").ToUpper() = ForDatabase.Text Or iteration_row("TABLE_SCHEM") = ForDatabase.Text Then ForTable.Items.Add(iteration_row("TABLE_NAME"))
        '    Next iteration_row
        '    If Strings.Len(tableTemp) > 0 Then ForTable.SelectedIndex = ForTable.Items.IndexOf(tableTemp)
        '    FormDisabled = False
        'Catch ex As System.Exception
        '    Throw New Exception("Exception in fillForTables: " & ex.Message)
        'End Try
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
            Throw New Exception("Could not retrieve schema information for databases in connection string: '" & dbsheetConnString & "',error: " & ex.Message)
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
                Throw New Exception("Exception: " & ex.Message)
            End Try
        Else
            FormDisabled = False
            Throw New Exception("Could not retrieve any databases with: " & dbGetAllStr & "!")
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
            Dim testcheck As String = ""
            If Strings.Len(Query.Text) > 0 Then
                If testQuery.Text = "&test DBSheet Query" Then
                    testTheQuery(Query.Text)
                ElseIf testQuery.Text = "remove &Testsheet" Then
                    'TODO: check for testsheet..
                    If (testcheck.IndexOf("TESTSHEET") + 1) = 0 Then
                        MessageBox.Show("Active sheet doesn't seem to be a query test sheet !!!", "DBAddin: DBSheet Testsheet Remove Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Else
                        ExcelDnaUtil.Application.ActiveWorkbook.Close(False)
                    End If
                    testQuery.Text = "&test DBSheet Query"
                End If
            Else
                MessageBox.Show("No Query created to test !!!", "DBAddin: DBSheet Query Test Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        Catch ex As System.Exception
            ErrorMsg("Error: " & ex.Message)
        End Try
    End Sub

    ''' <summary>for testing either the main query or the selected restriction query being given in theQueryText</summary>
    ''' <param name="theQueryText"></param>
    ''' <param name="isRestrictQuery"></param>
    Private Sub testTheQuery(ByVal theQueryText As String, Optional ByRef isRestrictQuery As Boolean = False)
        Dim rst As DataSet
        Dim Preview As Excel.Worksheet
        Dim newWB As Excel.Workbook
        Dim teststr() As String
        Dim paramVal As String = "", replacedStr As String = "", whereStr As String = ""

        theQueryText = theQueryText.Replace(vbCrLf, " ")
        theQueryText = theQueryText.Replace(vbLf, " ")
        If isRestrictQuery Then theQueryText = quotedReplace(theQueryText, "FT")

        ' quoted replace of "?" with parameter values
        ' needs splitting of WhereClause by quotes !
        ' only for main query !!
        If Not isRestrictQuery Then
            teststr = Split(WhereClause.Text, "'")
            whereStr = vbNullString
            Dim j, i As Integer
            Dim subresult As String
            j = 1
            For i = 0 To UBound(teststr)
                If i Mod 2 = 0 Then
                    replacedStr = teststr(i)
                    While InStr(1, replacedStr, "?")
                        paramVal = InputBox("Value for parameter " & j & " ?", "Enter parameter values..")
                        If Len(paramVal) = 0 Then Exit Sub
                        Dim questionMarkLoc As Integer
                        questionMarkLoc = InStr(1, replacedStr, "?")
                        replacedStr = Strings.Mid(replacedStr, 1, questionMarkLoc - 1) & paramVal & Strings.Mid(replacedStr, questionMarkLoc + 1)
                        j += 1
                    End While
                    subresult = replacedStr
                Else
                    subresult = teststr(i)
                End If
                whereStr = whereStr & subresult & IIf(i < UBound(teststr), "'", "")
            Next
        End If

        rst = New DataSet()
        Try
            Dim adap As OdbcDataAdapter = New OdbcDataAdapter(theQueryText, dbshcnn)
            rst.Tables.Clear()
            adap.Fill(rst)
        Catch ex As Exception
            LogWarn("Error in query: " & theQueryText & vbCrLf & ex.Message)
            Exit Sub
        End Try

        Try
            ExcelDnaUtil.Application.SheetsInNewWorkbook = 1
            newWB = ExcelDnaUtil.Application.Workbooks.Add
            Preview = newWB.Sheets(1)
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
            If isRestrictQuery Then
                testLookupQuery.Text = "remove &Lookup Testsheet"
            Else
                testQuery.Text = "remove &Testsheet"
            End If
            Exit Sub
        Catch ex As System.Exception
            ErrorMsg("Error: " & ex.Message)
        End Try
    End Sub

    ''' <summary>creates the main query from the column definitions found in DBsheetCols</summary>
    ''' <returns>the generated query</returns>
    Private Function createTheQuery() As String
        Dim result As String = ""
        Dim selectStr As String = "", orderByStr As String = ""
        Dim theTable As String, usedColumn As String, fromStr As String
        Dim tableCounter As Integer

        Try
            ' always take primary table database from connection definition !
            fromStr = "FROM " & ownerQualifier & Table.Text & " T1"
            tableCounter = 1
            Dim completeJoin As String = "", addRestrict As String = ""
            Dim restrPos As Integer
            Dim selectPart As String = ""
            For i As Integer = 0 To DBSheetCols.Rows.Count - 2
                ' plain table field
                usedColumn = correctNonNull(DBSheetCols.Rows(i).Cells("name").Value)
                tableCounter += 1
                Select Case DBSheetCols.Rows(i).Cells("sort").Value
                    Case "Ascending" : orderByStr = IIf(orderByStr = "", "", orderByStr & ", ") & CStr(i + 1) & " ASC"
                    Case "Descending" : orderByStr = IIf(orderByStr = "", "", orderByStr & ", ") & CStr(i + 1) & " DESC"
                End Select
                If Strings.Len(DBSheetCols.Rows(i).Cells("ftable").Value) = 0 Then
                    selectStr = selectStr & "T1." & usedColumn & ", "
                    ' create (inner or outer) joins for foreign key lookup id
                Else
                    If Strings.Len(DBSheetCols.Rows(i).Cells("lookup").Value) = 0 Then
                        DBSheetCols.Rows(i).Selected = True
                        result = ""
                        ErrorMsg("No Lookup Query created for field " & DBSheetCols.Rows(i).Cells("name").Value & ", can't proceed !")
                        Return result
                    End If
                    theTable = "T" & tableCounter
                    ' either we go for the whole part after the last join
                    completeJoin = fetch(DBSheetCols.Rows(i).Cells("lookup").Value, "JOIN ", "")
                    ' or we have a simple WHERE and just "AND" it to the created join
                    addRestrict = quotedReplace(fetch(DBSheetCols.Rows(i).Cells("lookup").Value, "WHERE ", ""), "T" & tableCounter)

                    ' remove any ORDER BY clause from additional restrict...
                    restrPos = addRestrict.ToUpper().LastIndexOf(" ORDER") + 1
                    If restrPos > 0 Then addRestrict = addRestrict.Substring(0, Math.Min(restrPos - 1, addRestrict.Length))
                    If Strings.Len(completeJoin) > 0 Then
                        ' when having the complete join, use additional restriction not for main subtable
                        addRestrict = ""
                        ' instead make it an additional condition for the join and replace placeholder with tablealias
                        completeJoin = quotedReplace(ciReplace(completeJoin, "WHERE", "AND"), "T" & tableCounter)
                    End If
                    If DBSheetCols.Rows(i).Cells("outer").Value Then
                        fromStr += " LEFT JOIN " & Environment.NewLine & DBSheetCols.Rows(i).Cells("ftable").Value & " " & theTable &
                                       " ON " & "T1." & usedColumn & " = " & theTable & "." & DBSheetCols.Rows(i).Cells("fkey").Value & IIf(Strings.Len(addRestrict) > 0, " AND " & addRestrict, "")
                    Else
                        fromStr += " INNER JOIN " & Environment.NewLine & DBSheetCols.Rows(i).Cells("ftable").Value & " " & theTable &
                                       " ON " & "T1." & usedColumn & " = " & theTable & "." & DBSheetCols.Rows(i).Cells("fkey").Value & IIf(Strings.Len(addRestrict) > 0, " AND " & addRestrict, "")
                    End If
                    ' we have additionally joined (an)other table(s) for the lookup display...
                    If Strings.Len(completeJoin) > 0 Then
                        ' remove any ORDER BY clause from completeJoin...
                        restrPos = completeJoin.ToUpper().LastIndexOf(" ORDER") + 1
                        If restrPos > 0 Then completeJoin = completeJoin.Substring(0, Math.Min(restrPos - 1, completeJoin.Length))
                        ' ..and add join of additional subtable(s) to the query
                        fromStr += " LEFT JOIN " & Environment.NewLine & completeJoin
                    End If

                    selectPart = fetch(DBSheetCols.Rows(i).Cells("lookup").Value, "SELECT ", " FROM ").Trim()
                    ' remove second field in lookup query's select clause
                    restrPos = selectPart.LastIndexOf(",") + 1
                    selectPart = selectPart.Substring(0, Math.Min(restrPos - 1, selectPart.Length))
                    ' complex select statement, take directly from lookup query..
                    If selectPart <> DBSheetCols.Rows(i).Cells("flookup").Value Then
                        selectStr += quotedReplace(selectPart, "T" & tableCounter) & ", "
                    Else
                        ' simple select statement (only the lookup field and id), put together...
                        selectStr += theTable & "." & DBSheetCols.Rows(i).Cells("flookup").Value & " AS " & usedColumn & ", "
                    End If
                End If
            Next
            Dim wherePart As String = ""
            wherePart = WhereClause.Text.Replace(Environment.NewLine, "")
            selectStr = "SELECT " & selectStr.Substring(0, Math.Min(Strings.Len(selectStr) - 2, selectStr.Length))
            result = selectStr & Environment.NewLine & fromStr.ToString() & Environment.NewLine &
                     IIf(Strings.Len(wherePart) > 0, "WHERE " & wherePart & Environment.NewLine, "") &
                     IIf(Strings.Len(orderByStr) > 0, "ORDER BY " & orderByStr, "")
            saveEnabled(True)
        Catch ex As System.Exception
            ErrorMsg("Error: " & ex.Message)
        End Try
        Return result
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
                createDefinitions(DBSheetParams)
            End If
        Catch ex As System.Exception
            ErrorMsg("Error: " & ex.Message)
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

    ''' <summary>opens a database connection with active connstring, optionally changing database in the connection string</summary>
    ''' <param name="database"></param>
    ''' <returns>true on success</returns>
    Function openConnection(Optional database As String = "") As Boolean
        openConnection = False
        ' connections are pooled by ADO depending on the connection string:
        If InStr(1, dbsheetConnString, dbPwdSpec) > 0 And Strings.Len(Password.Text) = 0 Then
            ErrorMsg("Password is required by connection string: " & dbsheetConnString, "Open Connection Error")
            Exit Function
        End If
        If database <> "" Then dbsheetConnString = Change(dbsheetConnString, dbidentifier, database, ";")
        If Strings.Len(Password.Text) > 0 Then
            If InStr(1, dbsheetConnString, dbPwdSpec) > 0 Then
                dbsheetConnString = Change(dbsheetConnString, dbPwdSpec, Password.Text, ";")
            Else
                dbsheetConnString = dbsheetConnString & ";" & dbPwdSpec & Password.Text
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
            ErrorMsg("Error connecting to DB: " & ex.Message & ", connection string: " & dbsheetConnString, "Open Connection Error")
            dbshcnn = Nothing
        End Try
    End Function

    ''' <summary>checks in existing columns (column 1 in DataGridView theCols) whether theColumnVal exists already in DBsheetCols And returns the found row of DBsheetCols</summary>
    ''' <param name="theColumnVal"></param>
    ''' <returns>found row in DataGridView</returns>
    Public Function checkForValue(theColumnVal As String) As Integer
        DBSheetCols.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Try
            For Each row As DataGridViewRow In DBSheetCols.Rows
                If (row.Cells(2).Value.ToString().Equals(theColumnVal)) Then Return row.Index
            Next
        Catch ex As Exception
            ErrorMsg(ex.Message)
        End Try
        Return -1
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
        Return Strings.Left(theString, replaceBeg - 1) & changed & Strings.Right(theString, Len(theString) - replaceBeg - Len(keystr) + 1)
    End Function

    ''' <summary>set UI to enable(choice=True)/disable(choice=False) changes of table</summary>
    ''' <param name="choice"></param>
    Private Sub TableEditable(ByRef choice As Boolean)
        Database.Enabled = choice
        LDatabase.Enabled = choice
        Table.Enabled = choice
        LTable.Enabled = choice
    End Sub

    ''' <summary>replaces tblPlaceHolder with changed in theString, quote aware (keystr is not replaced within quotes) !!</summary>
    ''' <param name="theString"></param>
    ''' <param name="changed"></param>
    ''' <returns>the replaced string</returns>
    Private Function quotedReplace(ByVal theString As String, ByVal changed As String) As String
        Dim teststr
        Dim subresult As String
        quotedReplace = ""
        teststr = Split(theString, "'")
        ' walk through quote1 splitted parts and replace keystr in even ones
        For i As Integer = 0 To UBound(teststr)
            If i Mod 2 = 0 Then
                subresult = Replace(teststr(i), tblPlaceHolder, changed)
            Else
                subresult = teststr(i)
            End If
            quotedReplace += subresult + IIf(i < UBound(teststr), "'", vbNullString)
        Next
    End Function

    Private Sub DBSheetCreateForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Dim env As Integer = Globals.selectedEnvironment + 1
        dbGetAllStr = fetchSetting("dbGetAll" & env.ToString, "NONEXISTENT")
        If dbGetAllStr = "NONEXISTENT" Then
            ErrorMsg("No dbGetAllStr given for environment: " & env.ToString & ", please correct and rerun.", "createDefinitions Error")
            Exit Sub
        End If
        DBGetAllFieldName = fetchSetting("dbGetAllFieldName" & env.ToString, "NONEXISTENT")
        If DBGetAllFieldName = "NONEXISTENT" Then
            ErrorMsg("No DBGetAllFieldName given for environment: " & env.ToString & ", please correct and rerun.", "createDefinitions Error")
            Exit Sub
        End If
        ownerQualifier = fetchSetting("ownerQualifier" & env.ToString, "NONEXISTENT")
        If ownerQualifier = "NONEXISTENT" Then
            ErrorMsg("No ownerQualifier given for environment: " & env.ToString & ", please correct and rerun.", "createDefinitions Error")
            Exit Sub
        End If
        dbsheetConnString = fetchSetting("DBSheetConnString" & env.ToString, "NONEXISTENT")
        If dbsheetConnString = "NONEXISTENT" Then
            ErrorMsg("No Connectionstring given for environment: " & env.ToString & ", please correct and rerun.", "createDefinitions Error")
            Exit Sub
        End If
        dbidentifier = fetchSetting("DBidentifierCCS" & env.ToString, "NONEXISTENT")
        If dbidentifier = "NONEXISTENT" Then
            ErrorMsg("No DB identifier given for environment: " & env.ToString & ", please correct and rerun.", "createDefinitions Error")
            Exit Sub
        End If
        dbPwdSpec = fetchSetting("dbPwdSpec" & env.ToString, "")
        tblPlaceHolder = fetchSetting("tblPlaceHolder" & env.ToString, "!T!")
        specialNonNullableChar = fetchSetting("specialNonNullableChar" & env.ToString, "*")

        ' columns for DBSheetCols
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

        If dbPwdSpec <> "" Then
            Me.Text = "DB Sheet creation: Please enter required Password into Pwd to access schema information"
            TableEditable(False)
            saveEnabled(False)
            loadDefs.Enabled = False
        Else
            fillDatabasesAndSetDropDown()
            ' initialize with empty DBSheet definitions is done by above call, changing Database.SelectedIndex (Database_SelectedIndexChanged)
        End If
    End Sub

    ''' <summary>enter is hit in Password textbox triggering initialisation</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Password_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Password.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            fillDatabasesAndSetDropDown()
        End If
    End Sub

    ''' <summary>fill the Database dropdown and set to database set in connection string</summary>
    Private Sub fillDatabasesAndSetDropDown()
        Try
            fillDatabases(Database)
        Catch ex As System.Exception
            ErrorMsg("Error: " & ex.Message)
            Exit Sub
        End Try
        Me.Text = "DB Sheet creation: Select Database and Table to start building a DBSheet Definition"
        Database.SelectedIndex = Database.Items.IndexOf(fetch(dbsheetConnString, dbidentifier, ";"))
        'initialization of everything else is triggered by above change and caught by Database_SelectedIndexChanged
    End Sub

    ''' <summary>database changed, initialize everything from scratch</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Database_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Database.SelectedIndexChanged
        createDefinitions()
    End Sub

    Private Sub DBSheetCols_RowStateChanged(sender As Object, e As DataGridViewRowStateChangedEventArgs) Handles DBSheetCols.RowStateChanged
        If e.StateChanged = DataGridViewElementStates.Selected Then
            columnEditMode(True)
        Else
            columnEditMode(False)
        End If
    End Sub
End Class