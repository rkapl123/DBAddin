Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.Odbc
Imports System.IO
Imports System.Text
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
    Private currentForDatabase As String = ""
    Private currentForTable As String = ""
    Private curConfig As String
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
            init_Form()
            ' loading defs from file...
            If Strings.Len(DBSheetParams) > 0 Then curConfig = DBSheetParams
            FormDisabled = True

            ' if we have a valid dbsheet definition (either selected a valid dbsheeet or loaded from file)
            ' fetch params into form from sheet or file
            If Not (curConfig = "") Then
                FormDisabled = True
                ' get Database from (legacy) connID (legacy connID prefixed with connIDPrefixDBtype)
                Dim configDatabase As String = Replace(DBSheetConfig.getEntry("connID", DBSheetParams), fetchSetting("connIDPrefixDBtype", "MSSQL"), "")
                Database.SelectedIndex = Database.Items.IndexOf(configDatabase)
                If Not openConnection(Database.Text) Then
                    ErrorMsg("Couldn't open connection to database " & Database.Text)
                    Exit Sub
                End If
                ExcelDnaUtil.Application.StatusBar = "fillTables..."
                fillTables()
                FormDisabled = True
                Table.SelectedIndex = Table.Items.IndexOf(DBSheetConfig.getEntry("table", DBSheetParams))
                ExcelDnaUtil.Application.StatusBar = "fillColumns..."
                fillColumns()
                ExcelDnaUtil.Application.StatusBar = "getEntryList..."
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
                ExcelDnaUtil.Application.StatusBar = "fillDatabases and ForTables..."
                fillDatabases(ForDatabase)
                fillForTables()
                ExcelDnaUtil.Application.StatusBar = False
                DBSheetCols.Enabled = True
                TableEditable(False)
                saveEnabled(True)
            Else
                ' start with empty columns list
                TableEditable(True)
            End If
            ' now show the dialog...
            columnEditMode(False)
        Catch ex As System.Exception
            ErrorMsg("Error: " & ex.Message)
        End Try
    End Sub

    ''' <summary>reusable init procedure</summary>
    Public Sub init_Form()
        Try
            Sorting.Items.Clear()
            Sorting.Items.Add("None")
            Sorting.Items.Add("Ascending")
            Sorting.Items.Add("Descending")
            Query.Text = ""
            WhereClause.Text = ""
            LookupQuery.Text = ""
            DBSheetCols.Rows.Clear()
            currentForTable = ""
            maxColCount = 0
            saveEnabled(False)

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
            If Strings.Len(LookupQuery.Text) > 0 Then
                If testLookupQuery.Text = "test &Lookup Query" Then
                    testTheQuery(LookupQuery.Text, True)
                ElseIf testLookupQuery.Text = "remove &Lookup Testsheet" Then
                    ' TODO: check for lookup testsheet...
                    If (testcheck.IndexOf("TESTSHEET") + 1) = 0 Then
                        ErrorMsg("Active sheet doesn't seem to be a query test sheet !!!", "DBSheet Testsheet Remove Warning")
                    Else
                        ExcelDnaUtil.Application.ActiveWorkbook.Close(False)
                    End If
                    testLookupQuery.Text = "test &Lookup Query"
                End If
            Else
                ErrorMsg("No restriction query created to test !!!", "DBSheet Query Test Warning")
            End If

        Catch ex As System.Exception
            ErrorMsg("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub ForDatabase_SelectedIndexChanged(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles ForDatabase.SelectedIndexChanged
        If FormDisabled Then Exit Sub
        fillForTables()
        currentForDatabase = ForDatabase.Text
    End Sub

    Private Sub Table_SelectedIndexChanged(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles Table.SelectedIndexChanged
        If FormDisabled Then Exit Sub
        Try
            FormDisabled = True
            If Table.SelectedIndex >= 0 Then
                addAllFields.Enabled = True
                addToDBsheetCols.Enabled = True
                DBSheetCols.Enabled = True
            End If
            ' just in case this wasn't cleared before...
            'TODO: change theDBSheetColumnList
            'theDBSheetColumnList.Clear()
            Query.Text = ""
            fillColumns()
            columnEditMode(False)
            FormDisabled = False

        Catch ex As System.Exception
            ErrorMsg("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub isPrimary_CheckStateChanged(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles IsPrimary.CheckStateChanged
        If FormDisabled Or Strings.Len(Column.Text) = 0 Then Exit Sub
        Try
            lookupAndSelect(IsPrimary)
            'TODO: change theDBSheetColumnList
            'If Not theDBSheetColumnList.hasRows() Then Exit Sub
            'If Not theDBSheetColumnList.selectionMade() Then Exit Sub
            'If Not theDBSheetColumnList.firstEntrySelected() Then
            '    If theDBSheetColumnList.Value(theDBSheetColumnList.Selection - 1, 5) = 0 And IsPrimary.CheckState = CheckState.Checked Then
            '        ErrorMsg("All primary keys have to be first and there is at least one non-primary key column before that one !", "DBSheet Definition Warning")
            '        IsPrimary.CheckState = CheckState.Unchecked
            '    End If
            '    If Not theDBSheetColumnList.lastEntrySelected() Then
            '        If theDBSheetColumnList.Value(theDBSheetColumnList.Selection + 1, 5) = 1 And IsPrimary.CheckState = CheckState.Unchecked Then
            '            ErrorMsg("All primary keys have to be first and there is at least one primary key column after that one !", "DBSheet Definition Warning")
            '            IsPrimary.CheckState = CheckState.Checked
            '        End If
            '    End If
            'ElseIf IsPrimary.CheckState = CheckState.Unchecked Then
            '    ErrorMsg("first column always has to be primary key", "DBSheet Definition Warning")
            '    IsPrimary.CheckState = CheckState.Checked
            'End If

        Catch ex As System.Exception
            ErrorMsg("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub isForeign_CheckStateChanged(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles IsForeign.CheckStateChanged
        If FormDisabled Or Strings.Len(Column.Text) = 0 Then Exit Sub
        Try
            lookupAndSelect(IsForeign)
            ' check whether this can't be done also on non selected fields (would be nicer !!)....
            'TODO: change  theDBSheetColumnList
            'If Not theDBSheetColumnList.selectionMade() Then Exit Sub
            fillDatabases(ForDatabase)
            ForDatabase.SelectedIndex = ForDatabase.Items.IndexOf(Database.Text)
            setForeignColFieldsVisibility()
            LLookupQuery.Enabled = True
            LookupQuery.Enabled = True
            regenLookupQueries.Enabled = True
            testLookupQuery.Enabled = True
        Catch ex As System.Exception
            ErrorMsg("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub ForTable_SelectedIndexChanged(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles ForTable.SelectedIndexChanged
        foreignTableChange()
    End Sub

    Private Sub Column_SelectedIndexChanged(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles Column.SelectedIndexChanged
        If FormDisabled Then Exit Sub
        TableEditable(False)
        FormDisabled = True
        IsPrimary.CheckState = CheckState.Unchecked
        IsForeign.CheckState = CheckState.Unchecked
        FormDisabled = False
    End Sub

    Private Sub addAllFields_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles addAllFields.Click
        Dim rstSchema As DataSet

        Try
            FormDisabled = True
            rstSchema = dbshcnn.GetSchema().DataSet
            Dim firstRow As Boolean : firstRow = True
            'TODO: change theDBSheetColumnList
            'theDBSheetColumnList.Clear()
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

    Private Sub addToDBsheetCols_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles addToDBsheetCols.Click
        Try
            If Strings.Len(Column.Text) = 0 Then Exit Sub
            If addToDBsheetCols.Text.StartsWith("&abort") Then
                columnEditMode(False)
                Exit Sub
            End If

            If maxColCount = 0 Then
                maxColCount = ExcelDnaUtil.ExcelLimits.MaxColumns
            End If
            'TODO: change theDBSheetColumnList
            'If theDBSheetColumnList.RowCount = maxColCount Then
            '    ErrorMsg("Max. Columns allowed in DBSheet: " & maxColCount & " (last column reserved for data status)", "DBSheet Definition Warning")
            '    Exit Sub
            'End If

            ExcelDnaUtil.Application.EnableEvents = False
            Dim newRow As Integer
            newRow = DBSheetCols.Rows.Add(New DataGridViewRow())
            FormDisabled = True

            ' Column
            'TODO: change theDBSheetColumnList
            'theDBSheetColumnList.Value(newRow, 0) = Column.Text

            ' Foreign Table information
            If Strings.Len(ForTable.Text) > 0 And Strings.Len(ForTableKey.Text) > 0 And Strings.Len(ForTableLookup.Text) > 0 Then
                'TODO: change theDBSheetColumnList
                'theDBSheetColumnList.Value(newRow, 1) = ForDatabase.Text & ownerQualifier & ForTable.Text
                'theDBSheetColumnList.Value(newRow, 2) = ForTableKey.Text
                'theDBSheetColumnList.Value(newRow, 3) = ForTableLookup.Text
                'If outerJoin.CheckState = CheckState.Checked Then theDBSheetColumnList.Value(newRow, 4) = 1
            ElseIf Strings.Len(ForTable.Text) > 0 Or Strings.Len(ForTableKey.Text) > 0 Or Strings.Len(ForTableLookup.Text) > 0 And Strings.Len(LookupQuery.Text) = 0 Then
                ErrorMsg("Please specify all 3 foreign column informations: ForeignTable, ForeignTableKey and ForeignTableLookup !", "DBSheet Definition Warning")
            End If

            ' Primary key
            If newRow = 0 Then ' always have first column as PK
                'TODO: change theDBSheetColumnList
                'theDBSheetColumnList.Value(newRow, 5) = 1
                IsPrimary.CheckState = CheckState.Checked
            End If
            ' check if primary keys are first
            Dim primaryAllowed As Boolean
            primaryAllowed = True
            For i As Integer = 0 To newRow
                'TODO: change theDBSheetColumnList
                'If Strings.Len(theDBSheetColumnList.Value(i, 5)) = 0 Then
                '    primaryAllowed = False
                '    Exit For
                'End If
            Next
            If IsPrimary.CheckState = CheckState.Checked Then
                If primaryAllowed Then
                    'TODO: change theDBSheetColumnList
                    'theDBSheetColumnList.Value(newRow, 5) = 1
                Else
                    MessageBox.Show("Primary Keys must be first in a DBSheet (please place above)", "DBAddin: DBSheet Definition Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    IsPrimary.CheckState = CheckState.Unchecked
                End If
            End If

            ' Column Type
            'TODO: change theDBSheetColumnList
            'theDBSheetColumnList.Value(newRow, 6) = getType_Renamed(Column.Text)

            ' Sort by this column ?
            'TODO: change theDBSheetColumnList
            'theDBSheetColumnList.Value(newRow, 7) = Sorting.Text

            columnEditMode(False)
            FormDisabled = False
            ExcelDnaUtil.Application.EnableEvents = True
            TableEditable(False) ' after changing the column no more change to table allowed !!

        Catch ex As System.Exception
            ExcelDnaUtil.Application.EnableEvents = True
            ErrorMsg("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub removeDBsheetCols_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles removeDBsheetCols.Click
        Try
            ExcelDnaUtil.Application.EnableEvents = False
            FormDisabled = True
            ' TODO: change theDBSheetColumnList
            'If theDBSheetColumnList.selectionMade() Then
            '    'TODO: theDBSheetColumnList.removeRow(theDBSheetColumnList.Selection)
            '    setEntryFields()
            'End If
            'If Not theDBSheetColumnList.hasRows() Then
            '    Query.Text = ""
            '    ' reset the current filename
            '    currentFilepath = ""
            '    saveEnabled(False)
            '    columnEditMode(False)
            '    ' after resetting columns changes to table/connection allowed again !!
            '    TableEditable(True)
            'End If
            FormDisabled = False
            ExcelDnaUtil.Application.EnableEvents = True

        Catch ex As System.Exception
            ExcelDnaUtil.Application.EnableEvents = True
            ErrorMsg("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub clearAllFields_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles clearAllFields.Click
        clearTablesColumnsAndQuery()
        ' reset the current filename
        currentFilepath = ""
        saveEnabled(False)
        columnEditMode(False)
    End Sub

    ''
    '  when entering into DBSheetCols then start editing the DBlookup column list
    Private Sub DBsheetCols_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles DBSheetCols.Click
        If FormDisabled Then Exit Sub

        setColumnListFields()
        columnEditMode(True)
        FormDisabled = True
        setEntryFields()
        FormDisabled = False
    End Sub

    ''
    ' copy/paste is implemented for DBsheet foreign key/primkey/calculated/lookup definitions
    ' @param KeyCode
    ' @param Shift
    Private Sub DBsheetCols_KeyDown(ByVal eventSender As Object, ByVal eventArgs As KeyEventArgs) Handles DBSheetCols.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData / 65536
        Try
            CtrlPressed = (Shift And 2) > 0
        Finally
            eventArgs.Handled = KeyCode = 0
        End Try
    End Sub

    ''
    ' @param KeyAscii
    Private Sub DBsheetCols_KeyPress(ByVal eventSender As Object, ByVal eventArgs As KeyPressEventArgs) Handles DBSheetCols.KeyPress
        Dim KeyAscii As Integer = Strings.Asc(eventArgs.KeyChar)
        Static restrictDef As String = "", PrimaryV As String = "", ForTableLookupV As String = "", ForTableV As String = "", ForTableKeyV As String = "", outerJoinV As String = "", isCalculatedV As String = "", SortingBy As String = ""

        Try
            If addToDBsheetCols.Text.StartsWith("&add") Or Not CtrlPressed Then
                If KeyAscii = 0 Then
                    eventArgs.Handled = True
                End If
                Exit Sub
            End If

            Dim curSel As Integer
            'TODO: change theDBSheetColumnList
            'curSel = theDBSheetColumnList.Selection
            ' copy into static variables
            If KeyAscii = 3 Then
                'TODO: change theDBSheetColumnList
                'ForTableV = theDBSheetColumnList.Value(curSel, 1)
                'ForTableKeyV = theDBSheetColumnList.Value(curSel, 2)
                'ForTableLookupV = theDBSheetColumnList.Value(curSel, 3)
                'outerJoinV = theDBSheetColumnList.Value(curSel, 4)
                'PrimaryV = theDBSheetColumnList.Value(curSel, 5)
                'SortingBy = theDBSheetColumnList.Value(curSel, 7)
                'restrictDef = theDBSheetColumnList.Value(curSel, 8)
                ' paste from static variables
            ElseIf KeyAscii = 22 Then
                'TODO: change theDBSheetColumnList
                'theDBSheetColumnList.Value(curSel, 1) = ForTableV
                'theDBSheetColumnList.Value(curSel, 2) = ForTableKeyV
                'theDBSheetColumnList.Value(curSel, 3) = ForTableLookupV
                'theDBSheetColumnList.Value(curSel, 4) = outerJoinV
                'theDBSheetColumnList.Value(curSel, 5) = PrimaryV
                '' exception if we overwrite isPrimary of first dbsheet column...
                'If curSel = 0 Then theDBSheetColumnList.Value(curSel, 5) = "Y"
                'theDBSheetColumnList.Value(curSel, 7) = SortingBy
                'theDBSheetColumnList.Value(curSel, 8) = restrictDef
                setEntryFields()
            End If
            If KeyAscii = 0 Then
                eventArgs.Handled = True
            End If

        Catch ex As System.Exception
            ErrorMsg("Error: " & ex.Message)
            If KeyAscii = 0 Then
                eventArgs.Handled = True
            End If
            eventArgs.KeyChar = Convert.ToChar(KeyAscii)
        End Try
    End Sub

    Private Sub regenLookupQueries_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles regenLookupQueries.Click
        Try
            FormDisabled = True
            Dim retval As MsgBoxResult
            If regenLookupQueries.Text = "re&generate this lookup query" Then
                LookupQuery.Text = "SELECT " & ForTableLookup.Text & "," & ForTableKey.Text & " FROM " & ForDatabase.Text & ownerQualifier & ForTable.Text & " ORDER BY " & ForTableLookup.Text
            Else
                retval = QuestionMsg("regenerating Foreign Lookups completely (overwriting all customizations there): yes or generate only new: no !", MsgBoxStyle.YesNoCancel, "DBSheet Definition")
                If retval = MsgBoxResult.Cancel Then
                    FormDisabled = False
                    Exit Sub
                End If
                'TODO: change theDBSheetColumnList
                'For i As Integer = 0 To theDBSheetColumnList.RowCount - 1
                '    If Strings.Len(theDBSheetColumnList.Value(i, 1)) > 0 Then
                '        ' only overwrite if forced regenerate or empty restriction def...
                '        If retval = MsgBoxResult.Yes Or Strings.Len(theDBSheetColumnList.Value(i, 9)) = 0 Then
                '            theDBSheetColumnList.Value(i, 9) = "SELECT " & theDBSheetColumnList.Value(i, 3) & "," & theDBSheetColumnList.Value(i, 2) & " FROM " & theDBSheetColumnList.Value(i, 1) & " ORDER BY " & theDBSheetColumnList.Value(i, 3)
                '        End If
                '    End If
                'Next
            End If
            FormDisabled = False
        Catch ex As System.Exception
            ErrorMsg("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub moveUp_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles moveUp.Click
        'TODO: change theDBSheetColumnList
        'If Not theDBSheetColumnList.selectionMade() Or theDBSheetColumnList.firstEntrySelected() Then Exit Sub
        Try
            If DBSheetCols.CurrentCell.RowIndex = 0 And IsPrimary.CheckState = CheckState.Unchecked Then
                ErrorMsg("first column always has to be primary key", "DBSheet Definition Warning")
                Exit Sub
            ElseIf DBSheetCols.DataSource.Item(DBSheetCols.CurrentCell.RowIndex + 1).Primary And IsPrimary.CheckState = CheckState.Unchecked Then
                ErrorMsg("All primary keys have to be first and there is a primary key column that would be shifted below this non-primary one !", "DBSheet Definition Warning")
                Exit Sub
            End If
            If IsNothing(DBSheetCols.CurrentRow) Then Return
            Dim rw As DataGridViewRow = DBSheetCols.CurrentRow
            Dim selIndex As Integer = DBSheetCols.CurrentRow.Index
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

    Private Sub moveDown_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles moveDown.Click
        Try
            'TODO: change theDBSheetColumnList
            'If Not theDBSheetColumnList.selectionMade() Or theDBSheetColumnList.lastEntrySelected() Then Exit Sub
            If Not DBSheetCols.DataSource.Item(DBSheetCols.CurrentCell.RowIndex + 1).Primary And IsPrimary.CheckState = CheckState.Checked Then
                ErrorMsg("All primary keys have to be first and there is a non primary key column that would be shifted above this primary one !", "DBSheet Definition Warning")
                Exit Sub
            End If
            If IsNothing(DBSheetCols.CurrentRow) Then Return
            Dim rw As DataGridViewRow = DBSheetCols.CurrentRow
            Dim selIndex As Integer = DBSheetCols.CurrentRow.Index
            ' avoid moving down of last row, DBSeqenceDataGrid.Rows.Count is 1 more than the actual inserted rows because of the "new" row, selIndex is 0 based
            If selIndex = DBSheetCols.Rows.Count - 2 Then Return
            DBSheetCols.Rows.RemoveAt(selIndex)
            DBSheetCols.Rows.Insert(selIndex + 1, rw)
            DBSheetCols.Rows(selIndex + 1).Cells(0).Selected = True
            last += 1
            columnEditMode(True)

        Catch ex As System.Exception
            ErrorMsg("Error: " & ex.Message)
        End Try
    End Sub

    ''' <summary>clears the defined columns and resets the selection fields (Table, ForTable) and the Query</summary>
    Private Sub clearTablesColumnsAndQuery()
        FormDisabled = True
        'TODO: change theDBSheetColumnList
        ' theDBSheetColumnList.Clear()
        TableEditable(True)
        Table.SelectedIndex = -1
        Column.SelectedIndex = -1
        Query.Text = ""
        LookupQuery.Text = ""
        columnEditMode(False)
        FormDisabled = False
    End Sub

    ''' <summary>lookup and set last to existing column, if it doesn't exist, set last to end of DBSheetCols list (used for isCalculated, isForeign and isPrimary changing)</summary>
    ''' <param name="changedField"></param>
    Private Sub lookupAndSelect(ByRef changedField As CheckBox)
        Try
            Dim columnBackup As String = ""

            last = checkForValue(Column.Text)
            If addToDBsheetCols.Text.StartsWith("&add") And last >= 0 Then
                FormDisabled = True
                'TODO: change theDBSheetColumnList
                ' theDBSheetColumnList.Selection = last
                columnBackup = CStr(changedField.CheckState)
                setEntryFields()
                changedField.CheckState = columnBackup
                columnEditMode(True)
            Else
                'last = theDBSheetColumnList.Selection
            End If

        Catch ex As System.Exception
            ErrorMsg("Error: " & ex.Message)
        End Try
    End Sub

    ''' <summary>sets the column list fields to the values in the entry fields (on top of the col list)</summary>
    Private Sub setColumnListFields()
        Try
            If addToDBsheetCols.Text.StartsWith("&abort") Then
                ' only fill column def if column is really filled (schema errors can lead to empty values here !
                'TODO: change theDBSheetColumnList
                'If Strings.Len(Column.Text) > 0 Then theDBSheetColumnList.Value(last, 0) = Column.Text
                'theDBSheetColumnList.Value(last, 5) = IsPrimary.CheckState
                'theDBSheetColumnList.Value(last, 7) = Sorting.Text
                'theDBSheetColumnList.Value(last, 8) = LookupQuery.Text
                'If IsForeign.CheckState = CheckState.Checked And Strings.Len(ForTable.Text) > 0 And Strings.Len(ForTableKey.Text) > 0 And Strings.Len(ForTableLookup.Text) > 0 Then
                '    theDBSheetColumnList.Value(last, 1) = ForDatabase.Text & ownerQualifier & ForTable.Text
                '    theDBSheetColumnList.Value(last, 2) = ForTableKey.Text
                '    theDBSheetColumnList.Value(last, 3) = ForTableLookup.Text
                '    theDBSheetColumnList.Value(last, 4) = outerJoin.CheckState
                'Else
                '    theDBSheetColumnList.Value(last, 1) = ""
                '    theDBSheetColumnList.Value(last, 2) = ""
                '    theDBSheetColumnList.Value(last, 3) = ""
                '    theDBSheetColumnList.Value(last, 4) = 0
                'End If
            End If

        Catch ex As System.Exception
            ErrorMsg("Error: " & ex.Message)
        End Try
    End Sub

    ''' <summary>sets the entry fields on top of the DBSheet col list</summary>
    Private Sub setEntryFields()
        Dim ForTableKeyV, ForTableV, ForTableLookupV As String

        Try
            ' remember last selection for a) moveRows and b) lookupAndSelect
            'TODO: change theDBSheetColumnList
            'last = theDBSheetColumnList.Selection
            If last = -1 Then Exit Sub

            ' lookup the current selected row in the available columns
            Dim newIndex As Integer
            'TODO: change theDBSheetColumnList
            'newIndex = Column.Items.IndexOf(theDBSheetColumnList.Value(last, 0))
            ' if column name changed isnullable flag (specialNonNullableChar in front of col name changed)
            ' then try again including/excluding it (prevent strange GUI effects later)
            'If newIndex = -1 Then
            '    If theDBSheetColumnList.Value(last, 0).StartsWith(specialNonNullableChar) Then
            '        ' skip specialNonNullableChar in front
            '        newIndex = Column.Items.IndexOf(theDBSheetColumnList.Value(last, 0).Substring(1))
            '    Else
            '        ' add specialNonNullableChar in front
            '        newIndex = Column.Items.IndexOf(specialNonNullableChar & theDBSheetColumnList.Value(last, 0))
            '    End If
            'End If
            'If newIndex = -1 Then
            '    ErrorMsg("couldn't find the column " & theDBSheetColumnList.Value(last, 0) & " in current table's columns. Did the table schema change?")
            'End If
            Column.SelectedIndex = newIndex

            ' set the plain entry fields to the row's values
            'TODO: change theDBSheetColumnList
            'ForTableV = theDBSheetColumnList.Value(last, 1)
            'ForTableKeyV = theDBSheetColumnList.Value(last, 2)
            'ForTableLookupV = theDBSheetColumnList.Value(last, 3)
            'outerJoin.CheckState = theDBSheetColumnList.Value(last, 4)
            'IsPrimary.CheckState = theDBSheetColumnList.Value(last, 5)
            'Sorting.SelectedIndex = Sorting.Items.IndexOf(theDBSheetColumnList.Value(last, 7))
            'LookupQuery.Text = theDBSheetColumnList.Value(last, 8)

            ' set the foreign lookup entry fields to the row's values: special care needs to be taken for switching databases !!
            If Strings.Len(ForTableV) > 0 Then
                IsForeign.CheckState = CheckState.Checked
                setForeignColFieldsVisibility()
                ' in case of qualified table name (pubs.dbo.employee), set ForDatabase to end at first "."
                If InStr(1, ForTableV, ".") = 0 Then
                    ErrorMsg("No database information can be extracted from (not fully qualified) foreign table name!")
                    ForDatabase.SelectedIndex = -1
                Else
                    ForDatabase.SelectedIndex = ForDatabase.Items.IndexOf(Strings.Left(ForTableV, InStr(1, ForTableV, ".") - 1))
                End If
                If ForDatabase.Text <> currentForDatabase Then
                    fillForTables()
                    currentForDatabase = ForDatabase.Text
                End If
                ' in case of qualified table name (pubs.dbo.employee), set table to begin at last "."
                Dim lookupForTable As String = Strings.Mid(ForTableV, IIf(InStrRev(ForTableV, ".") = 0, 1, InStrRev(ForTableV, ".") + 1))
                ForTable.SelectedIndex = ForTable.Items.IndexOf(lookupForTable)
                If ForTable.SelectedIndex = -1 Then
                    ErrorMsg("foreign table '" & lookupForTable & "' was not found! Did the table's name change (case sensitive !)?")
                    Exit Sub
                End If
                ' update foreign column dropdowns, try to reassign existing values ...
                ForceFieldUpdate = True
                foreignTableChange(ForTableKeyV, ForTableLookupV)
                ForceFieldUpdate = False
            Else
                IsForeign.CheckState = CheckState.Unchecked
                setForeignColFieldsVisibility()
            End If

        Catch ex As System.Exception
            ErrorMsg("Error: " & ex.Message)
        End Try
    End Sub

    ''' <summary>sets the entry foreign lookup and id fields depending on value in foreign table</summary>
    ''' <param name="oldForTableKey"></param>
    ''' <param name="oldForTableLookup"></param>
    Private Sub foreignTableChange(Optional ByRef oldForTableKey As String = "", Optional ByRef oldForTableLookup As String = "")
        Dim rstSchema As DataSet

        Try
            If FormDisabled And Not ForceFieldUpdate Then Exit Sub
            If ForTable.Text <> currentForTable Then
                ForTableKey.Items.Clear()
                ForTableLookup.Items.Clear()
                rstSchema = dbshcnn.GetSchema().DataSet
                For Each iteration_row As DataRow In rstSchema.Tables(0).Rows
                    ForTableKey.Items.Add(iteration_row("COLUMN_NAME"))
                    ForTableLookup.Items.Add(iteration_row("COLUMN_NAME"))
                Next iteration_row
                currentForTable = ForTable.Text
            End If
            ' restore backuped settings
            Dim newIndex As Integer
            If Strings.Len(oldForTableKey) > 0 Then
                newIndex = ForTableKey.Items.IndexOf(oldForTableKey)
                If newIndex = -1 Then
                    ErrorMsg("couldn't find the foreign table key column " & oldForTableKey & " in current foreign table columns. Did the table schema change?")
                Else
                    ForTableKey.SelectedIndex = newIndex
                End If
            End If
            If Strings.Len(oldForTableLookup) > 0 Then
                newIndex = ForTableLookup.Items.IndexOf(oldForTableLookup)
                If newIndex = -1 Then
                    ErrorMsg("couldn't find the foreign table lookup column " & oldForTableLookup & " in current foreign table columns. Did the table schema change?")
                Else
                    ForTableLookup.SelectedIndex = newIndex
                End If
            End If
        Catch ex As System.Exception
            ErrorMsg("Error: " & ex.Message)
        End Try
    End Sub

    ''' <summary>add to DBsheet/abort column edit, regenerate this/all lookups, visible restrictions (moveUp/Down, remove Cols, )</summary>
    ''' <param name="choice"></param>
    ''' <remarks>sets(choice=true) or resets(choice=false) column "edit" mode</remarks>
    Private Sub columnEditMode(ByRef choice As Boolean)
        FormDisabled = True
        removeDBsheetCols.Enabled = choice
        moveDown.Visible = choice
        moveUp.Visible = choice
        If choice Then
            Column.Enabled = False
            LColumn.Enabled = False
            addToDBsheetCols.Text = "&abort column edit"
            regenLookupQueries.Text = "re&generate this lookup query"
        Else
            Column.Enabled = True
            LColumn.Enabled = True
            'TODO: change theDBSheetColumnList
            'theDBSheetColumnList.Selection = -1
            addToDBsheetCols.Text = "&add to DBsheet"
            regenLookupQueries.Text = "re&generate all lookup queries"
            IsForeign.CheckState = CheckState.Unchecked
            setForeignColFieldsVisibility()
        End If
        FormDisabled = False
    End Sub

    ''' <summary>shows/hides additional foreign column entry fields (when isForeign is checked)</summary>
    Private Sub setForeignColFieldsVisibility()
        If IsForeign.CheckState = CheckState.Checked Then
            ForDatabase.Visible = True
            ForTable.Visible = True
            ForTableKey.Visible = True
            ForTableLookup.Visible = True
            LForDatabase.Visible = True
            LForTable.Visible = True
            LForTableKey.Visible = True
            LForTableLookup.Visible = True
            outerJoin.Visible = True
        Else
            ForDatabase.Visible = False
            LForDatabase.Visible = False
            ForTable.SelectedIndex = -1
            ForTableKey.SelectedIndex = -1
            ForTableLookup.SelectedIndex = -1
            ForTable.Visible = False
            ForTableKey.Visible = False
            ForTableLookup.Visible = False
            LForTable.Visible = False
            LForTableKey.Visible = False
            LForTableLookup.Visible = False
            outerJoin.CheckState = CheckState.Unchecked
            outerJoin.Visible = False
        End If
    End Sub

    Private TableDataTypes As Dictionary(Of String, String)

    ''' <summary>gets the types of currently selected table including size, precision and scale into DataTypes</summary>
    Private Sub getTableDataTypes()
        TableDataTypes = New Dictionary(Of String, String)
        Dim rstSchema As OdbcDataReader
        If Not openConnection() Then Exit Sub

        Dim sqlCommand As OdbcCommand = New OdbcCommand("SELECT TOP 1 * FROM " + Table.Text, dbshcnn)
        rstSchema = sqlCommand.ExecuteReader()
        Try
            Dim schemaInfo As DataTable = rstSchema.GetSchemaTable()
            For Each schemaRow As DataRow In schemaInfo.Rows
                Dim appendInfo As String = If(schemaRow("AllowDBNull"), "", specialNonNullableChar)
                TableDataTypes(appendInfo + schemaRow("ColumnName")) = schemaRow("DataType").Name + "(" + schemaRow("ColumnSize").ToString + "/" + schemaRow("NumericPrecision").ToString + "/" + schemaRow("NumericScale").ToString + ")"
            Next
        Catch ex As Exception
            ErrorMsg("Could not get type information for table: '" & Table.Text & "', error: " & ex.Message)
        End Try
        rstSchema.Close()
    End Sub

    ''' <summary>fill all possible columns of currently selected table</summary>
    Private Sub fillColumns()
        getTableDataTypes()
        Dim columnTemp As String = Column.Text
        Column.Items.Clear()
        Try
            For Each colname As String In TableDataTypes.Keys
                Column.Items.Add(colname)
            Next
            If Strings.Len(columnTemp) > 0 Then Column.SelectedIndex = Column.Items.IndexOf(columnTemp)
            FormDisabled = False
        Catch ex As System.Exception
            Throw New Exception("Exception in fillColumns: " & ex.Message)
        End Try

        'Dim schemaTable As DataTable
        'Dim attached As String

        'If Not openConnection() Then Exit Sub
        'FormDisabled = True

        'Try
        '    schemaTable = dbshcnn.GetSchema("Columns")
        '    If schemaTable.Rows.Count = 0 Then Throw New Exception("No Columns could be fetched from Schema")
        'Catch ex As Exception
        '    FormDisabled = False
        '    Throw New Exception("Error getting schema information for columns in connection strings database ' " & Database.Text & "'." & ",error: " & ex.Message)
        'End Try

        'Try
        '    For Each iteration_row As DataRow In schemaTable.Rows
        '        attached = ""
        '        If iteration_row("IS_NULLABLE") <> "NO" Then attached = specialNonNullableChar
        '        If iteration_row("TABLE_CAT").ToUpper() = Database.Text Or iteration_row("TABLE_SCHEM").ToUpper() = Database.Text Then Column.Items.Add(attached & iteration_row("COLUMN_NAME"))
        '    Next iteration_row
        '    If Strings.Len(columnTemp) > 0 Then Column.SelectedIndex = Column.Items.IndexOf(columnTemp)
        '    FormDisabled = False
        'Catch ex As System.Exception
        '    Throw New Exception("Exception in fillColumns: " & ex.Message)
        'End Try
    End Sub

    ''' <summary>fill all possible tables Of currently selected database/schema</summary>
    Private Sub fillTables()
        Dim schemaTable As DataTable
        Dim tableTemp As String

        If Not openConnection() Then
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
                If iteration_row("TABLE_CAT").ToUpper() = Database.Text Or iteration_row("TABLE_SCHEM").ToUpper() = Database.Text Then Table.Items.Add(iteration_row("TABLE_NAME"))
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

        If Not openConnection(ForDatabase.Text) Then
            FormDisabled = False
            ForTable.Items.Clear() : ForTableKey.Items.Clear() : ForTableLookup.Items.Clear()
            Throw New Exception("could not open connection for foreign database '" & ForDatabase.Text & "' in connection string '" & dbsheetConnString & "'.")
        End If

        FormDisabled = True
        tableTemp = ForTable.Text
        ForTable.Items.Clear()
        Try
            schemaTable = dbshcnn.GetSchema("Tables")
            If schemaTable.Rows.Count = 0 Then Throw New Exception("No Tables could be fetched from Schema")
        Catch ex As Exception
            FormDisabled = False
            Throw New Exception("Error getting schema information for tables in connection strings database ' " & Database.Text & "'." & ",error: " & ex.Message)
        End Try
        Try
            For Each iteration_row As DataRow In schemaTable.Rows
                If iteration_row("TABLE_CAT").ToUpper() = ForDatabase.Text.ToUpper() Or iteration_row("TABLE_SCHEM").ToUpper() = ForDatabase.Text.ToUpper() Then ForTable.Items.Add(iteration_row("TABLE_NAME"))
            Next iteration_row
            If Strings.Len(tableTemp) > 0 Then ForTable.SelectedIndex = ForTable.Items.IndexOf(tableTemp)
            FormDisabled = False
        Catch ex As System.Exception
            Throw New Exception("Exception in fillForTables: " & ex.Message)
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
            'TODO: change theDBSheetColumnList
            'For i As Integer = 0 To theDBSheetColumnList.RowCount - 1
            '    ' plain table field
            '    usedColumn = correctNonNull(theDBSheetColumnList.Value(i, 0))
            '    tableCounter += 1
            '    Select Case theDBSheetColumnList.Value(i, 8)
            '        Case "Ascending" : orderByStr = IIf(orderByStr = "", "", orderByStr & ", ") & CStr(i + 1) & " ASC"
            '        Case "Descending" : orderByStr = IIf(orderByStr = "", "", orderByStr & ", ") & CStr(i + 1) & " DESC"
            '    End Select
            '    If Strings.Len(theDBSheetColumnList.Value(i, 1)) = 0 Then
            '        selectStr = selectStr & "T1." & usedColumn & ", "
            '        ' create (inner or outer) joins for foreign key lookup id
            '    Else
            '        If Strings.Len(theDBSheetColumnList.Value(i, 9)) = 0 Then
            '            theDBSheetColumnList.Selection = i
            '            result = ""
            '            LogWarn("No Lookup Query created for field " & theDBSheetColumnList.Value(i, 0) & ", can't proceed !")
            '            Return result
            '        End If
            '        theTable = "T" & tableCounter
            '        ' either we go for the whole part after the last join
            '        completeJoin = fetch(theDBSheetColumnList.Value(i, 9), "JOIN ", "")
            '        ' or we have a simple WHERE and just "AND" it to the created join
            '        addRestrict = quotedReplace(fetch(theDBSheetColumnList.Value(i, 9), "WHERE ", ""), "T" & tableCounter)

            '        ' remove any ORDER BY clause from additional restrict...
            '        restrPos = addRestrict.ToUpper().LastIndexOf(" ORDER") + 1
            '        If restrPos > 0 Then addRestrict = addRestrict.Substring(0, Math.Min(restrPos - 1, addRestrict.Length))
            '        If Strings.Len(completeJoin) > 0 Then
            '            ' when having the complete join, use additional restriction not for main subtable
            '            addRestrict = ""
            '            ' instead make it an additional condition for the join and replace placeholder with tablealias
            '            completeJoin = quotedReplace(ciReplace(completeJoin, "WHERE", "AND"), "T" & tableCounter)
            '        End If
            '        If theDBSheetColumnList.Value(i, 4) = 1 Then
            '            fromStr += " LEFT JOIN " & Environment.NewLine & theDBSheetColumnList.Value(i, 1) & " " & theTable &
            '                           " ON " & "T1." & usedColumn & " = " & theTable & "." & theDBSheetColumnList.Value(i, 2) & IIf(Strings.Len(addRestrict) > 0, " AND " & addRestrict, "")
            '        Else
            '            fromStr += " INNER JOIN " & Environment.NewLine & theDBSheetColumnList.Value(i, 1) & " " & theTable &
            '                           " ON " & "T1." & usedColumn & " = " & theTable & "." & theDBSheetColumnList.Value(i, 2) & IIf(Strings.Len(addRestrict) > 0, " AND " & addRestrict, "")
            '        End If
            '        ' we have additionally joined (an)other table(s) for the lookup display...
            '        If Strings.Len(completeJoin) > 0 Then
            '            ' remove any ORDER BY clause from completeJoin...
            '            restrPos = completeJoin.ToUpper().LastIndexOf(" ORDER") + 1
            '            If restrPos > 0 Then completeJoin = completeJoin.Substring(0, Math.Min(restrPos - 1, completeJoin.Length))
            '            ' ..and add join of additional subtable(s) to the query
            '            fromStr += " LEFT JOIN " & Environment.NewLine & completeJoin
            '        End If

            '        selectPart = fetch(theDBSheetColumnList.Value(i, 9), "SELECT ", " FROM ").Trim()
            '        ' remove second field in lookup query's select clause
            '        restrPos = selectPart.LastIndexOf(",") + 1
            '        selectPart = selectPart.Substring(0, Math.Min(restrPos - 1, selectPart.Length))
            '        ' complex select statement, take directly from lookup query..
            '        If selectPart <> theDBSheetColumnList.Value(i, 3) Then
            '            selectStr += quotedReplace(selectPart, "T" & tableCounter) & ", "
            '        Else
            '            ' simple select statement (only the lookup field and id), put together...
            '            selectStr += theTable & "." & theDBSheetColumnList.Value(i, 3) & " AS " & usedColumn & ", "
            '        End If
            '    End If
            'Next
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
        Try
            fillDatabases(Database)
        Catch ex As System.Exception
            ErrorMsg("Error: " & ex.Message)
        End Try
        Database.SelectedIndex = Database.Items.IndexOf(fetch(dbsheetConnString, dbidentifier, ";"))
        ' initialize with empty DBSheet definitions
        createDefinitions("")
    End Sub

    Dim oldPwdEntry As String = ""
    Private Sub Password_Leave(sender As Object, e As EventArgs) Handles Password.Leave
        If Password.Text <> oldPwdEntry Then
            Try
                fillDatabases(Database)
            Catch ex As System.Exception
                ErrorMsg("Error: " & ex.Message)
            End Try
            Database.SelectedIndex = Database.Items.IndexOf(fetch(dbsheetConnString, dbidentifier, ";"))
            ' initialize with empty DBSheet definitions
            createDefinitions("")
        End If
    End Sub

    Private Sub Password_Enter(sender As Object, e As EventArgs) Handles Password.Enter
        oldPwdEntry = Password.Text
    End Sub
End Class