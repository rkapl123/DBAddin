Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports Microsoft.Vbe.Interop
Imports System.Windows.Forms
Imports System.Collections.Generic


''' <summary>Dialog for creating DB Modifier configurations</summary>
Public Class DBModifCreate
    ''' <summary>on loading of Form catch Data Errors produced when filling DBSeqenceDataGrid here</summary>
    Private DBSeqStepValidationErrors As String = ""
    ''' <summary>only catch errors until Form is displayed</summary>
    Private DBSeqStepValidationErrorsShown As Boolean = False

    ''' <summary>check for required fields before closing</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Dim NameValidationResult As String = ""
        ' Check for valid range name
        If Me.DBModifName.Text <> "" Then
            Dim checkName As String = Me.Tag + Me.DBModifName.Text
            If checkName.Length() > 255 Then
                NameValidationResult = "more than 255 characters long (including " + Me.Tag + ") !"
            ElseIf IsNumeric(Strings.Left(checkName, 1)) Then
                NameValidationResult = "starts with a number !"
            Else
                For i As Integer = 0 To checkName.Length - 1
                    If Not Char.IsLetterOrDigit(checkName.Chars(i)) And checkName.Chars(i) <> "_" Then
                        NameValidationResult = "contains non-alphanumeric character: " + checkName.Chars(i).ToString() + " !"
                        Exit For
                    End If
                Next
            End If
        End If
        Dim primKeys As Integer = 0
        ' besides valid range name, also check for requirements: 
        ' mandatory fields filled (visible Tablename, Primary keys And Database), NameValidation above OK, no Double invocation for execOnSave in DB Sequences And sequence parts and only one primary key for AutoInc Flag
        ' Beware: All If/ElseIf branches have to contain an validation error message, because the dialog stays open in this case. Only the Else branch closes the dialog.
        If NameValidationResult <> "" Then
            Globals.UserMsg("Invalid DBModifier name '" + Me.DBModifName.Text + "', Error: " + NameValidationResult, "DBModification Validation Error")
        ElseIf Me.Tablename.Text = "" And Me.Tablename.Visible Then
            Globals.UserMsg("Field Tablename is required, please fill in!", "DBModification Validation Error")
        ElseIf Me.PrimaryKeys.Visible AndAlso Not Integer.TryParse(Me.PrimaryKeys.Text, primKeys) Then
            Globals.UserMsg("Field Primary Keys is required and has to be an integer number, please fill in accordingly!", "DBModification Validation  Error")
        ElseIf Me.Database.Text = "" And Me.Database.Visible Then
            Globals.UserMsg("Field Database is required, please fill in!", "DBModification Validation Error")
        ElseIf Me.Tag = "DBMapper" AndAlso Me.AutoIncFlag.Checked AndAlso primKeys > 1 Then
            Globals.UserMsg("Only one primary key is allowed when Auto Incrementing is enabled!", "DBModification Validation Error")
        Else
            ' check for double invocation because of execOnSave both being set on current DB Modifier ...
            If Me.execOnSave.Checked And Globals.DBModifDefColl.ContainsKey("DBSeqnce") Then
                Dim MyDBModifName As String = Me.Tag + Me.DBModifName.Text
                ' and on DB Sequence that contains the current DB Mapper or DB Action:
                If Me.Tag <> "DBSeqnce" Then
                    For Each DBModifierCheck As DBSeqnce In Globals.DBModifDefColl("DBSeqnce").Values
                        ' check for Sequences that have execOnSave set...
                        If DBModifierCheck.execOnSave Then
                            ' ...if they contain the current DBAction/DBMapper
                            For Each sequenceParam As String In DBModifierCheck.getSequenceSteps
                                Dim definition() As String = Split(sequenceParam, ":")
                                If MyDBModifName = definition(1) Then
                                    Dim DBModifTargetAddress As String = "(Target Address could not be found...)"
                                    If Globals.DBModifDefColl(definition(0)).ContainsKey(definition(1)) Then DBModifTargetAddress = Globals.DBModifDefColl(definition(0)).Item(definition(1)).getTargetRangeAddress()
                                    Dim foundDBModifName As String = IIf(DBModifierCheck.getName = "DBSeqnce", "Unnamed DBSequence", DBModifierCheck.getName)
                                    Globals.UserMsg(Me.Tag + Me.DBModifName.Text + " in " + DBModifTargetAddress + " will be executed twice on saving, because it is part of '" + foundDBModifName + "', which is also executed on saving." + vbCrLf + IIf(Me.execOnSave.Checked, "Disabling 'Execute on save' now, you ", "You") + " can reenable after disabling it on '" + foundDBModifName + "'", "DBModification Validation")
                                    Me.execOnSave.Checked = False
                                End If
                            Next
                        End If
                    Next
                    If Not Me.execOnSave.Checked Then Exit Sub
                Else ' or on any DB Modifier being contained in current DB Sequence:
                    For i As Integer = 0 To Me.DBSeqenceDataGrid.Rows().Count - 2
                        Dim definition() As String = Split(Me.DBSeqenceDataGrid.Rows(i).Cells(0).Value, ":")
                        If (definition(0) = "DBAction" Or definition(0) = "DBMapper") AndAlso Globals.DBModifDefColl(definition(0)).ContainsKey(definition(1)) AndAlso Globals.DBModifDefColl(definition(0)).Item(definition(1)).execOnSave Then
                            Dim foundDBModifName As String = IIf(definition(1) = "", "Unnamed " + definition(0), definition(1))
                            Dim DBModifTargetAddress As String = Globals.DBModifDefColl(definition(0)).Item(definition(1)).getTargetRangeAddress()
                            Globals.UserMsg(foundDBModifName + " in " + DBModifTargetAddress + " will be executed twice on saving, because it is part of this DBSequence, which is also executed on saving." + vbCrLf + IIf(Me.execOnSave.Checked, "Disabling 'Execute on save' now, you ", "You") + " can reenable after disabling it on '" + foundDBModifName + "'", "DBModification Validation")
                            Me.execOnSave.Checked = False
                        End If
                    Next
                    If Not Me.execOnSave.Checked Then Exit Sub
                End If
            End If
            If Me.Tag = "DBSeqnce" Then
                Dim TransactionOpened As Boolean = False
                For i As Integer = 0 To Me.DBSeqenceDataGrid.Rows().Count - 2
                    Dim definition() As String = Split(Me.DBSeqenceDataGrid.Rows(i).Cells(0).Value, ":")
                    If (definition(0) = "DBBegin") Then
                        TransactionOpened = True
                    ElseIf (definition(0) = "DBCommitRollback") Then
                        TransactionOpened = False
                    End If
                    If (TransactionOpened And Strings.Left(definition(0), 7) = "Refresh") Then
                        Globals.UserMsg("You placed a " + definition(0) + " inside of a transaction, currently this might lead to deadlocks as DB functions use a different connection method (ADODB) than DB Modifiers (ADO.NET)." + vbCrLf + "If the DB function done in the refresh doesn't query any data being modified inside the transaction, you may ignore this warning.", "DBModification Validation", MsgBoxStyle.Exclamation)
                    End If
                Next
            End If
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End If
    End Sub

    ''' <summary>ignore all done changes in dialog</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    ''' <summary>in case of (actually impossible) data errors in DBSequence DataGridView row entries, catch and log them here</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DBSeqenceDataGrid_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DBSeqenceDataGrid.DataError
        If Not DBSeqStepValidationErrorsShown Then
            DBSeqStepValidationErrors += "Error in row " + (e.RowIndex + 1).ToString() + ",content: " + Me.DBSeqenceDataGrid.Rows(e.RowIndex).Cells(0).Value + vbCrLf
        End If
    End Sub

    ''' <summary>the DBMapper and DBAction Target Range Address is displayed as a hyperlink, simulate this link here</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TargetRangeAddress_Click(sender As Object, e As EventArgs) Handles TargetRangeAddress.Click
        If Me.TargetRangeAddress.Text = "" Then Exit Sub
        ' only get TargetRangeAddress up to bracket (possibly contained named range formula)
        Dim clickAddress = Me.TargetRangeAddress.Text.Substring(0, IIf(Me.TargetRangeAddress.Text.IndexOf("(") < 0, Me.TargetRangeAddress.Text.Length, Me.TargetRangeAddress.Text.IndexOf("(")))
        Dim rangePart() As String = Split(clickAddress, "!")
        Try
            ExcelDnaUtil.Application.Worksheets(rangePart(0)).Select()
            ExcelDnaUtil.Application.Range(rangePart(1)).Select()
        Catch ex As Exception
            Globals.UserMsg("Couldn't select " + clickAddress + ":" + ex.Message)
        End Try
    End Sub

    ''' <summary>move row up in DataGridView of DB Sequence</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub MoveRowUp_Click(sender As Object, e As EventArgs) Handles MoveRowUp.Click
        If IsNothing(DBSeqenceDataGrid.CurrentRow) Then Return
        Dim rw As DataGridViewRow = DBSeqenceDataGrid.CurrentRow
        Dim selIndex As Integer = DBSeqenceDataGrid.CurrentRow.Index
        ' avoid moving up of first row
        If selIndex = 0 Then Return
        DBSeqenceDataGrid.Rows.RemoveAt(selIndex)
        DBSeqenceDataGrid.Rows.Insert(selIndex - 1, rw)
        DBSeqenceDataGrid.Rows(selIndex - 1).Cells(0).Selected = True
    End Sub

    ''' <summary>move row down in DataGridView of DB Sequence</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub MoveRowDown_Click(sender As Object, e As EventArgs) Handles MoveRowDown.Click
        If IsNothing(DBSeqenceDataGrid.CurrentRow) Then Return
        Dim rw As DataGridViewRow = DBSeqenceDataGrid.CurrentRow
        Dim selIndex As Integer = DBSeqenceDataGrid.CurrentRow.Index
        ' avoid moving down of last row, DBSeqenceDataGrid.Rows.Count is 1 more than the actual inserted rows because of the "new" row, selIndex is 0 based
        If selIndex = DBSeqenceDataGrid.Rows.Count - 2 Then Return
        DBSeqenceDataGrid.Rows.RemoveAt(selIndex)
        DBSeqenceDataGrid.Rows.Insert(selIndex + 1, rw)
        DBSeqenceDataGrid.Rows(selIndex + 1).Cells(0).Selected = True
    End Sub

    Private selRowIndex As Integer
    Private selColIndex As Integer

    ''' <summary>prepare context menus to be displayed after right mouse click</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DBSeqenceDataGrid_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DBSeqenceDataGrid.CellMouseDown
        selRowIndex = e.RowIndex
        selColIndex = e.ColumnIndex
        If e.Button = System.Windows.Forms.MouseButtons.Right Then
            DBSeqenceDataGrid.ContextMenuStrip = MoveMenu
        End If
    End Sub

    ''' <summary>Shown Event to display Data Errors when adding DBSequence Grid elements</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DBModifCreate_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        ' on creating the form and DBSequenceDataGrid in DBModif.createDBModif, Data Errors produced when filling DBSeqenceDataGrid are caught by 
        ' DBModifCreate.DBSeqenceDataGrid_DataError event procedure and stored in DBSeqStepValidationErrors. 
        ' If any errors have been caught, display these in alternate textform RepairDBSeqnce along with instructions on how to repair them
        If DBSeqStepValidationErrors <> "" Then
            ' first get allowed values from filled DataGridView DataSource
            Dim cb As DataGridViewComboBoxColumn = DBSeqenceDataGrid.Columns(0)
            Dim ds As List(Of String) = cb.DataSource()
            Dim allowedValues As String = ""
            For Each def As String In ds
                allowedValues += def + vbCrLf
            Next
            ' then display allowed values along with error messages and instruction on how to repair.
            Me.RepairDBSeqnce.Text = DBSeqStepValidationErrors + vbCrLf + "Allowed Entries are:" + vbCrLf + allowedValues + vbCrLf + "Repair existing definitions below and remove all above incl. this line to fix it by clicking OK:" + vbCrLf + Me.RepairDBSeqnce.Text
            Me.RepairDBSeqnce.Show()
            Me.RepairDBSeqnce.Width = Me.DBSeqenceDataGrid.Width
            Me.RepairDBSeqnce.Height = 325
            Me.RepairDBSeqnce.Top = Me.DatabaseLabel.Top
            Me.DBSeqenceDataGrid.Hide()
            ' go into "repaired" mode (indicating rewriting DBSequence Steps in DBModif.createDBModif)
            Me.Tag = "repaired"
            Globals.UserMsg("Defined DBSequence steps did not match allowed values." + vbCrLf + "Please follow the instructions in textbox to fix it...", "DBSequence definition Insert error")
        End If
        DBSeqStepValidationErrorsShown = True
    End Sub

    ''' <summary>Create Commandbutton Click event</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CreateCB_Click(sender As Object, e As EventArgs) Handles CreateCB.Click
        ' create a commandbutton for the current DBmodification?
        Dim cbshp As Excel.OLEObject = Nothing
        Dim cb As Forms.CommandButton = Nothing
        Try
            cbshp = ExcelDnaUtil.Application.ActiveSheet.OLEObjects.Add(ClassType:="Forms.CommandButton.1", Link:=False, DisplayAsIcon:=False, Left:=600, Top:=70, Width:=120, Height:=24)
            cb = cbshp.Object
        Catch ex As Exception
            Globals.UserMsg("Can't create command button: " + ex.Message, "CommandButton create Error")
            cbshp.Delete()
            Exit Sub
        End Try
        Dim cbName As String = Me.Tag + Me.DBModifName.Text
        Try
            cb.Name = cbName
            cb.Caption = IIf(Me.DBModifName.Text = "", "Unnamed " + Me.Tag, Me.Tag + Me.DBModifName.Text)
        Catch ex As Exception
            cbshp.Delete()
            If ex.Message.Contains("HRESULT: 0x8002802C (TYPE_E_AMBIGUOUSNAME)") Then
                Globals.UserMsg("Can't name the new command button '" + cbName + "' as there already exists a button with that name", "CommandButton create Error")
            Else
                Globals.UserMsg("Can't name command button '" + cbName + "': " + ex.Message, "CommandButton create Error")
            End If
            Exit Sub
        End Try
        If Len(cbName) > 31 Then
            cbshp.Delete()
            Globals.UserMsg("CommandButton codenames cannot be longer than 31 characters ! '" + cbName + "': ", "CommandButton create Error")
            Exit Sub
        End If
        ' fail to assign a handler? remove commandbutton (otherwise it gets hard to edit an existing DBModification with a different name).
        If Not AddInEvents.assignHandler(ExcelDnaUtil.Application.ActiveSheet) Then
            cbshp.Delete()
        End If
    End Sub

End Class