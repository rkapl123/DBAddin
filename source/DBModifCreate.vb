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
        If Me.DBModifName.Text <> String.Empty Then
            ' Add doesn't work directly with ExcelDnaUtil.Application.ActiveWorkbook.Names (late binding), so create an object here...
            Dim NamesList As Excel.Names = ExcelDnaUtil.Application.ActiveWorkbook.Names
            Try
                NamesList.Add(Name:=Me.Tag & Me.DBModifName.Text, RefersTo:=ExcelDnaUtil.Application.ActiveCell)
            Catch ex As Exception
                NameValidationResult = ex.Message
            End Try
            Try : NamesList.Item(Me.Tag & Me.DBModifName.Text).Delete() : Catch ex As Exception : End Try
        End If
        ' check for requirements: mandatory fields filled (visible Tablename, Primary keys and Database), NameValidation above OK and no double invocation for execOnSave in DB Sequences and sequence parts
        If Me.Tablename.Text = String.Empty And Me.Tablename.Visible Then
            MsgBox("Field Tablename is required, please fill in!", MsgBoxStyle.Critical + vbOKOnly, "DBModification Validation")
        ElseIf Me.PrimaryKeys.Text = String.Empty And Me.PrimaryKeys.Visible Then
            MsgBox("Field Primary Keys is required, please fill in!", MsgBoxStyle.Critical + vbOKOnly, "DBModification Validation")
        ElseIf Me.Database.Text = String.Empty And Me.Database.Visible Then
            MsgBox("Field Database is required, please fill in!", MsgBoxStyle.Critical + vbOKOnly, "DBModification Validation")
        ElseIf NameValidationResult <> "" Then
            MsgBox("Invalid " & Me.NameLabel.Text & ", Error: " & NameValidationResult, MsgBoxStyle.Critical + vbOKOnly, "DBModification Validation")
        Else
            ' check for double invocation because of execOnSave both being set on current DB Modifier ...
            If Me.execOnSave.Checked And Globals.DBModifDefColl.ContainsKey("DBSeqnce") Then
                Dim MyDBModifName As String = Me.Tag & Me.DBModifName.Text
                ' and on DB Sequence that contains the current DB Mapper or DB Action:
                If Me.Tag <> "DBSeqnce" Then
                    For Each DBModifierCheck As DBSeqnce In Globals.DBModifDefColl("DBSeqnce").Values
                        ' check for Sequences that have execOnSave set...
                        If DBModifierCheck.DBModifSaveNeeded Then
                            ' ...if they contain the current DBAction/DBMapper
                            For Each sequenceParam As String In DBModifierCheck.getSequenceSteps
                                Dim definition() As String = Split(sequenceParam, ":")
                                If MyDBModifName = definition(1) Then
                                    Dim DBModifTargetAddress As String = "(Target Address could not be found...)"
                                    If Globals.DBModifDefColl(definition(0)).ContainsKey(definition(1)) Then DBModifTargetAddress = Globals.DBModifDefColl(definition(0)).Item(definition(1)).getTargetRangeAddress()
                                    Dim foundDBModifName As String = IIf(DBModifierCheck.getName = "DBSeqnce", "Unnamed DBSequence", DBModifierCheck.getName)
                                    MsgBox(Me.Tag & Me.DBModifName.Text & " in " & DBModifTargetAddress & " will be executed twice on saving, because it is part of '" & foundDBModifName & "', which is also executed on saving." & vbCrLf & IIf(Me.execOnSave.Checked, "Disabling 'Execute on save' now, you ", "You") & " can reenable after disabling it on '" & foundDBModifName & "'", MsgBoxStyle.Critical + vbOKOnly, "DBModification Validation")
                                    Me.execOnSave.Checked = False
                                End If
                            Next
                        End If
                    Next
                    If Not Me.execOnSave.Checked Then Exit Sub
                Else ' or on any DB Modifier being contained in current DB Sequence:
                    For i As Integer = 0 To Me.DBSeqenceDataGrid.Rows().Count - 2
                        Dim definition() As String = Split(Me.DBSeqenceDataGrid.Rows(i).Cells(0).Value, ":")
                        If (definition(0) = "DBAction" Or definition(0) = "DBMapper") AndAlso Globals.DBModifDefColl(definition(0)).ContainsKey(definition(1)) AndAlso Globals.DBModifDefColl(definition(0)).Item(definition(1)).DBModifSaveNeeded Then
                            Dim foundDBModifName As String = IIf(definition(1) = "", "Unnamed " & definition(0), definition(1))
                            Dim DBModifTargetAddress As String = Globals.DBModifDefColl(definition(0)).Item(definition(1)).getTargetRangeAddress()
                            MsgBox(foundDBModifName & " in " & DBModifTargetAddress & " will be executed twice on saving, because it is part of this DBSequence, which is also executed on saving." & vbCrLf & IIf(Me.execOnSave.Checked, "Disabling 'Execute on save' now, you ", "You") & " can reenable after disabling it on '" & foundDBModifName & "'", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "DBModification Validation")
                            Me.execOnSave.Checked = False
                        End If
                    Next
                    If Not Me.execOnSave.Checked Then Exit Sub
                End If
            End If
        End If
        Me.DialogResult = DialogResult.OK
        Me.Close()
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
            DBSeqStepValidationErrors += "Error in row " & e.RowIndex + 1 & ",content: " & Me.DBSeqenceDataGrid.Rows(e.RowIndex).Cells(0).Value & vbCrLf
        End If
    End Sub

    ''' <summary>the DBMapper and DBAction Target Range Address is displayed as a hyperlink, simulate this link here</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TargetRangeAddress_Click(sender As Object, e As EventArgs) Handles TargetRangeAddress.Click
        If Me.TargetRangeAddress.Text = "" Then Exit Sub
        Dim rangePart() As String = Split(Me.TargetRangeAddress.Text, "!")
        Try
            ExcelDnaUtil.Application.Worksheets(rangePart(0)).Select()
            ExcelDnaUtil.Application.Range(rangePart(1)).Select()
        Catch ex As Exception
            MsgBox("Couldn't select " & Me.TargetRangeAddress.Text & ":" & ex.Message)
        End Try
    End Sub

    ''' <summary>move row up in DataGridView of DB Sequence</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Up_Click(sender As Object, e As EventArgs) Handles up.Click
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
    Private Sub Down_Click(sender As Object, e As EventArgs) Handles down.Click
        If IsNothing(DBSeqenceDataGrid.CurrentRow) Then Return
        Dim rw As DataGridViewRow = DBSeqenceDataGrid.CurrentRow
        Dim selIndex As Integer = DBSeqenceDataGrid.CurrentRow.Index
        ' avoid moving down of last row, DBSeqenceDataGrid.Rows.Count is 1 more than the actual inserted rows because of the "new" row, selIndex is 0 based
        If selIndex = DBSeqenceDataGrid.Rows.Count - 2 Then Return
        DBSeqenceDataGrid.Rows.RemoveAt(selIndex)
        DBSeqenceDataGrid.Rows.Insert(selIndex + 1, rw)
        DBSeqenceDataGrid.Rows(selIndex + 1).Cells(0).Selected = True
    End Sub

    ''' <summary>Shown Event to display Data Errors when adding DBSequence Grid elements</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DBModifCreate_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        If DBSeqStepValidationErrors <> "" Then
            Dim cb As DataGridViewComboBoxColumn = DBSeqenceDataGrid.Columns(0)
            Dim ds As List(Of String) = cb.DataSource()
            Dim allowedValues As String = ""
            For Each def As String In ds
                allowedValues += def + vbCrLf
            Next
            Me.RepairDBSeqnce.Text = DBSeqStepValidationErrors & vbCrLf & "Allowed Entries are:" & vbCrLf & allowedValues & vbCrLf & "Repair existing definitions below and remove all above incl. this line to fix it by clicking OK:" & vbCrLf & Me.RepairDBSeqnce.Text
            Me.RepairDBSeqnce.Show()
            Me.RepairDBSeqnce.Width = Me.DBSeqenceDataGrid.Width
            Me.RepairDBSeqnce.Height = 325
            Me.RepairDBSeqnce.Top = Me.DatabaseLabel.Top
            Me.DBSeqenceDataGrid.Hide()
            Me.Tag = "repaired"
            MsgBox("Defined DBSequence steps did not match allowed values." & vbCrLf & "Please follow the instructions in textbox to fix it...", MsgBoxStyle.Critical, "DBSequence definition Insert error")
        End If
        DBSeqStepValidationErrorsShown = True
    End Sub

    ''' <summary>Create Commandbutton Click event</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CreateCB_Click(sender As Object, e As EventArgs) Handles CreateCB.Click
        ' create a commandbutton for the current DBmodification?
        Dim cbshp As Excel.OLEObject = ExcelDnaUtil.Application.ActiveSheet.OLEObjects.Add(ClassType:="Forms.CommandButton.1", Link:=False, DisplayAsIcon:=False, Left:=600, Top:=70, Width:=120, Height:=24)
        Dim cb As Forms.CommandButton = cbshp.Object
        Dim cbName As String = Me.Tag & Me.DBModifName.Text
        Try
            cb.Name = cbName
            cb.Caption = IIf(Me.DBModifName.Text = "", "Unnamed " & Me.Tag, Me.Tag & Me.DBModifName.Text)
        Catch ex As Exception
            cbshp.Delete()
            MsgBox("Couldn't name CommandButton '" & cbName & "': " & ex.Message, MsgBoxStyle.Critical, "CommandButton create Error")
            Exit Sub
        End Try
        If Len(cbName) > 31 Then
            cbshp.Delete()
            MsgBox("CommandButton codenames cannot be longer than 31 characters ! '" & cbName & "': ", MsgBoxStyle.Critical, "CommandButton create Error")
            Exit Sub
        End If
        ' fail to assign a handler? remove commandbutton (otherwise it gets hard to edit an existing DBModification with a different name).
        If Not AddInEvents.assignHandler(ExcelDnaUtil.Application.ActiveSheet) Then
            cbshp.Delete()
        End If
    End Sub
End Class
