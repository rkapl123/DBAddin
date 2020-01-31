Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports Microsoft.Vbe.Interop
Imports System.Windows.Forms
Imports System.Collections.Generic


''' <summary>Dialog for creating DB Modifier configurations</summary>
Public Class DBModifCreate
    Private DBSeqStepValidationErrors As String = ""
    Private DBSeqStepValidationErrorsShown As Boolean = False


    ''' <summary>check for required fields before closing</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Dim NameValidationResult As String = ""
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
        If Me.Tablename.Text = String.Empty And Me.Tablename.Visible Then
            MsgBox("Field Tablename is required, please fill in!")
        ElseIf Me.PrimaryKeys.Text = String.Empty And Me.PrimaryKeys.Visible Then
            MsgBox("Field Primary Keys is required, please fill in!")
        ElseIf Me.Database.Text = String.Empty And Me.Database.Visible Then
            MsgBox("Field Database is required, please fill in!")
        ElseIf NameValidationResult <> "" Then
            MsgBox("Invalid " & Me.NameLabel.Text & ", Error: " & NameValidationResult)
        Else
            ' check for double invocation because of execOnSave being set on DBAction/DBMapper
            If Me.Tag <> "DBSeqnce" Then
                For Each docproperty In ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties
                    If TypeName(docproperty.Value) = "String" And Strings.Left(docproperty.Name, 8) = "DBSeqnce" Then
                        Dim dbseqName As String = Replace(docproperty.Name, "DBSeqnce", "")
                        Dim params() As String = Split(docproperty.Value, ",")
                        Dim execSeqElemOnSave As Boolean = Convert.ToBoolean(params(0))
                        Dim i As Integer
                        For i = 1 To UBound(params)
                            Dim definition() As String = Split(params(i), ":")
                            If definition(0) = Me.Tag AndAlso definition(1) = Me.DBModifName.Text AndAlso
                            DBModifDefColl.ContainsKey(definition(0)) AndAlso DBModifDefColl(definition(0)).ContainsKey(definition(1)) AndAlso
                            Me.execOnSave.Checked AndAlso execSeqElemOnSave Then
                                Dim retval As MsgBoxResult = MsgBox(Me.Tag & Me.DBModifName.Text & " in " & DBModifDefColl(definition(0)).Item(definition(1)).getTargetRangeAddress() & " will be executed twice on saving, because it is part of DBSequence " & dbseqName & ", which is also executed on saving." & vbCrLf & "Please disable 'Execute on save' either here or on " & dbseqName & " !", MsgBoxStyle.Critical + vbOKCancel, "DBModification Validation")
                                Exit Sub
                            End If
                        Next
                    End If
                Next
                ' check for double invocation because of execOnSave being set on DBSequence
            Else
                For Each docproperty In ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties
                    For i As Integer = 0 To Me.DBSeqenceDataGrid.Rows().Count - 2
                        Dim definition() As String = Split(Me.DBSeqenceDataGrid.Rows(i).Cells(0).Value, ":")
                        If TypeName(docproperty.Value) = "String" And docproperty.Name = definition(0) & definition(1) Then
                            Dim DBModifParams() As String = functionSplit(docproperty.Value, ",", """", "def", "(", ")")
                            Dim storeDBMapOnSave As Boolean = False
                            If definition(0) = "DBAction" Then
                                If DBModifParams.Length > 2 AndAlso DBModifParams(2) <> "" Then storeDBMapOnSave = Convert.ToBoolean(DBModifParams(2))
                            ElseIf definition(0) = "DBMapper" Then
                                If DBModifParams(7) <> "" Then storeDBMapOnSave = Convert.ToBoolean(DBModifParams(7))
                            End If
                            If Me.execOnSave.Checked And storeDBMapOnSave Then
                                Dim foundDBModifName As String = definition(0) & IIf(definition(1) = "", "Unnamed " & definition(0), definition(1))
                                Dim retval As MsgBoxResult = MsgBox(foundDBModifName & " will be executed twice on saving, because it is part of this DBSequence, which is also executed on saving." & vbCrLf & "Please disable Execute on save either here or on '" & foundDBModifName & "'", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "DBModification Validation")
                                Exit Sub
                            End If
                        End If
                    Next
                Next
            End If
            ' need to create a commandbutton for the current DBmodification?
            If Me.CBCreate.Checked Then
                Dim cbshp As Excel.OLEObject = ExcelDnaUtil.Application.ActiveSheet.OLEObjects.Add(ClassType:="Forms.CommandButton.1", Link:=False, DisplayAsIcon:=False, Left:=600, Top:=70, Width:=120, Height:=24)
                Dim cb As Forms.CommandButton = cbshp.Object
                Dim cbName As String = Me.Tag & Me.DBModifName.Text
                Try
                    cb.Name = cbName
                    cb.Caption = IIf(Me.DBModifName.Text = "", "Unnamed " & Me.Tag, Me.Tag & Me.DBModifName.Text)
                Catch ex As Exception
                    MsgBox("Couldn't name CommandButton '" & cbName & "': " & ex.Message, MsgBoxStyle.Critical, "CommandButton create Error")
                    cbshp.Delete()
                    Exit Sub
                End Try
                If Len(cbName) > 31 Then
                    MsgBox("CommandButton codenames cannot be longer than 31 characters ! '" & cbName & "': ", MsgBoxStyle.Critical, "CommandButton create Error")
                    cbshp.Delete()
                    Exit Sub
                End If
                ' fail to assign a handler? remove commandbutton (otherwise it gets hard to edit an existing DBModification with a different name).
                If Not AddInEvents.assignHandler(ExcelDnaUtil.Application.ActiveSheet) Then
                    cbshp.Delete()
                    Exit Sub
                End If
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
            DBSeqStepValidationErrors += "Error in row " & e.RowIndex + 1 & ",content: " & Me.DBSeqenceDataGrid.Rows(e.RowIndex).Cells(0).Value & vbCrLf
        End If
    End Sub

    ''' <summary>the DBMapper and DBAction Target Range Address is displayed as a hyperlink, simulate this link here</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TargetRangeAddress_Click(sender As Object, e As EventArgs) Handles TargetRangeAddress.Click
        Dim rangePart() As String
        rangePart = Split(Me.TargetRangeAddress.Text, "!")
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

    Private Sub DBModifCreate_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        If DBSeqStepValidationErrors <> "" Then
            Dim cb As DataGridViewComboBoxColumn = DBSeqenceDataGrid.Columns(0)
            Dim ds As List(Of String) = cb.DataSource()
            Dim allowedValues As String = ""
            For Each def As String In ds
                allowedValues += def + vbCrLf
            Next
            Me.RepairDBSeqnce.Text = DBSeqStepValidationErrors & vbCrLf & "allowed Entries are:" & vbCrLf & allowedValues & vbCrLf & "modify existing definition below and remove everything else to repair (clicking OK):" & vbCrLf & Me.RepairDBSeqnce.Text
            Me.RepairDBSeqnce.Show()
            Me.RepairDBSeqnce.Width = Me.DBSeqenceDataGrid.Width
            Me.RepairDBSeqnce.Height = 325
            Me.RepairDBSeqnce.Top = Me.DatabaseLabel.Top
            Me.DBSeqenceDataGrid.Hide()
            Me.Tag = "repaired"
            MsgBox("Defined DBSequence steps did not fit definitions." & vbCrLf & "Please follow the instructions in textbox to repair it...", MsgBoxStyle.Critical, "DBSequence definition Insert error")
        End If
        DBSeqStepValidationErrorsShown = True
    End Sub

End Class
