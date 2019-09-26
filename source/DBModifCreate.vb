Imports System.Drawing ' for clientPoint in DBSeqenceDataGrid_DragDrop
Imports System.Windows.Forms

''' <summary>Dialog for creating DB Modifier configurations</summary>
Public Class DBModifCreate

    ''' <summary>check for required fields before closing</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Dim NameValidation As String = ""
        If Me.DBModifName.Text <> String.Empty Then
            Try
                hostApp.Names.Add(Name:=Me.DBModifName.Text, RefersTo:=hostApp.ActiveCell)
            Catch ex As Exception
                NameValidation = ex.Message
            End Try
            Try : hostApp.Names.Item(Me.DBModifName.Text).Delete() : Catch ex As Exception : End Try
        End If
        If Me.Tablename.Text = String.Empty And Me.Tablename.Visible Then
            MsgBox("Field Tablename is required, please fill in!")
        ElseIf Me.PrimaryKeys.Text = String.Empty And Me.PrimaryKeys.Visible Then
            MsgBox("Field Primary Keys is required, please fill in!")
        ElseIf Me.Database.Text = String.Empty And Me.Database.Visible Then
            MsgBox("Field Database is required, please fill in!")
        ElseIf NameValidation <> "" Then
            MsgBox("Invalid " & Me.NameLabel.Text & NameValidation)
        Else
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
        LogWarn(e.Exception.Message & ":" & e.RowIndex & ":" & e.Context.ToString())
    End Sub

    ''' <summary>the DBMapper and DBAction Target Range Address is displayed as a hyperlink, simulate this link here</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TargetRangeAddress_Click(sender As Object, e As EventArgs) Handles TargetRangeAddress.Click
        Dim rangePart() As String
        rangePart = Split(Me.TargetRangeAddress.Text, "!")
        Try
            hostApp.Worksheets(rangePart(0)).Select()
            hostApp.Range(rangePart(1)).Select()
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
    End Sub

End Class
