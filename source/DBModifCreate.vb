Imports System.Windows.Forms

''' <summary>Dialog for creating DB Modifier configurations</summary>
Public Class DBModifCreate

    ''' <summary>check for required fields before closing</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Me.Tablename.Text = String.Empty And Me.Tablename.Visible Then
            MsgBox("Field Tablename is required, please fill !")
        ElseIf Me.PrimaryKeys.Text = String.Empty And Me.PrimaryKeys.Visible Then
            MsgBox("Field Primary Keys is required, please fill !")
        ElseIf Me.Database.Text = String.Empty And Me.Database.Visible Then
            MsgBox("Field Database is required, please fill !")
        Else
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End If
    End Sub

    ''' <summary>ignore all changes</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub DBSeqenceDataGrid_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DBSeqenceDataGrid.DataError
        LogWarn(e.Exception.Message & ":" & e.RowIndex & ":" & e.Context.ToString())
    End Sub

End Class
