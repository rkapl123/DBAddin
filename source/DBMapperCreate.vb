Imports System.Windows.Forms

''' <summary>Dialog for creating DB Mapper configurations</summary>
Public Class DBMapperCreate

    ''' <summary>check for required fields before closing</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Me.Tablename.Text = String.Empty Then
            MsgBox("Field Tablename is required, please fill !")
        ElseIf Me.PrimaryKeys.Text = String.Empty Then
            MsgBox("Field Primary Keys is required, please fill !")
        ElseIf Me.Database.Text = String.Empty Then
            MsgBox("Field Database is required, please fill !")
        Else
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        End If
    End Sub

    ''' <summary>ignore all changes</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

End Class
