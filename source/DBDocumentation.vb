Imports System.Windows.Forms

Public Class DBDocumentation
    Private Sub CancelBtn_Click(sender As Object, e As EventArgs) Handles CancelBtn.Click
        Me.Close()
    End Sub

    Private Sub OKBtn_Click(sender As Object, e As EventArgs) Handles OKBtn.Click
        Me.Close()
    End Sub

    Private Sub DBDocTextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles DBDocTextBox.KeyDown
        If e.KeyCode = Keys.Enter Then Me.Close()
    End Sub
End Class