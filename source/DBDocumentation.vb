Imports System.Windows.Forms

''' <summary>Simple Popup Window for displaying Database documentation</summary>
Public Class DBDocumentation
    ''' <summary>for handling the "Enter" and "Escape" key press</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DBDocTextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles DBDocTextBox.KeyDown
        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Escape Then
            Try : theAdHocSQLDlg.propagatedFromDoc = True : Catch ex As Exception : End Try ' needed to not propagate escape key to AdHocSql form
            Me.Close()
        End If
    End Sub
    ''' <summary>for handling the "Enter" and "Escape" key press</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DBDocumentation_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Escape Then
            Try : theAdHocSQLDlg.propagatedFromDoc = True : Catch ex As Exception : End Try ' needed to not propagate escape key to AdHocSql form
            Me.Close()
        End If
    End Sub
End Class