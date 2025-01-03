<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class DBDocumentation
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.DBDocTextBox = New System.Windows.Forms.RichTextBox()
        Me.SuspendLayout()
        '
        'DBDocTextBox
        '
        Me.DBDocTextBox.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DBDocTextBox.Location = New System.Drawing.Point(12, 12)
        Me.DBDocTextBox.Name = "DBDocTextBox"
        Me.DBDocTextBox.ReadOnly = True
        Me.DBDocTextBox.Size = New System.Drawing.Size(378, 251)
        Me.DBDocTextBox.TabIndex = 1
        Me.DBDocTextBox.TabStop = False
        Me.DBDocTextBox.Text = ""
        '
        'DBDocumentation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(402, 275)
        Me.Controls.Add(Me.DBDocTextBox)
        Me.Name = "DBDocumentation"
        Me.Text = "DBAddin: Documentation for Config Object"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents DBDocTextBox As Windows.Forms.RichTextBox
End Class
