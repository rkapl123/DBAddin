<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DBDocumentation
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.DBDocTextBox = New System.Windows.Forms.RichTextBox()
        Me.CancelBtn = New System.Windows.Forms.Button()
        Me.OKBtn = New System.Windows.Forms.Button()
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
        'CancelBtn
        '
        Me.CancelBtn.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CancelBtn.Location = New System.Drawing.Point(12, 270)
        Me.CancelBtn.Name = "CancelBtn"
        Me.CancelBtn.Size = New System.Drawing.Size(0, 0)
        Me.CancelBtn.TabIndex = 0
        Me.CancelBtn.Text = "OKCancel"
        Me.CancelBtn.UseVisualStyleBackColor = True
        '
        'OKBtn
        '
        Me.OKBtn.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.OKBtn.Location = New System.Drawing.Point(18, 269)
        Me.OKBtn.Name = "OKBtn"
        Me.OKBtn.Size = New System.Drawing.Size(0, 0)
        Me.OKBtn.TabIndex = 2
        Me.OKBtn.Text = "OK"
        Me.OKBtn.UseVisualStyleBackColor = True
        '
        'DBDocumentation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.CancelBtn
        Me.ClientSize = New System.Drawing.Size(402, 275)
        Me.Controls.Add(Me.OKBtn)
        Me.Controls.Add(Me.CancelBtn)
        Me.Controls.Add(Me.DBDocTextBox)
        Me.Name = "DBDocumentation"
        Me.Text = "DBAddin: Documentation for Config Object"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents DBDocTextBox As Windows.Forms.RichTextBox
    Friend WithEvents CancelBtn As Windows.Forms.Button
    Friend WithEvents OKBtn As Windows.Forms.Button
End Class
