<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DBMapperErrors
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DBMapperErrors))
        Me.ErrorDataGrid = New System.Windows.Forms.DataGridView()
        CType(Me.ErrorDataGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ErrorDataGrid
        '
        Me.ErrorDataGrid.AllowUserToAddRows = False
        Me.ErrorDataGrid.AllowUserToDeleteRows = False
        Me.ErrorDataGrid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ErrorDataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.ErrorDataGrid.Location = New System.Drawing.Point(12, 12)
        Me.ErrorDataGrid.Name = "ErrorDataGrid"
        Me.ErrorDataGrid.ReadOnly = True
        Me.ErrorDataGrid.Size = New System.Drawing.Size(699, 253)
        Me.ErrorDataGrid.TabIndex = 0
        '
        'DBMapperErrors
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(723, 278)
        Me.Controls.Add(Me.ErrorDataGrid)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "DBMapperErrors"
        Me.Text = "Storing DBMapper modifications in database had following errors, fix them and ret" &
    "ry..."
        CType(Me.ErrorDataGrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ErrorDataGrid As Windows.Forms.DataGridView
End Class
