<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AdHocSQL
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AdHocSQL))
        Me.SQLText = New System.Windows.Forms.TextBox()
        Me.AdHocSQLQueryResult = New System.Windows.Forms.DataGridView()
        Me.Execute = New System.Windows.Forms.Button()
        Me.LDatabase = New System.Windows.Forms.Label()
        Me.Database = New System.Windows.Forms.ComboBox()
        Me.Transfer = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.CloseBtn = New System.Windows.Forms.Button()
        Me.CloseAndAdd = New System.Windows.Forms.Button()
        Me.RowsReturned = New System.Windows.Forms.Label()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        CType(Me.AdHocSQLQueryResult, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SQLText
        '
        Me.SQLText.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SQLText.Location = New System.Drawing.Point(12, 12)
        Me.SQLText.Multiline = True
        Me.SQLText.Name = "SQLText"
        Me.SQLText.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.SQLText.Size = New System.Drawing.Size(776, 115)
        Me.SQLText.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.SQLText, "Enter SQL Code here and click Execute (or press Ctrl-Return) to execute it")
        '
        'AdHocSQLQueryResult
        '
        Me.AdHocSQLQueryResult.AllowUserToAddRows = False
        Me.AdHocSQLQueryResult.AllowUserToDeleteRows = False
        Me.AdHocSQLQueryResult.AllowUserToOrderColumns = True
        Me.AdHocSQLQueryResult.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AdHocSQLQueryResult.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.AdHocSQLQueryResult.Location = New System.Drawing.Point(13, 133)
        Me.AdHocSQLQueryResult.Name = "AdHocSQLQueryResult"
        Me.AdHocSQLQueryResult.ReadOnly = True
        Me.AdHocSQLQueryResult.ShowEditingIcon = False
        Me.AdHocSQLQueryResult.Size = New System.Drawing.Size(775, 305)
        Me.AdHocSQLQueryResult.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.AdHocSQLQueryResult, "Results of executed SQL Commands")
        '
        'Execute
        '
        Me.Execute.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Execute.Location = New System.Drawing.Point(589, 449)
        Me.Execute.Name = "Execute"
        Me.Execute.Size = New System.Drawing.Size(61, 23)
        Me.Execute.TabIndex = 4
        Me.Execute.Text = "Execute"
        Me.ToolTip1.SetToolTip(Me.Execute, "click (or press Ctrl+Return) to execute entered SQL Commands ")
        Me.Execute.UseVisualStyleBackColor = True
        '
        'LDatabase
        '
        Me.LDatabase.AllowDrop = True
        Me.LDatabase.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LDatabase.AutoSize = True
        Me.LDatabase.BackColor = System.Drawing.Color.Transparent
        Me.LDatabase.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LDatabase.Location = New System.Drawing.Point(16, 453)
        Me.LDatabase.Name = "LDatabase"
        Me.LDatabase.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LDatabase.Size = New System.Drawing.Size(56, 13)
        Me.LDatabase.TabIndex = 102
        Me.LDatabase.Text = "Database:"
        '
        'Database
        '
        Me.Database.AllowDrop = True
        Me.Database.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Database.BackColor = System.Drawing.SystemColors.Window
        Me.Database.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Database.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Database.Location = New System.Drawing.Point(75, 449)
        Me.Database.Name = "Database"
        Me.Database.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Database.Size = New System.Drawing.Size(115, 21)
        Me.Database.Sorted = True
        Me.Database.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.Database, "select currently active database for SQL Command")
        '
        'Transfer
        '
        Me.Transfer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Transfer.Location = New System.Drawing.Point(522, 449)
        Me.Transfer.Name = "Transfer"
        Me.Transfer.Size = New System.Drawing.Size(61, 23)
        Me.Transfer.TabIndex = 3
        Me.Transfer.Text = "Transfer"
        Me.ToolTip1.SetToolTip(Me.Transfer, "click (or press Shift+Return) to transfer SQL Command to selected cell")
        Me.Transfer.UseVisualStyleBackColor = True
        '
        'CloseBtn
        '
        Me.CloseBtn.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CloseBtn.Location = New System.Drawing.Point(656, 449)
        Me.CloseBtn.Name = "CloseBtn"
        Me.CloseBtn.Size = New System.Drawing.Size(62, 23)
        Me.CloseBtn.TabIndex = 5
        Me.CloseBtn.Text = "Close"
        Me.ToolTip1.SetToolTip(Me.CloseBtn, "click to leave Ad Hoc SQL Editor without action")
        Me.CloseBtn.UseVisualStyleBackColor = True
        '
        'CloseAndAdd
        '
        Me.CloseAndAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CloseAndAdd.Location = New System.Drawing.Point(724, 449)
        Me.CloseAndAdd.Name = "CloseAndAdd"
        Me.CloseAndAdd.Size = New System.Drawing.Size(64, 23)
        Me.CloseAndAdd.TabIndex = 6
        Me.CloseAndAdd.Text = "Add"
        Me.ToolTip1.SetToolTip(Me.CloseAndAdd, "Close and add the SQL Comman to the AdHocSQL Combo")
        Me.CloseAndAdd.UseVisualStyleBackColor = True
        '
        'RowsReturned
        '
        Me.RowsReturned.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.RowsReturned.AutoSize = True
        Me.RowsReturned.Location = New System.Drawing.Point(196, 453)
        Me.RowsReturned.Name = "RowsReturned"
        Me.RowsReturned.Size = New System.Drawing.Size(0, 13)
        Me.RowsReturned.TabIndex = 104
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerSupportsCancellation = True
        '
        'AdHocSQL
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(799, 484)
        Me.Controls.Add(Me.RowsReturned)
        Me.Controls.Add(Me.CloseAndAdd)
        Me.Controls.Add(Me.CloseBtn)
        Me.Controls.Add(Me.Transfer)
        Me.Controls.Add(Me.LDatabase)
        Me.Controls.Add(Me.Database)
        Me.Controls.Add(Me.Execute)
        Me.Controls.Add(Me.AdHocSQLQueryResult)
        Me.Controls.Add(Me.SQLText)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "AdHocSQL"
        Me.Text = "Ad Hoc SQL Editor"
        CType(Me.AdHocSQLQueryResult, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents SQLText As Windows.Forms.TextBox
    Friend WithEvents AdHocSQLQueryResult As Windows.Forms.DataGridView
    Friend WithEvents Execute As Windows.Forms.Button
    Public WithEvents LDatabase As Windows.Forms.Label
    Public WithEvents Database As Windows.Forms.ComboBox
    Friend WithEvents Transfer As Windows.Forms.Button
    Friend WithEvents ToolTip1 As Windows.Forms.ToolTip
    Friend WithEvents CloseBtn As Windows.Forms.Button
    Friend WithEvents CloseAndAdd As Windows.Forms.Button
    Friend WithEvents RowsReturned As Windows.Forms.Label
    Friend WithEvents Timer1 As Windows.Forms.Timer
    Friend WithEvents BackgroundWorker1 As ComponentModel.BackgroundWorker
End Class
