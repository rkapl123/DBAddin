<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class AdHocSQL
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AdHocSQL))
        Me.AdHocSQLQueryResult = New System.Windows.Forms.DataGridView()
        Me.Execute = New System.Windows.Forms.Button()
        Me.LDatabase = New System.Windows.Forms.Label()
        Me.Database = New System.Windows.Forms.ComboBox()
        Me.Transfer = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.CloseBtn = New System.Windows.Forms.Button()
        Me.EnvSwitch = New System.Windows.Forms.ComboBox()
        Me.RowsReturned = New System.Windows.Forms.Label()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.TransferType = New System.Windows.Forms.ComboBox()
        Me.LEnv1 = New System.Windows.Forms.Label()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.SQLText = New System.Windows.Forms.RichTextBox()
        Me.ConfigMenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
        CType(Me.AdHocSQLQueryResult, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
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
        Me.AdHocSQLQueryResult.Location = New System.Drawing.Point(3, 3)
        Me.AdHocSQLQueryResult.Name = "AdHocSQLQueryResult"
        Me.AdHocSQLQueryResult.ReadOnly = True
        Me.AdHocSQLQueryResult.ShowEditingIcon = False
        Me.AdHocSQLQueryResult.Size = New System.Drawing.Size(770, 314)
        Me.AdHocSQLQueryResult.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.AdHocSQLQueryResult, "Results of executed SQL Commands")
        '
        'Execute
        '
        Me.Execute.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Execute.Location = New System.Drawing.Point(659, 448)
        Me.Execute.Name = "Execute"
        Me.Execute.Size = New System.Drawing.Size(61, 23)
        Me.Execute.TabIndex = 6
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
        Me.LDatabase.Location = New System.Drawing.Point(135, 453)
        Me.LDatabase.Name = "LDatabase"
        Me.LDatabase.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LDatabase.Size = New System.Drawing.Size(25, 13)
        Me.LDatabase.TabIndex = 102
        Me.LDatabase.Text = "DB:"
        '
        'Database
        '
        Me.Database.AllowDrop = True
        Me.Database.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Database.BackColor = System.Drawing.SystemColors.Window
        Me.Database.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Database.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Database.Location = New System.Drawing.Point(161, 450)
        Me.Database.Name = "Database"
        Me.Database.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Database.Size = New System.Drawing.Size(115, 21)
        Me.Database.Sorted = True
        Me.Database.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.Database, "select currently active database for SQL Command")
        '
        'Transfer
        '
        Me.Transfer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Transfer.Location = New System.Drawing.Point(512, 449)
        Me.Transfer.Name = "Transfer"
        Me.Transfer.Size = New System.Drawing.Size(61, 23)
        Me.Transfer.TabIndex = 4
        Me.Transfer.Text = "Transfer"
        Me.ToolTip1.SetToolTip(Me.Transfer, resources.GetString("Transfer.ToolTip"))
        Me.Transfer.UseVisualStyleBackColor = True
        '
        'CloseBtn
        '
        Me.CloseBtn.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CloseBtn.Location = New System.Drawing.Point(726, 448)
        Me.CloseBtn.Name = "CloseBtn"
        Me.CloseBtn.Size = New System.Drawing.Size(62, 23)
        Me.CloseBtn.TabIndex = 7
        Me.CloseBtn.Text = "Close"
        Me.ToolTip1.SetToolTip(Me.CloseBtn, "click (or press ESC) to finish Ad Hoc SQL Editor")
        Me.CloseBtn.UseVisualStyleBackColor = True
        '
        'EnvSwitch
        '
        Me.EnvSwitch.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.EnvSwitch.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.EnvSwitch.FormattingEnabled = True
        Me.EnvSwitch.Location = New System.Drawing.Point(42, 450)
        Me.EnvSwitch.Name = "EnvSwitch"
        Me.EnvSwitch.Size = New System.Drawing.Size(87, 21)
        Me.EnvSwitch.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.EnvSwitch, "select currently active environment (connection string) for SQL Command")
        '
        'RowsReturned
        '
        Me.RowsReturned.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.RowsReturned.AutoSize = True
        Me.RowsReturned.Location = New System.Drawing.Point(282, 453)
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
        'TransferType
        '
        Me.TransferType.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TransferType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.TransferType.FormattingEnabled = True
        Me.TransferType.Location = New System.Drawing.Point(579, 450)
        Me.TransferType.Name = "TransferType"
        Me.TransferType.Size = New System.Drawing.Size(74, 21)
        Me.TransferType.TabIndex = 5
        '
        'LEnv1
        '
        Me.LEnv1.AllowDrop = True
        Me.LEnv1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LEnv1.AutoSize = True
        Me.LEnv1.BackColor = System.Drawing.Color.Transparent
        Me.LEnv1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LEnv1.Location = New System.Drawing.Point(12, 453)
        Me.LEnv1.Name = "LEnv1"
        Me.LEnv1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LEnv1.Size = New System.Drawing.Size(29, 13)
        Me.LEnv1.TabIndex = 102
        Me.LEnv1.Text = "Env:"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(12, 12)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.SQLText)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.AdHocSQLQueryResult)
        Me.SplitContainer1.Size = New System.Drawing.Size(776, 430)
        Me.SplitContainer1.SplitterDistance = 106
        Me.SplitContainer1.TabIndex = 105
        '
        'SQLText
        '
        Me.SQLText.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SQLText.Location = New System.Drawing.Point(3, 3)
        Me.SQLText.Name = "SQLText"
        Me.SQLText.Size = New System.Drawing.Size(770, 100)
        Me.SQLText.TabIndex = 1
        Me.SQLText.Text = ""
        '
        'ConfigMenuStrip
        '
        Me.ConfigMenuStrip.Name = "ConfigMenuStrip"
        Me.ConfigMenuStrip.Size = New System.Drawing.Size(181, 26)
        '
        'AdHocSQL
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(799, 484)
        Me.Controls.Add(Me.EnvSwitch)
        Me.Controls.Add(Me.TransferType)
        Me.Controls.Add(Me.RowsReturned)
        Me.Controls.Add(Me.CloseBtn)
        Me.Controls.Add(Me.Transfer)
        Me.Controls.Add(Me.LEnv1)
        Me.Controls.Add(Me.LDatabase)
        Me.Controls.Add(Me.Database)
        Me.Controls.Add(Me.Execute)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MinimumSize = New System.Drawing.Size(815, 523)
        Me.Name = "AdHocSQL"
        Me.Text = "Ad Hoc SQL Editor"
        CType(Me.AdHocSQLQueryResult, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents AdHocSQLQueryResult As Windows.Forms.DataGridView
    Friend WithEvents Execute As Windows.Forms.Button
    Public WithEvents LDatabase As Windows.Forms.Label
    Public WithEvents Database As Windows.Forms.ComboBox
    Friend WithEvents Transfer As Windows.Forms.Button
    Friend WithEvents ToolTip1 As Windows.Forms.ToolTip
    Friend WithEvents CloseBtn As Windows.Forms.Button
    Friend WithEvents RowsReturned As Windows.Forms.Label
    Friend WithEvents Timer1 As Windows.Forms.Timer
    Friend WithEvents BackgroundWorker1 As ComponentModel.BackgroundWorker
    Friend WithEvents TransferType As Windows.Forms.ComboBox
    Friend WithEvents EnvSwitch As Windows.Forms.ComboBox
    Public WithEvents LEnv1 As Windows.Forms.Label
    Friend WithEvents SplitContainer1 As Windows.Forms.SplitContainer
    Friend WithEvents SQLText As Windows.Forms.RichTextBox
    Friend WithEvents ConfigMenuStrip As Windows.Forms.ContextMenuStrip
End Class
