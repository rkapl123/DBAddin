<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class DBModifCreate
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DBModifCreate))
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.CreateCB = New System.Windows.Forms.Button()
        Me.DBModifName = New System.Windows.Forms.TextBox()
        Me.NameLabel = New System.Windows.Forms.Label()
        Me.Tablename = New System.Windows.Forms.TextBox()
        Me.PrimaryKeys = New System.Windows.Forms.TextBox()
        Me.Database = New System.Windows.Forms.TextBox()
        Me.IgnoreColumns = New System.Windows.Forms.TextBox()
        Me.addStoredProc = New System.Windows.Forms.TextBox()
        Me.TablenameLabel = New System.Windows.Forms.Label()
        Me.PrimaryKeysLabel = New System.Windows.Forms.Label()
        Me.DatabaseLabel = New System.Windows.Forms.Label()
        Me.IgnoreColumnsLabel = New System.Windows.Forms.Label()
        Me.AdditionalStoredProcLabel = New System.Windows.Forms.Label()
        Me.insertIfMissing = New System.Windows.Forms.CheckBox()
        Me.execOnSave = New System.Windows.Forms.CheckBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.envSel = New System.Windows.Forms.ComboBox()
        Me.DBSeqenceDataGrid = New System.Windows.Forms.DataGridView()
        Me.TargetRangeAddress = New System.Windows.Forms.Label()
        Me.CUDflags = New System.Windows.Forms.CheckBox()
        Me.RepairDBSeqnce = New System.Windows.Forms.TextBox()
        Me.AskForExecute = New System.Windows.Forms.CheckBox()
        Me.IgnoreDataErrors = New System.Windows.Forms.CheckBox()
        Me.EnvironmentLabel = New System.Windows.Forms.Label()
        Me.MoveMenu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.MoveRowUp = New System.Windows.Forms.ToolStripMenuItem()
        Me.MoveRowDown = New System.Windows.Forms.ToolStripMenuItem()
        CType(Me.DBSeqenceDataGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MoveMenu.SuspendLayout()
        Me.SuspendLayout()
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(401, 434)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(67, 23)
        Me.Cancel_Button.TabIndex = 3
        Me.Cancel_Button.Text = "Cancel"
        Me.ToolTip1.SetToolTip(Me.Cancel_Button, "discard changes in DB Modifier Creation")
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.OK_Button.Location = New System.Drawing.Point(361, 434)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(34, 23)
        Me.OK_Button.TabIndex = 2
        Me.OK_Button.Text = "OK"
        Me.ToolTip1.SetToolTip(Me.OK_Button, "use changes done in DB Modifier Creation")
        Me.OK_Button.UseVisualStyleBackColor = True
        '
        'CreateCB
        '
        Me.CreateCB.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CreateCB.Location = New System.Drawing.Point(287, 434)
        Me.CreateCB.Name = "CreateCB"
        Me.CreateCB.Size = New System.Drawing.Size(68, 23)
        Me.CreateCB.TabIndex = 1
        Me.CreateCB.Text = "Create CB"
        Me.ToolTip1.SetToolTip(Me.CreateCB, "Create a Commandbutton for the DB Modifier Definition (max. 5 Buttons possible pe" &
        "r Workbook)")
        '
        'DBModifName
        '
        Me.DBModifName.Location = New System.Drawing.Point(167, 3)
        Me.DBModifName.Name = "DBModifName"
        Me.DBModifName.Size = New System.Drawing.Size(297, 20)
        Me.DBModifName.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.DBModifName, resources.GetString("DBModifName.ToolTip"))
        '
        'NameLabel
        '
        Me.NameLabel.AutoSize = True
        Me.NameLabel.Location = New System.Drawing.Point(9, 6)
        Me.NameLabel.Name = "NameLabel"
        Me.NameLabel.Size = New System.Drawing.Size(91, 13)
        Me.NameLabel.TabIndex = 2
        Me.NameLabel.Text = "DBModifier name:"
        '
        'Tablename
        '
        Me.Tablename.Location = New System.Drawing.Point(167, 55)
        Me.Tablename.Name = "Tablename"
        Me.Tablename.Size = New System.Drawing.Size(297, 20)
        Me.Tablename.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.Tablename, "Database Table, where Data is to be stored")
        '
        'PrimaryKeys
        '
        Me.PrimaryKeys.Location = New System.Drawing.Point(167, 81)
        Me.PrimaryKeys.Name = "PrimaryKeys"
        Me.PrimaryKeys.Size = New System.Drawing.Size(297, 20)
        Me.PrimaryKeys.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.PrimaryKeys, "Number of primary keys in DBMapper datatable (starting from the left)")
        '
        'Database
        '
        Me.Database.Location = New System.Drawing.Point(167, 29)
        Me.Database.Name = "Database"
        Me.Database.Size = New System.Drawing.Size(297, 20)
        Me.Database.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.Database, "Database to store DBMaps Data  into/ do DBActions")
        '
        'IgnoreColumns
        '
        Me.IgnoreColumns.Location = New System.Drawing.Point(167, 107)
        Me.IgnoreColumns.Name = "IgnoreColumns"
        Me.IgnoreColumns.Size = New System.Drawing.Size(297, 20)
        Me.IgnoreColumns.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.IgnoreColumns, "columns to be ignored (e.g. helper columns), comma separated")
        '
        'addStoredProc
        '
        Me.addStoredProc.Location = New System.Drawing.Point(167, 133)
        Me.addStoredProc.Name = "addStoredProc"
        Me.addStoredProc.Size = New System.Drawing.Size(297, 20)
        Me.addStoredProc.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.addStoredProc, "additional stored procedure to be executed after saving")
        '
        'TablenameLabel
        '
        Me.TablenameLabel.AutoSize = True
        Me.TablenameLabel.Location = New System.Drawing.Point(9, 58)
        Me.TablenameLabel.Name = "TablenameLabel"
        Me.TablenameLabel.Size = New System.Drawing.Size(63, 13)
        Me.TablenameLabel.TabIndex = 2
        Me.TablenameLabel.Text = "Tablename:"
        '
        'PrimaryKeysLabel
        '
        Me.PrimaryKeysLabel.AutoSize = True
        Me.PrimaryKeysLabel.Location = New System.Drawing.Point(9, 84)
        Me.PrimaryKeysLabel.Name = "PrimaryKeysLabel"
        Me.PrimaryKeysLabel.Size = New System.Drawing.Size(99, 13)
        Me.PrimaryKeysLabel.TabIndex = 2
        Me.PrimaryKeysLabel.Text = "Primary keys count:"
        '
        'DatabaseLabel
        '
        Me.DatabaseLabel.AutoSize = True
        Me.DatabaseLabel.Location = New System.Drawing.Point(9, 32)
        Me.DatabaseLabel.Name = "DatabaseLabel"
        Me.DatabaseLabel.Size = New System.Drawing.Size(56, 13)
        Me.DatabaseLabel.TabIndex = 2
        Me.DatabaseLabel.Text = "Database:"
        '
        'IgnoreColumnsLabel
        '
        Me.IgnoreColumnsLabel.AutoSize = True
        Me.IgnoreColumnsLabel.Location = New System.Drawing.Point(9, 110)
        Me.IgnoreColumnsLabel.Name = "IgnoreColumnsLabel"
        Me.IgnoreColumnsLabel.Size = New System.Drawing.Size(82, 13)
        Me.IgnoreColumnsLabel.TabIndex = 2
        Me.IgnoreColumnsLabel.Text = "Ignore columns:"
        '
        'AdditionalStoredProcLabel
        '
        Me.AdditionalStoredProcLabel.AutoSize = True
        Me.AdditionalStoredProcLabel.Location = New System.Drawing.Point(9, 136)
        Me.AdditionalStoredProcLabel.Name = "AdditionalStoredProcLabel"
        Me.AdditionalStoredProcLabel.Size = New System.Drawing.Size(139, 13)
        Me.AdditionalStoredProcLabel.TabIndex = 2
        Me.AdditionalStoredProcLabel.Text = "Additional stored procedure:"
        '
        'insertIfMissing
        '
        Me.insertIfMissing.AutoSize = True
        Me.insertIfMissing.Location = New System.Drawing.Point(216, 163)
        Me.insertIfMissing.Name = "insertIfMissing"
        Me.insertIfMissing.Size = New System.Drawing.Size(97, 17)
        Me.insertIfMissing.TabIndex = 9
        Me.insertIfMissing.Text = "Insert if missing"
        Me.ToolTip1.SetToolTip(Me.insertIfMissing, "if set, then insert row into table if primary key is missing there. Default = Fal" &
        "se (only update)")
        Me.insertIfMissing.UseVisualStyleBackColor = True
        '
        'execOnSave
        '
        Me.execOnSave.AutoSize = True
        Me.execOnSave.Location = New System.Drawing.Point(12, 163)
        Me.execOnSave.Name = "execOnSave"
        Me.execOnSave.Size = New System.Drawing.Size(91, 17)
        Me.execOnSave.TabIndex = 7
        Me.execOnSave.Text = "Exec on save"
        Me.ToolTip1.SetToolTip(Me.execOnSave, "should DB Modifier automatically be done on Excel Workbook Saving? (default no)")
        Me.execOnSave.UseVisualStyleBackColor = True
        '
        'envSel
        '
        Me.envSel.FormattingEnabled = True
        Me.envSel.Location = New System.Drawing.Point(351, 159)
        Me.envSel.Name = "envSel"
        Me.envSel.Size = New System.Drawing.Size(113, 21)
        Me.envSel.TabIndex = 10
        Me.ToolTip1.SetToolTip(Me.envSel, "The Environment, where connection id should be taken from (if not existing, take " &
        "from selected Environment in DB Addin General Settings Group)")
        '
        'DBSeqenceDataGrid
        '
        Me.DBSeqenceDataGrid.AllowDrop = True
        Me.DBSeqenceDataGrid.AllowUserToResizeRows = False
        Me.DBSeqenceDataGrid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DBSeqenceDataGrid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DBSeqenceDataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DBSeqenceDataGrid.Location = New System.Drawing.Point(12, 209)
        Me.DBSeqenceDataGrid.MultiSelect = False
        Me.DBSeqenceDataGrid.Name = "DBSeqenceDataGrid"
        Me.DBSeqenceDataGrid.Size = New System.Drawing.Size(452, 219)
        Me.DBSeqenceDataGrid.TabIndex = 13
        Me.ToolTip1.SetToolTip(Me.DBSeqenceDataGrid, "Define the steps for the DB Sequence in the order of their desired execution here" &
        ". Any DBMapper and/or DBAction can be selected.")
        '
        'TargetRangeAddress
        '
        Me.TargetRangeAddress.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TargetRangeAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TargetRangeAddress.ForeColor = System.Drawing.Color.DodgerBlue
        Me.TargetRangeAddress.Location = New System.Drawing.Point(12, 439)
        Me.TargetRangeAddress.Name = "TargetRangeAddress"
        Me.TargetRangeAddress.Size = New System.Drawing.Size(269, 16)
        Me.TargetRangeAddress.TabIndex = 11
        Me.ToolTip1.SetToolTip(Me.TargetRangeAddress, "click to select Target Range with Data for DBMapper or SQL DML for DBAction")
        '
        'CUDflags
        '
        Me.CUDflags.AutoSize = True
        Me.CUDflags.Location = New System.Drawing.Point(12, 186)
        Me.CUDflags.Name = "CUDflags"
        Me.CUDflags.Size = New System.Drawing.Size(84, 17)
        Me.CUDflags.TabIndex = 11
        Me.CUDflags.Text = "C/U/D flags"
        Me.ToolTip1.SetToolTip(Me.CUDflags, "if set, then only insert/update/delete row if special CUDFlags column contains i," &
        " u or d. Default = False (only update)")
        Me.CUDflags.UseVisualStyleBackColor = True
        '
        'RepairDBSeqnce
        '
        Me.RepairDBSeqnce.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RepairDBSeqnce.Location = New System.Drawing.Point(12, 209)
        Me.RepairDBSeqnce.Multiline = True
        Me.RepairDBSeqnce.Name = "RepairDBSeqnce"
        Me.RepairDBSeqnce.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.RepairDBSeqnce.Size = New System.Drawing.Size(456, 219)
        Me.RepairDBSeqnce.TabIndex = 14
        Me.ToolTip1.SetToolTip(Me.RepairDBSeqnce, "use this textbox to repair DB Sequence entries...")
        '
        'AskForExecute
        '
        Me.AskForExecute.AutoSize = True
        Me.AskForExecute.Location = New System.Drawing.Point(102, 163)
        Me.AskForExecute.Name = "AskForExecute"
        Me.AskForExecute.Size = New System.Drawing.Size(108, 17)
        Me.AskForExecute.TabIndex = 8
        Me.AskForExecute.Text = "Ask for execution"
        Me.ToolTip1.SetToolTip(Me.AskForExecute, "ask for confirmation before execution?")
        Me.AskForExecute.UseVisualStyleBackColor = True
        '
        'IgnoreDataErrors
        '
        Me.IgnoreDataErrors.AutoSize = True
        Me.IgnoreDataErrors.Location = New System.Drawing.Point(101, 186)
        Me.IgnoreDataErrors.Name = "IgnoreDataErrors"
        Me.IgnoreDataErrors.Size = New System.Drawing.Size(109, 17)
        Me.IgnoreDataErrors.TabIndex = 12
        Me.IgnoreDataErrors.Text = "Ignore data errors"
        Me.ToolTip1.SetToolTip(Me.IgnoreDataErrors, "if set, don't notify user of error values in cells during update/insert, null val" &
        "ues are used instead")
        Me.IgnoreDataErrors.UseVisualStyleBackColor = True
        '
        'EnvironmentLabel
        '
        Me.EnvironmentLabel.AutoSize = True
        Me.EnvironmentLabel.Location = New System.Drawing.Point(316, 164)
        Me.EnvironmentLabel.Name = "EnvironmentLabel"
        Me.EnvironmentLabel.Size = New System.Drawing.Size(29, 13)
        Me.EnvironmentLabel.TabIndex = 6
        Me.EnvironmentLabel.Text = "Env:"
        '
        'MoveMenu
        '
        Me.MoveMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MoveRowUp, Me.MoveRowDown})
        Me.MoveMenu.Name = "ContextMenuStrip1"
        Me.MoveMenu.Size = New System.Drawing.Size(165, 48)
        '
        'MoveRowUp
        '
        Me.MoveRowUp.Name = "MoveRowUp"
        Me.MoveRowUp.Size = New System.Drawing.Size(164, 22)
        Me.MoveRowUp.Text = "Move Row Up"
        '
        'MoveRowDown
        '
        Me.MoveRowDown.Name = "MoveRowDown"
        Me.MoveRowDown.Size = New System.Drawing.Size(164, 22)
        Me.MoveRowDown.Text = "Move Row Down"
        '
        'DBModifCreate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(479, 469)
        Me.ControlBox = False
        Me.Controls.Add(Me.CreateCB)
        Me.Controls.Add(Me.OK_Button)
        Me.Controls.Add(Me.Cancel_Button)
        Me.Controls.Add(Me.IgnoreDataErrors)
        Me.Controls.Add(Me.EnvironmentLabel)
        Me.Controls.Add(Me.AskForExecute)
        Me.Controls.Add(Me.CUDflags)
        Me.Controls.Add(Me.insertIfMissing)
        Me.Controls.Add(Me.DBSeqenceDataGrid)
        Me.Controls.Add(Me.envSel)
        Me.Controls.Add(Me.execOnSave)
        Me.Controls.Add(Me.AdditionalStoredProcLabel)
        Me.Controls.Add(Me.IgnoreColumnsLabel)
        Me.Controls.Add(Me.DatabaseLabel)
        Me.Controls.Add(Me.PrimaryKeysLabel)
        Me.Controls.Add(Me.TablenameLabel)
        Me.Controls.Add(Me.NameLabel)
        Me.Controls.Add(Me.addStoredProc)
        Me.Controls.Add(Me.IgnoreColumns)
        Me.Controls.Add(Me.Database)
        Me.Controls.Add(Me.PrimaryKeys)
        Me.Controls.Add(Me.Tablename)
        Me.Controls.Add(Me.DBModifName)
        Me.Controls.Add(Me.RepairDBSeqnce)
        Me.Controls.Add(Me.TargetRangeAddress)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(490, 485)
        Me.Name = "DBModifCreate"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        CType(Me.DBSeqenceDataGrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MoveMenu.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents DBModifName As Windows.Forms.TextBox
    Friend WithEvents NameLabel As Windows.Forms.Label
    Friend WithEvents Tablename As Windows.Forms.TextBox
    Friend WithEvents PrimaryKeys As Windows.Forms.TextBox
    Friend WithEvents Database As Windows.Forms.TextBox
    Friend WithEvents IgnoreColumns As Windows.Forms.TextBox
    Friend WithEvents addStoredProc As Windows.Forms.TextBox
    Friend WithEvents TablenameLabel As Windows.Forms.Label
    Friend WithEvents PrimaryKeysLabel As Windows.Forms.Label
    Friend WithEvents DatabaseLabel As Windows.Forms.Label
    Friend WithEvents IgnoreColumnsLabel As Windows.Forms.Label
    Friend WithEvents AdditionalStoredProcLabel As Windows.Forms.Label
    Friend WithEvents insertIfMissing As Windows.Forms.CheckBox
    Friend WithEvents execOnSave As Windows.Forms.CheckBox
    Friend WithEvents ToolTip1 As Windows.Forms.ToolTip
    Friend WithEvents envSel As Windows.Forms.ComboBox
    Friend WithEvents EnvironmentLabel As Windows.Forms.Label
    Friend WithEvents DBSeqenceDataGrid As Windows.Forms.DataGridView
    Friend WithEvents TargetRangeAddress As Windows.Forms.Label
    Friend WithEvents CUDflags As Windows.Forms.CheckBox
    Friend WithEvents RepairDBSeqnce As Windows.Forms.TextBox
    Friend WithEvents AskForExecute As Windows.Forms.CheckBox
    Friend WithEvents CreateCB As Windows.Forms.Button
    Friend WithEvents IgnoreDataErrors As Windows.Forms.CheckBox
    Friend WithEvents MoveMenu As Windows.Forms.ContextMenuStrip
    Friend WithEvents MoveRowUp As Windows.Forms.ToolStripMenuItem
    Friend WithEvents MoveRowDown As Windows.Forms.ToolStripMenuItem
End Class
