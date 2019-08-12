<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DBMapperCreate
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
        Me.components = New System.ComponentModel.Container()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.DBMapperName = New System.Windows.Forms.TextBox()
        Me.NameLabel = New System.Windows.Forms.Label()
        Me.Tablename = New System.Windows.Forms.TextBox()
        Me.PrimaryKeys = New System.Windows.Forms.TextBox()
        Me.Database = New System.Windows.Forms.TextBox()
        Me.IgnoreColumns = New System.Windows.Forms.TextBox()
        Me.addStoredProc = New System.Windows.Forms.TextBox()
        Me.TablenameLabel = New System.Windows.Forms.Label()
        Me.PrimaryKeysLabel = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.IgnoreColumnsLabel = New System.Windows.Forms.Label()
        Me.AdditionalStoredProcLabel = New System.Windows.Forms.Label()
        Me.insertIfMissing = New System.Windows.Forms.CheckBox()
        Me.storeDBMapOnSave = New System.Windows.Forms.CheckBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.envSel = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(275, 238)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Location = New System.Drawing.Point(3, 3)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(67, 23)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "OK"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(76, 3)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(67, 23)
        Me.Cancel_Button.TabIndex = 0
        Me.Cancel_Button.Text = "Abbrechen"
        '
        'DBMapperName
        '
        Me.DBMapperName.Location = New System.Drawing.Point(161, 25)
        Me.DBMapperName.Name = "DBMapperName"
        Me.DBMapperName.Size = New System.Drawing.Size(259, 20)
        Me.DBMapperName.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.DBMapperName, "Enter the name for the selected Range that will be used to identify the DBMap in " &
        "the ""store DBMapper Data"" Group dropdowns." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "If no name is given here, then Unnam" &
        "edDBMapper will be used to identify it.")
        '
        'NameLabel
        '
        Me.NameLabel.AutoSize = True
        Me.NameLabel.Location = New System.Drawing.Point(12, 28)
        Me.NameLabel.Name = "NameLabel"
        Me.NameLabel.Size = New System.Drawing.Size(92, 13)
        Me.NameLabel.TabIndex = 2
        Me.NameLabel.Text = "DBMapper Name:"
        '
        'Tablename
        '
        Me.Tablename.Location = New System.Drawing.Point(161, 77)
        Me.Tablename.Name = "Tablename"
        Me.Tablename.Size = New System.Drawing.Size(259, 20)
        Me.Tablename.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.Tablename, "Database Table, where Data is to be stored")
        '
        'PrimaryKeys
        '
        Me.PrimaryKeys.Location = New System.Drawing.Point(161, 103)
        Me.PrimaryKeys.Name = "PrimaryKeys"
        Me.PrimaryKeys.Size = New System.Drawing.Size(259, 20)
        Me.PrimaryKeys.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.PrimaryKeys, "String containing primary Key names for updating table data, comma separated")
        '
        'Database
        '
        Me.Database.Location = New System.Drawing.Point(161, 51)
        Me.Database.Name = "Database"
        Me.Database.Size = New System.Drawing.Size(259, 20)
        Me.Database.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.Database, "Database to store DBMaps Data into")
        '
        'IgnoreColumns
        '
        Me.IgnoreColumns.Location = New System.Drawing.Point(161, 129)
        Me.IgnoreColumns.Name = "IgnoreColumns"
        Me.IgnoreColumns.Size = New System.Drawing.Size(259, 20)
        Me.IgnoreColumns.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.IgnoreColumns, "columns to be ignored (e.g. helper columns), comma separated")
        '
        'addStoredProc
        '
        Me.addStoredProc.Location = New System.Drawing.Point(161, 155)
        Me.addStoredProc.Name = "addStoredProc"
        Me.addStoredProc.Size = New System.Drawing.Size(259, 20)
        Me.addStoredProc.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.addStoredProc, "additional stored procedure to be executed after saving")
        '
        'TablenameLabel
        '
        Me.TablenameLabel.AutoSize = True
        Me.TablenameLabel.Location = New System.Drawing.Point(12, 80)
        Me.TablenameLabel.Name = "TablenameLabel"
        Me.TablenameLabel.Size = New System.Drawing.Size(63, 13)
        Me.TablenameLabel.TabIndex = 2
        Me.TablenameLabel.Text = "Tablename:"
        '
        'PrimaryKeysLabel
        '
        Me.PrimaryKeysLabel.AutoSize = True
        Me.PrimaryKeysLabel.Location = New System.Drawing.Point(12, 106)
        Me.PrimaryKeysLabel.Name = "PrimaryKeysLabel"
        Me.PrimaryKeysLabel.Size = New System.Drawing.Size(70, 13)
        Me.PrimaryKeysLabel.TabIndex = 2
        Me.PrimaryKeysLabel.Text = "Primary Keys:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 54)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 13)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Database:"
        '
        'IgnoreColumnsLabel
        '
        Me.IgnoreColumnsLabel.AutoSize = True
        Me.IgnoreColumnsLabel.Location = New System.Drawing.Point(12, 132)
        Me.IgnoreColumnsLabel.Name = "IgnoreColumnsLabel"
        Me.IgnoreColumnsLabel.Size = New System.Drawing.Size(83, 13)
        Me.IgnoreColumnsLabel.TabIndex = 2
        Me.IgnoreColumnsLabel.Text = "Ignore Columns:"
        '
        'AdditionalStoredProcLabel
        '
        Me.AdditionalStoredProcLabel.AutoSize = True
        Me.AdditionalStoredProcLabel.Location = New System.Drawing.Point(12, 158)
        Me.AdditionalStoredProcLabel.Name = "AdditionalStoredProcLabel"
        Me.AdditionalStoredProcLabel.Size = New System.Drawing.Size(142, 13)
        Me.AdditionalStoredProcLabel.TabIndex = 2
        Me.AdditionalStoredProcLabel.Text = "Additional Stored Procedure:"
        '
        'insertIfMissing
        '
        Me.insertIfMissing.AutoSize = True
        Me.insertIfMissing.Location = New System.Drawing.Point(15, 189)
        Me.insertIfMissing.Name = "insertIfMissing"
        Me.insertIfMissing.Size = New System.Drawing.Size(99, 17)
        Me.insertIfMissing.TabIndex = 7
        Me.insertIfMissing.Text = "Insert If Missing"
        Me.ToolTip1.SetToolTip(Me.insertIfMissing, "if set, then insert row into table if primary key is missing there. Default = Fal" &
        "se (only update)")
        Me.insertIfMissing.UseVisualStyleBackColor = True
        '
        'storeDBMapOnSave
        '
        Me.storeDBMapOnSave.AutoSize = True
        Me.storeDBMapOnSave.Location = New System.Drawing.Point(14, 212)
        Me.storeDBMapOnSave.Name = "storeDBMapOnSave"
        Me.storeDBMapOnSave.Size = New System.Drawing.Size(133, 17)
        Me.storeDBMapOnSave.TabIndex = 8
        Me.storeDBMapOnSave.Text = "Store DBMap on Save"
        Me.ToolTip1.SetToolTip(Me.storeDBMapOnSave, "should DBMap also be saved on Excel Workbook Saving? (default no)")
        Me.storeDBMapOnSave.UseVisualStyleBackColor = True
        '
        'envSel
        '
        Me.envSel.FormattingEnabled = True
        Me.envSel.Location = New System.Drawing.Point(297, 187)
        Me.envSel.Name = "envSel"
        Me.envSel.Size = New System.Drawing.Size(123, 21)
        Me.envSel.TabIndex = 9
        Me.ToolTip1.SetToolTip(Me.envSel, "The Environment, where connection id should be taken from (if not existing, take " &
        "from selected Environment in DB Addin General Settings Group)")
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(211, 190)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(69, 13)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "Environment:"
        '
        'DBMapperCreate
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(433, 279)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.envSel)
        Me.Controls.Add(Me.storeDBMapOnSave)
        Me.Controls.Add(Me.insertIfMissing)
        Me.Controls.Add(Me.AdditionalStoredProcLabel)
        Me.Controls.Add(Me.IgnoreColumnsLabel)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.PrimaryKeysLabel)
        Me.Controls.Add(Me.TablenameLabel)
        Me.Controls.Add(Me.NameLabel)
        Me.Controls.Add(Me.addStoredProc)
        Me.Controls.Add(Me.IgnoreColumns)
        Me.Controls.Add(Me.Database)
        Me.Controls.Add(Me.PrimaryKeys)
        Me.Controls.Add(Me.Tablename)
        Me.Controls.Add(Me.DBMapperName)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DBMapperCreate"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents DBMapperName As Windows.Forms.TextBox
    Friend WithEvents NameLabel As Windows.Forms.Label
    Friend WithEvents Tablename As Windows.Forms.TextBox
    Friend WithEvents PrimaryKeys As Windows.Forms.TextBox
    Friend WithEvents Database As Windows.Forms.TextBox
    Friend WithEvents IgnoreColumns As Windows.Forms.TextBox
    Friend WithEvents addStoredProc As Windows.Forms.TextBox
    Friend WithEvents TablenameLabel As Windows.Forms.Label
    Friend WithEvents PrimaryKeysLabel As Windows.Forms.Label
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents IgnoreColumnsLabel As Windows.Forms.Label
    Friend WithEvents AdditionalStoredProcLabel As Windows.Forms.Label
    Friend WithEvents insertIfMissing As Windows.Forms.CheckBox
    Friend WithEvents storeDBMapOnSave As Windows.Forms.CheckBox
    Friend WithEvents ToolTip1 As Windows.Forms.ToolTip
    Friend WithEvents envSel As Windows.Forms.ComboBox
    Friend WithEvents Label7 As Windows.Forms.Label
End Class
