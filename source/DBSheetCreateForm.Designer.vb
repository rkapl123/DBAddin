<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class DBSheetCreateForm

#Region "Windows Form Designer generated code "
    Public Shared Function CreateInstance() As DBSheetCreateForm
        Dim theInstance As New DBSheetCreateForm()
        Return theInstance
    End Function
    Private visualControls() As String = New String() {"components", "ToolTipMain", "createQuery", "Query", "testQuery", "WhereClause", "ForTableKey", "Table", "Column", "ForTable", "ForTableLookup", "IsForeign", "addToDBsheetCols", "removeDBsheetCols", "clearAllFields", "outerJoin", "LookupQuery", "regenLookupQueries", "IsPrimary", "moveUp", "moveDown", "addAllFields", "testLookupQuery", "ForDatabase", "Sorting", "cmdAssignDBSheet", "saveDefs", "saveDefsAs", "loadDefs", "DBsheetCols", "Label3", "LTable", "LColumn", "LForTableKey", "LForTableLookup", "LLookupQuery", "LForDatabase", "LForTable", "Label28"}
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Public ToolTipMain As System.Windows.Forms.ToolTip
    Public WithEvents createQuery As System.Windows.Forms.Button
    Public WithEvents Query As System.Windows.Forms.TextBox
    Public WithEvents testQuery As System.Windows.Forms.Button
    Public WithEvents WhereClause As System.Windows.Forms.TextBox
    Public WithEvents Table As System.Windows.Forms.ComboBox
    Public WithEvents clearAllFields As System.Windows.Forms.Button
    Public WithEvents addAllFields As System.Windows.Forms.Button
    Public WithEvents saveDefs As System.Windows.Forms.Button
    Public WithEvents saveDefsAs As System.Windows.Forms.Button
    Public WithEvents loadDefs As System.Windows.Forms.Button
    Public WithEvents LWhereParamClause As System.Windows.Forms.Label
    Public WithEvents LTable As System.Windows.Forms.Label

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DBSheetCreateForm))
        Me.ToolTipMain = New System.Windows.Forms.ToolTip(Me.components)
        Me.createQuery = New System.Windows.Forms.Button()
        Me.Query = New System.Windows.Forms.TextBox()
        Me.testQuery = New System.Windows.Forms.Button()
        Me.WhereClause = New System.Windows.Forms.TextBox()
        Me.Table = New System.Windows.Forms.ComboBox()
        Me.clearAllFields = New System.Windows.Forms.Button()
        Me.addAllFields = New System.Windows.Forms.Button()
        Me.saveDefs = New System.Windows.Forms.Button()
        Me.saveDefsAs = New System.Windows.Forms.Button()
        Me.loadDefs = New System.Windows.Forms.Button()
        Me.Password = New System.Windows.Forms.TextBox()
        Me.Database = New System.Windows.Forms.ComboBox()
        Me.DBSheetCols = New System.Windows.Forms.DataGridView()
        Me.assignDBSheet = New System.Windows.Forms.Button()
        Me.LWhereParamClause = New System.Windows.Forms.Label()
        Me.LTable = New System.Windows.Forms.Label()
        Me.DBSheetColsMoveMenu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.MoveRowUp = New System.Windows.Forms.ToolStripMenuItem()
        Me.MoveRowDown = New System.Windows.Forms.ToolStripMenuItem()
        Me.LDatabase = New System.Windows.Forms.Label()
        Me.LPwd = New System.Windows.Forms.Label()
        Me.LQuery = New System.Windows.Forms.Label()
        Me.DBSheetColsLookupMenu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.RegenerateThisLookupQuery = New System.Windows.Forms.ToolStripMenuItem()
        Me.RegenerateAllLookupQueries = New System.Windows.Forms.ToolStripMenuItem()
        Me.TestLookupQuery = New System.Windows.Forms.ToolStripMenuItem()
        Me.RemoveLookupQueryTest = New System.Windows.Forms.ToolStripMenuItem()
        Me.CurrentFileLinkLabel = New System.Windows.Forms.LinkLabel()
        Me.Lenvironment = New System.Windows.Forms.Label()
        Me.DBSheetColsForDatabases = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.Environment = New System.Windows.Forms.ComboBox()
        CType(Me.DBSheetCols, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.DBSheetColsMoveMenu.SuspendLayout()
        Me.DBSheetColsLookupMenu.SuspendLayout()
        Me.SuspendLayout()
        '
        'createQuery
        '
        Me.createQuery.AllowDrop = True
        Me.createQuery.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.createQuery.BackColor = System.Drawing.SystemColors.Control
        Me.createQuery.ForeColor = System.Drawing.SystemColors.ControlText
        Me.createQuery.Location = New System.Drawing.Point(428, 337)
        Me.createQuery.Name = "createQuery"
        Me.createQuery.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.createQuery.Size = New System.Drawing.Size(133, 24)
        Me.createQuery.TabIndex = 14
        Me.createQuery.Text = "&create DBSheet query"
        Me.createQuery.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolTipMain.SetToolTip(Me.createQuery, "create query for DBsheet and foreign lookup informations")
        Me.createQuery.UseVisualStyleBackColor = False
        '
        'Query
        '
        Me.Query.AcceptsReturn = True
        Me.Query.AllowDrop = True
        Me.Query.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Query.BackColor = System.Drawing.SystemColors.Window
        Me.Query.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Query.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Query.Location = New System.Drawing.Point(5, 367)
        Me.Query.MaxLength = 0
        Me.Query.Multiline = True
        Me.Query.Name = "Query"
        Me.Query.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Query.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.Query.Size = New System.Drawing.Size(840, 164)
        Me.Query.TabIndex = 17
        Me.Query.Text = "__"
        Me.ToolTipMain.SetToolTip(Me.Query, "Query: If needed, modify the query for the DBSheet data being displayed. Attentio" &
        "n: create DBSheet query will destroy all custom information here !!")
        '
        'testQuery
        '
        Me.testQuery.AllowDrop = True
        Me.testQuery.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.testQuery.BackColor = System.Drawing.SystemColors.Control
        Me.testQuery.ForeColor = System.Drawing.SystemColors.ControlText
        Me.testQuery.Location = New System.Drawing.Point(568, 337)
        Me.testQuery.Name = "testQuery"
        Me.testQuery.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.testQuery.Size = New System.Drawing.Size(121, 24)
        Me.testQuery.TabIndex = 15
        Me.testQuery.Text = "&test DBSheet Query"
        Me.testQuery.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolTipMain.SetToolTip(Me.testQuery, "test the DBSheet query  in a new Excel Sheet...")
        Me.testQuery.UseVisualStyleBackColor = False
        '
        'WhereClause
        '
        Me.WhereClause.AcceptsReturn = True
        Me.WhereClause.AllowDrop = True
        Me.WhereClause.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.WhereClause.BackColor = System.Drawing.SystemColors.Window
        Me.WhereClause.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.WhereClause.ForeColor = System.Drawing.SystemColors.WindowText
        Me.WhereClause.Location = New System.Drawing.Point(851, 367)
        Me.WhereClause.MaxLength = 0
        Me.WhereClause.Multiline = True
        Me.WhereClause.Name = "WhereClause"
        Me.WhereClause.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.WhereClause.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.WhereClause.Size = New System.Drawing.Size(250, 164)
        Me.WhereClause.TabIndex = 18
        Me.ToolTipMain.SetToolTip(Me.WhereClause, "Where Parameter Clause: Restrict the data displayed with the Where part of an SQL" &
        " Select statement (enter without ""WHERE"" !).")
        '
        'Table
        '
        Me.Table.AllowDrop = True
        Me.Table.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Table.BackColor = System.Drawing.SystemColors.Window
        Me.Table.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Table.DropDownWidth = 300
        Me.Table.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Table.IntegralHeight = False
        Me.Table.Location = New System.Drawing.Point(907, 41)
        Me.Table.MaxDropDownItems = 50
        Me.Table.Name = "Table"
        Me.Table.Size = New System.Drawing.Size(194, 21)
        Me.Table.Sorted = True
        Me.Table.TabIndex = 6
        Me.ToolTipMain.SetToolTip(Me.Table, "Main Table of DBSheet: on this table the created DBSheet definition will allow ed" &
        "iting data")
        '
        'clearAllFields
        '
        Me.clearAllFields.AllowDrop = True
        Me.clearAllFields.BackColor = System.Drawing.SystemColors.Control
        Me.clearAllFields.ForeColor = System.Drawing.SystemColors.ControlText
        Me.clearAllFields.Location = New System.Drawing.Point(114, 38)
        Me.clearAllFields.Name = "clearAllFields"
        Me.clearAllFields.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.clearAllFields.Size = New System.Drawing.Size(132, 24)
        Me.clearAllFields.TabIndex = 8
        Me.clearAllFields.Text = "reset DBSheet definition"
        Me.clearAllFields.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolTipMain.SetToolTip(Me.clearAllFields, "clear all DBSheet definitions and enable database/table choice")
        Me.clearAllFields.UseVisualStyleBackColor = False
        '
        'addAllFields
        '
        Me.addAllFields.AllowDrop = True
        Me.addAllFields.BackColor = System.Drawing.SystemColors.Control
        Me.addAllFields.ForeColor = System.Drawing.SystemColors.ControlText
        Me.addAllFields.Location = New System.Drawing.Point(5, 38)
        Me.addAllFields.Name = "addAllFields"
        Me.addAllFields.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.addAllFields.Size = New System.Drawing.Size(103, 24)
        Me.addAllFields.TabIndex = 7
        Me.addAllFields.Text = "add all &Fields"
        Me.addAllFields.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolTipMain.SetToolTip(Me.addAllFields, "add all columns to DBsheet definition")
        Me.addAllFields.UseVisualStyleBackColor = False
        '
        'saveDefs
        '
        Me.saveDefs.AllowDrop = True
        Me.saveDefs.BackColor = System.Drawing.SystemColors.Control
        Me.saveDefs.ForeColor = System.Drawing.SystemColors.ControlText
        Me.saveDefs.Location = New System.Drawing.Point(114, 8)
        Me.saveDefs.Name = "saveDefs"
        Me.saveDefs.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.saveDefs.Size = New System.Drawing.Size(109, 24)
        Me.saveDefs.TabIndex = 4
        Me.saveDefs.Text = "&save DBSheet def"
        Me.saveDefs.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolTipMain.SetToolTip(Me.saveDefs, "save current DB Sheet definition context ")
        Me.saveDefs.UseVisualStyleBackColor = False
        '
        'saveDefsAs
        '
        Me.saveDefsAs.AllowDrop = True
        Me.saveDefsAs.BackColor = System.Drawing.SystemColors.Control
        Me.saveDefsAs.ForeColor = System.Drawing.SystemColors.ControlText
        Me.saveDefsAs.Location = New System.Drawing.Point(229, 8)
        Me.saveDefsAs.Name = "saveDefsAs"
        Me.saveDefsAs.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.saveDefsAs.Size = New System.Drawing.Size(130, 24)
        Me.saveDefsAs.TabIndex = 5
        Me.saveDefsAs.Text = "save DBSheet def As..."
        Me.saveDefsAs.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolTipMain.SetToolTip(Me.saveDefsAs, "save current DB Sheet definition context As...")
        Me.saveDefsAs.UseVisualStyleBackColor = False
        '
        'loadDefs
        '
        Me.loadDefs.AllowDrop = True
        Me.loadDefs.BackColor = System.Drawing.SystemColors.Control
        Me.loadDefs.ForeColor = System.Drawing.SystemColors.ControlText
        Me.loadDefs.Location = New System.Drawing.Point(5, 8)
        Me.loadDefs.Name = "loadDefs"
        Me.loadDefs.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.loadDefs.Size = New System.Drawing.Size(103, 24)
        Me.loadDefs.TabIndex = 3
        Me.loadDefs.Text = "&load DBSheet def"
        Me.loadDefs.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolTipMain.SetToolTip(Me.loadDefs, "load DB Sheet definitions into current context")
        Me.loadDefs.UseVisualStyleBackColor = False
        '
        'Password
        '
        Me.Password.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Password.Location = New System.Drawing.Point(708, 10)
        Me.Password.Name = "Password"
        Me.Password.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.Password.Size = New System.Drawing.Size(134, 20)
        Me.Password.TabIndex = 0
        Me.ToolTipMain.SetToolTip(Me.Password, "enter the password for the user required to access schema information (given in D" &
        "BSheetConnString)")
        '
        'Database
        '
        Me.Database.AllowDrop = True
        Me.Database.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Database.BackColor = System.Drawing.SystemColors.Window
        Me.Database.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Database.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Database.Location = New System.Drawing.Point(907, 10)
        Me.Database.MaxDropDownItems = 50
        Me.Database.Name = "Database"
        Me.Database.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Database.Size = New System.Drawing.Size(194, 21)
        Me.Database.Sorted = True
        Me.Database.TabIndex = 2
        Me.ToolTipMain.SetToolTip(Me.Database, "choose Database to select tables from.")
        '
        'DBSheetCols
        '
        Me.DBSheetCols.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DBSheetCols.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DBSheetCols.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DBSheetCols.Location = New System.Drawing.Point(5, 68)
        Me.DBSheetCols.MultiSelect = False
        Me.DBSheetCols.Name = "DBSheetCols"
        Me.DBSheetCols.Size = New System.Drawing.Size(1096, 263)
        Me.DBSheetCols.TabIndex = 13
        Me.ToolTipMain.SetToolTip(Me.DBSheetCols, "Select columns (fields) adding possible foreign key lookup information in foreign" &
        " tables")
        '
        'assignDBSheet
        '
        Me.assignDBSheet.AllowDrop = True
        Me.assignDBSheet.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.assignDBSheet.Location = New System.Drawing.Point(696, 338)
        Me.assignDBSheet.Name = "assignDBSheet"
        Me.assignDBSheet.Size = New System.Drawing.Size(98, 23)
        Me.assignDBSheet.TabIndex = 16
        Me.assignDBSheet.Text = "assign DBSheet"
        Me.ToolTipMain.SetToolTip(Me.assignDBSheet, "assigns the current definition to active sheet.")
        Me.assignDBSheet.UseVisualStyleBackColor = False
        '
        'LWhereParamClause
        '
        Me.LWhereParamClause.AllowDrop = True
        Me.LWhereParamClause.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LWhereParamClause.AutoSize = True
        Me.LWhereParamClause.BackColor = System.Drawing.Color.Transparent
        Me.LWhereParamClause.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LWhereParamClause.Location = New System.Drawing.Point(848, 348)
        Me.LWhereParamClause.Name = "LWhereParamClause"
        Me.LWhereParamClause.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LWhereParamClause.Size = New System.Drawing.Size(131, 13)
        Me.LWhereParamClause.TabIndex = 100
        Me.LWhereParamClause.Text = "Where Parameter Clause:"
        '
        'LTable
        '
        Me.LTable.AllowDrop = True
        Me.LTable.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LTable.AutoSize = True
        Me.LTable.BackColor = System.Drawing.Color.Transparent
        Me.LTable.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LTable.Location = New System.Drawing.Point(868, 44)
        Me.LTable.Name = "LTable"
        Me.LTable.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LTable.Size = New System.Drawing.Size(37, 13)
        Me.LTable.TabIndex = 100
        Me.LTable.Text = "Table:"
        '
        'DBSheetColsMoveMenu
        '
        Me.DBSheetColsMoveMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MoveRowUp, Me.MoveRowDown})
        Me.DBSheetColsMoveMenu.Name = "DBSheetColsContextMenu"
        Me.DBSheetColsMoveMenu.Size = New System.Drawing.Size(161, 48)
        '
        'MoveRowUp
        '
        Me.MoveRowUp.Name = "MoveRowUp"
        Me.MoveRowUp.Size = New System.Drawing.Size(160, 22)
        Me.MoveRowUp.Text = "move row up"
        '
        'MoveRowDown
        '
        Me.MoveRowDown.Name = "MoveRowDown"
        Me.MoveRowDown.Size = New System.Drawing.Size(160, 22)
        Me.MoveRowDown.Text = "move row down"
        '
        'LDatabase
        '
        Me.LDatabase.AllowDrop = True
        Me.LDatabase.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LDatabase.AutoSize = True
        Me.LDatabase.BackColor = System.Drawing.Color.Transparent
        Me.LDatabase.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LDatabase.Location = New System.Drawing.Point(848, 14)
        Me.LDatabase.Name = "LDatabase"
        Me.LDatabase.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LDatabase.Size = New System.Drawing.Size(57, 13)
        Me.LDatabase.TabIndex = 100
        Me.LDatabase.Text = "Database:"
        '
        'LPwd
        '
        Me.LPwd.AllowDrop = True
        Me.LPwd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LPwd.AutoSize = True
        Me.LPwd.BackColor = System.Drawing.Color.Transparent
        Me.LPwd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LPwd.Location = New System.Drawing.Point(675, 14)
        Me.LPwd.Name = "LPwd"
        Me.LPwd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LPwd.Size = New System.Drawing.Size(31, 13)
        Me.LPwd.TabIndex = 100
        Me.LPwd.Text = "Pwd:"
        '
        'LQuery
        '
        Me.LQuery.AllowDrop = True
        Me.LQuery.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LQuery.AutoSize = True
        Me.LQuery.BackColor = System.Drawing.Color.Transparent
        Me.LQuery.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LQuery.Location = New System.Drawing.Point(12, 348)
        Me.LQuery.Name = "LQuery"
        Me.LQuery.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LQuery.Size = New System.Drawing.Size(85, 13)
        Me.LQuery.TabIndex = 100
        Me.LQuery.Text = "DBSheet Query:"
        '
        'DBSheetColsLookupMenu
        '
        Me.DBSheetColsLookupMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.RegenerateThisLookupQuery, Me.RegenerateAllLookupQueries, Me.TestLookupQuery, Me.RemoveLookupQueryTest})
        Me.DBSheetColsLookupMenu.Name = "DBSheetColsContextMenu"
        Me.DBSheetColsLookupMenu.Size = New System.Drawing.Size(227, 92)
        '
        'RegenerateThisLookupQuery
        '
        Me.RegenerateThisLookupQuery.Name = "RegenerateThisLookupQuery"
        Me.RegenerateThisLookupQuery.Size = New System.Drawing.Size(226, 22)
        Me.RegenerateThisLookupQuery.Text = "regenerate this lookup query"
        '
        'RegenerateAllLookupQueries
        '
        Me.RegenerateAllLookupQueries.Name = "RegenerateAllLookupQueries"
        Me.RegenerateAllLookupQueries.Size = New System.Drawing.Size(226, 22)
        Me.RegenerateAllLookupQueries.Text = "regenerate all lookup queries"
        '
        'TestLookupQuery
        '
        Me.TestLookupQuery.Name = "TestLookupQuery"
        Me.TestLookupQuery.Size = New System.Drawing.Size(226, 22)
        Me.TestLookupQuery.Text = "test lookup query"
        '
        'RemoveLookupQueryTest
        '
        Me.RemoveLookupQueryTest.Name = "RemoveLookupQueryTest"
        Me.RemoveLookupQueryTest.Size = New System.Drawing.Size(226, 22)
        Me.RemoveLookupQueryTest.Text = "remove lookup query test"
        '
        'CurrentFileLinkLabel
        '
        Me.CurrentFileLinkLabel.AutoEllipsis = True
        Me.CurrentFileLinkLabel.Location = New System.Drawing.Point(252, 44)
        Me.CurrentFileLinkLabel.MaximumSize = New System.Drawing.Size(430, 18)
        Me.CurrentFileLinkLabel.Name = "CurrentFileLinkLabel"
        Me.CurrentFileLinkLabel.Size = New System.Drawing.Size(417, 18)
        Me.CurrentFileLinkLabel.TabIndex = 100
        '
        'Lenvironment
        '
        Me.Lenvironment.AllowDrop = True
        Me.Lenvironment.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lenvironment.AutoSize = True
        Me.Lenvironment.BackColor = System.Drawing.Color.Transparent
        Me.Lenvironment.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Lenvironment.Location = New System.Drawing.Point(675, 44)
        Me.Lenvironment.Name = "Lenvironment"
        Me.Lenvironment.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Lenvironment.Size = New System.Drawing.Size(71, 13)
        Me.Lenvironment.TabIndex = 100
        Me.Lenvironment.Text = "Environment:"
        '
        'DBSheetColsForDatabases
        '
        Me.DBSheetColsForDatabases.Name = "DBSheetColsContextMenu"
        Me.DBSheetColsForDatabases.Size = New System.Drawing.Size(61, 4)
        '
        'Environment
        '
        Me.Environment.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Environment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Environment.FormattingEnabled = True
        Me.Environment.Location = New System.Drawing.Point(752, 41)
        Me.Environment.Name = "Environment"
        Me.Environment.Size = New System.Drawing.Size(90, 21)
        Me.Environment.TabIndex = 1
        Me.ToolTipMain.SetToolTip(Me.Environment, "Environment, can be used to switch environment taken from Ribbon Menu")
        '
        'DBSheetCreateForm
        '
        Me.AllowDrop = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(1113, 544)
        Me.Controls.Add(Me.Environment)
        Me.Controls.Add(Me.Lenvironment)
        Me.Controls.Add(Me.assignDBSheet)
        Me.Controls.Add(Me.CurrentFileLinkLabel)
        Me.Controls.Add(Me.LQuery)
        Me.Controls.Add(Me.LPwd)
        Me.Controls.Add(Me.LDatabase)
        Me.Controls.Add(Me.Database)
        Me.Controls.Add(Me.Password)
        Me.Controls.Add(Me.createQuery)
        Me.Controls.Add(Me.Query)
        Me.Controls.Add(Me.testQuery)
        Me.Controls.Add(Me.WhereClause)
        Me.Controls.Add(Me.Table)
        Me.Controls.Add(Me.clearAllFields)
        Me.Controls.Add(Me.addAllFields)
        Me.Controls.Add(Me.saveDefs)
        Me.Controls.Add(Me.saveDefsAs)
        Me.Controls.Add(Me.loadDefs)
        Me.Controls.Add(Me.LWhereParamClause)
        Me.Controls.Add(Me.LTable)
        Me.Controls.Add(Me.DBSheetCols)
        Me.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MinimumSize = New System.Drawing.Size(1129, 582)
        Me.Name = "DBSheetCreateForm"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "DB Sheet creation"
        CType(Me.DBSheetCols, System.ComponentModel.ISupportInitialize).EndInit()
        Me.DBSheetColsMoveMenu.ResumeLayout(False)
        Me.DBSheetColsLookupMenu.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DBSheetCols As Windows.Forms.DataGridView
    Friend WithEvents Password As Windows.Forms.TextBox
    Public WithEvents Database As Windows.Forms.ComboBox
    Public WithEvents LDatabase As Windows.Forms.Label
    Public WithEvents LPwd As Windows.Forms.Label
    Public WithEvents LQuery As Windows.Forms.Label
    Friend WithEvents DBSheetColsMoveMenu As Windows.Forms.ContextMenuStrip
    Friend WithEvents MoveRowUp As Windows.Forms.ToolStripMenuItem
    Friend WithEvents MoveRowDown As Windows.Forms.ToolStripMenuItem
    Friend WithEvents DBSheetColsLookupMenu As Windows.Forms.ContextMenuStrip
    Friend WithEvents RegenerateThisLookupQuery As Windows.Forms.ToolStripMenuItem
    Friend WithEvents TestLookupQuery As Windows.Forms.ToolStripMenuItem
    Friend WithEvents RemoveLookupQueryTest As Windows.Forms.ToolStripMenuItem
    Friend WithEvents CurrentFileLinkLabel As Windows.Forms.LinkLabel
    Friend WithEvents RegenerateAllLookupQueries As Windows.Forms.ToolStripMenuItem
    Public WithEvents Lenvironment As Windows.Forms.Label
    Friend WithEvents DBSheetColsForDatabases As Windows.Forms.ContextMenuStrip
    Friend WithEvents assignDBSheet As Windows.Forms.Button
    Friend WithEvents Environment As Windows.Forms.ComboBox
#End Region
End Class