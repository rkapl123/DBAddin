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
    Public WithEvents regenLookupQueries As System.Windows.Forms.Button
    Public WithEvents moveUp As System.Windows.Forms.Button
    Public WithEvents moveDown As System.Windows.Forms.Button
    Public WithEvents addAllFields As System.Windows.Forms.Button
    Public WithEvents testLookupQuery As System.Windows.Forms.Button
    Public WithEvents saveDefs As System.Windows.Forms.Button
    Public WithEvents saveDefsAs As System.Windows.Forms.Button
    Public WithEvents loadDefs As System.Windows.Forms.Button
    Public WithEvents Label3 As System.Windows.Forms.Label
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
        Me.regenLookupQueries = New System.Windows.Forms.Button()
        Me.moveUp = New System.Windows.Forms.Button()
        Me.moveDown = New System.Windows.Forms.Button()
        Me.addAllFields = New System.Windows.Forms.Button()
        Me.testLookupQuery = New System.Windows.Forms.Button()
        Me.saveDefs = New System.Windows.Forms.Button()
        Me.saveDefsAs = New System.Windows.Forms.Button()
        Me.loadDefs = New System.Windows.Forms.Button()
        Me.Password = New System.Windows.Forms.TextBox()
        Me.Database = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.LTable = New System.Windows.Forms.Label()
        Me.DBSheetCols = New System.Windows.Forms.DataGridView()
        Me.LDatabase = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.LQuery = New System.Windows.Forms.Label()
        CType(Me.DBSheetCols, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'createQuery
        '
        Me.createQuery.AllowDrop = True
        Me.createQuery.BackColor = System.Drawing.SystemColors.Control
        Me.createQuery.ForeColor = System.Drawing.SystemColors.ControlText
        Me.createQuery.Location = New System.Drawing.Point(428, 337)
        Me.createQuery.Name = "createQuery"
        Me.createQuery.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.createQuery.Size = New System.Drawing.Size(133, 24)
        Me.createQuery.TabIndex = 12
        Me.createQuery.Text = "create DBSheet &query"
        Me.createQuery.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolTipMain.SetToolTip(Me.createQuery, "create query for DBsheet and foreign lookup informations")
        Me.createQuery.UseVisualStyleBackColor = False
        '
        'Query
        '
        Me.Query.AcceptsReturn = True
        Me.Query.AllowDrop = True
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
        Me.Query.TabIndex = 14
        Me.Query.Text = "__"
        Me.ToolTipMain.SetToolTip(Me.Query, "Query: If needed, modify the query for the DBSheet data being displayed. Attentio" &
        "n: create DBSheet query will destroy all custom information here !!")
        '
        'testQuery
        '
        Me.testQuery.AllowDrop = True
        Me.testQuery.BackColor = System.Drawing.SystemColors.Control
        Me.testQuery.ForeColor = System.Drawing.SystemColors.ControlText
        Me.testQuery.Location = New System.Drawing.Point(568, 337)
        Me.testQuery.Name = "testQuery"
        Me.testQuery.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.testQuery.Size = New System.Drawing.Size(121, 24)
        Me.testQuery.TabIndex = 13
        Me.testQuery.Text = "&test DBSheet Query"
        Me.testQuery.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolTipMain.SetToolTip(Me.testQuery, "test the DBSheet query  in a new Excel Sheet...")
        Me.testQuery.UseVisualStyleBackColor = False
        '
        'WhereClause
        '
        Me.WhereClause.AcceptsReturn = True
        Me.WhereClause.AllowDrop = True
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
        Me.WhereClause.TabIndex = 15
        Me.ToolTipMain.SetToolTip(Me.WhereClause, "Where Parameter Clause: Restrict the data displayed with the Where part of an SQL" &
        " Select statement (enter without ""WHERE"" !).")
        '
        'Table
        '
        Me.Table.AllowDrop = True
        Me.Table.BackColor = System.Drawing.SystemColors.Window
        Me.Table.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Table.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Table.Location = New System.Drawing.Point(51, 37)
        Me.Table.Name = "Table"
        Me.Table.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Table.Size = New System.Drawing.Size(190, 21)
        Me.Table.Sorted = True
        Me.Table.TabIndex = 5
        Me.ToolTipMain.SetToolTip(Me.Table, "Main Table of DBSheet: on this table the created DBSheet definition will allow ed" &
        "iting data")
        '
        'clearAllFields
        '
        Me.clearAllFields.AllowDrop = True
        Me.clearAllFields.BackColor = System.Drawing.SystemColors.Control
        Me.clearAllFields.ForeColor = System.Drawing.SystemColors.ControlText
        Me.clearAllFields.Location = New System.Drawing.Point(346, 36)
        Me.clearAllFields.Name = "clearAllFields"
        Me.clearAllFields.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.clearAllFields.Size = New System.Drawing.Size(96, 24)
        Me.clearAllFields.TabIndex = 7
        Me.clearAllFields.Text = "&clear all Fields"
        Me.clearAllFields.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolTipMain.SetToolTip(Me.clearAllFields, "clear all column definitions in DBsheet")
        Me.clearAllFields.UseVisualStyleBackColor = False
        '
        'regenLookupQueries
        '
        Me.regenLookupQueries.AllowDrop = True
        Me.regenLookupQueries.BackColor = System.Drawing.SystemColors.Control
        Me.regenLookupQueries.ForeColor = System.Drawing.SystemColors.ControlText
        Me.regenLookupQueries.Location = New System.Drawing.Point(787, 37)
        Me.regenLookupQueries.Name = "regenLookupQueries"
        Me.regenLookupQueries.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.regenLookupQueries.Size = New System.Drawing.Size(163, 24)
        Me.regenLookupQueries.TabIndex = 10
        Me.regenLookupQueries.Text = "re&generate all lookup queries"
        Me.regenLookupQueries.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolTipMain.SetToolTip(Me.regenLookupQueries, "regenerate restrictions for foreign table queries")
        Me.regenLookupQueries.UseVisualStyleBackColor = False
        '
        'moveUp
        '
        Me.moveUp.AllowDrop = True
        Me.moveUp.BackColor = System.Drawing.SystemColors.Control
        Me.moveUp.Font = New System.Drawing.Font("Arial", 11.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.moveUp.ForeColor = System.Drawing.SystemColors.ControlText
        Me.moveUp.Location = New System.Drawing.Point(517, 37)
        Me.moveUp.Name = "moveUp"
        Me.moveUp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.moveUp.Size = New System.Drawing.Size(25, 24)
        Me.moveUp.TabIndex = 8
        Me.moveUp.Text = "^"
        Me.moveUp.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolTipMain.SetToolTip(Me.moveUp, "move column in order of appearance 1 up ")
        Me.moveUp.UseVisualStyleBackColor = False
        '
        'moveDown
        '
        Me.moveDown.AllowDrop = True
        Me.moveDown.BackColor = System.Drawing.SystemColors.Control
        Me.moveDown.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.moveDown.ForeColor = System.Drawing.SystemColors.ControlText
        Me.moveDown.Location = New System.Drawing.Point(548, 37)
        Me.moveDown.Name = "moveDown"
        Me.moveDown.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.moveDown.Size = New System.Drawing.Size(24, 24)
        Me.moveDown.TabIndex = 9
        Me.moveDown.Text = "v"
        Me.moveDown.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolTipMain.SetToolTip(Me.moveDown, "move column in order of appearance 1 down")
        Me.moveDown.UseVisualStyleBackColor = False
        '
        'addAllFields
        '
        Me.addAllFields.AllowDrop = True
        Me.addAllFields.BackColor = System.Drawing.SystemColors.Control
        Me.addAllFields.ForeColor = System.Drawing.SystemColors.ControlText
        Me.addAllFields.Location = New System.Drawing.Point(262, 36)
        Me.addAllFields.Name = "addAllFields"
        Me.addAllFields.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.addAllFields.Size = New System.Drawing.Size(78, 24)
        Me.addAllFields.TabIndex = 6
        Me.addAllFields.Text = "add all &Fields"
        Me.addAllFields.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolTipMain.SetToolTip(Me.addAllFields, "add all columns to DBsheet")
        Me.addAllFields.UseVisualStyleBackColor = False
        '
        'testLookupQuery
        '
        Me.testLookupQuery.AllowDrop = True
        Me.testLookupQuery.BackColor = System.Drawing.SystemColors.Control
        Me.testLookupQuery.CausesValidation = False
        Me.testLookupQuery.ForeColor = System.Drawing.SystemColors.ControlText
        Me.testLookupQuery.Location = New System.Drawing.Point(985, 35)
        Me.testLookupQuery.Name = "testLookupQuery"
        Me.testLookupQuery.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.testLookupQuery.Size = New System.Drawing.Size(116, 25)
        Me.testLookupQuery.TabIndex = 11
        Me.testLookupQuery.Text = "test &Lookup Query"
        Me.testLookupQuery.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolTipMain.SetToolTip(Me.testLookupQuery, "test Lookup query in a new Excel Sheet...")
        Me.testLookupQuery.UseVisualStyleBackColor = False
        '
        'saveDefs
        '
        Me.saveDefs.AllowDrop = True
        Me.saveDefs.BackColor = System.Drawing.SystemColors.Control
        Me.saveDefs.ForeColor = System.Drawing.SystemColors.ControlText
        Me.saveDefs.Location = New System.Drawing.Point(346, 6)
        Me.saveDefs.Name = "saveDefs"
        Me.saveDefs.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.saveDefs.Size = New System.Drawing.Size(109, 23)
        Me.saveDefs.TabIndex = 3
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
        Me.saveDefsAs.Location = New System.Drawing.Point(461, 6)
        Me.saveDefsAs.Name = "saveDefsAs"
        Me.saveDefsAs.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.saveDefsAs.Size = New System.Drawing.Size(130, 23)
        Me.saveDefsAs.TabIndex = 4
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
        Me.loadDefs.Location = New System.Drawing.Point(237, 6)
        Me.loadDefs.Name = "loadDefs"
        Me.loadDefs.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.loadDefs.Size = New System.Drawing.Size(103, 23)
        Me.loadDefs.TabIndex = 2
        Me.loadDefs.Text = "&load DBSheet def"
        Me.loadDefs.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolTipMain.SetToolTip(Me.loadDefs, "load DB Sheet definitions into current context")
        Me.loadDefs.UseVisualStyleBackColor = False
        '
        'Password
        '
        Me.Password.Location = New System.Drawing.Point(992, 9)
        Me.Password.Name = "Password"
        Me.Password.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.Password.Size = New System.Drawing.Size(109, 20)
        Me.Password.TabIndex = 0
        Me.ToolTipMain.SetToolTip(Me.Password, "enter the password for the user required to access schema information (given in D" &
        "BSheetConnString)")
        '
        'Database
        '
        Me.Database.AllowDrop = True
        Me.Database.BackColor = System.Drawing.SystemColors.Window
        Me.Database.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Database.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Database.Location = New System.Drawing.Point(787, 8)
        Me.Database.Name = "Database"
        Me.Database.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Database.Size = New System.Drawing.Size(163, 21)
        Me.Database.Sorted = True
        Me.Database.TabIndex = 1
        Me.ToolTipMain.SetToolTip(Me.Database, "choose Database to select foreign tables from.")
        '
        'Label3
        '
        Me.Label3.AllowDrop = True
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(848, 348)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(131, 13)
        Me.Label3.TabIndex = 100
        Me.Label3.Text = "Where Parameter Clause:"
        '
        'LTable
        '
        Me.LTable.AllowDrop = True
        Me.LTable.AutoSize = True
        Me.LTable.BackColor = System.Drawing.Color.Transparent
        Me.LTable.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LTable.Location = New System.Drawing.Point(12, 40)
        Me.LTable.Name = "LTable"
        Me.LTable.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LTable.Size = New System.Drawing.Size(33, 13)
        Me.LTable.TabIndex = 100
        Me.LTable.Text = "Table"
        '
        'DBSheetCols
        '
        Me.DBSheetCols.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DBSheetCols.Location = New System.Drawing.Point(5, 68)
        Me.DBSheetCols.Name = "DBSheetCols"
        Me.DBSheetCols.Size = New System.Drawing.Size(1096, 258)
        Me.DBSheetCols.TabIndex = 10
        '
        'LDatabase
        '
        Me.LDatabase.AllowDrop = True
        Me.LDatabase.AutoSize = True
        Me.LDatabase.BackColor = System.Drawing.Color.Transparent
        Me.LDatabase.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LDatabase.Location = New System.Drawing.Point(728, 11)
        Me.LDatabase.Name = "LDatabase"
        Me.LDatabase.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LDatabase.Size = New System.Drawing.Size(53, 13)
        Me.LDatabase.TabIndex = 100
        Me.LDatabase.Text = "Database"
        '
        'Label1
        '
        Me.Label1.AllowDrop = True
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(956, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(27, 13)
        Me.Label1.TabIndex = 100
        Me.Label1.Text = "Pwd"
        '
        'LQuery
        '
        Me.LQuery.AllowDrop = True
        Me.LQuery.AutoSize = True
        Me.LQuery.BackColor = System.Drawing.Color.Transparent
        Me.LQuery.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LQuery.Location = New System.Drawing.Point(12, 343)
        Me.LQuery.Name = "LQuery"
        Me.LQuery.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LQuery.Size = New System.Drawing.Size(85, 13)
        Me.LQuery.TabIndex = 100
        Me.LQuery.Text = "DBSheet Query:"
        '
        'DBSheetCreateForm
        '
        Me.AllowDrop = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(1113, 543)
        Me.Controls.Add(Me.LQuery)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.LDatabase)
        Me.Controls.Add(Me.Database)
        Me.Controls.Add(Me.Password)
        Me.Controls.Add(Me.DBSheetCols)
        Me.Controls.Add(Me.createQuery)
        Me.Controls.Add(Me.Query)
        Me.Controls.Add(Me.testQuery)
        Me.Controls.Add(Me.WhereClause)
        Me.Controls.Add(Me.Table)
        Me.Controls.Add(Me.clearAllFields)
        Me.Controls.Add(Me.regenLookupQueries)
        Me.Controls.Add(Me.moveUp)
        Me.Controls.Add(Me.moveDown)
        Me.Controls.Add(Me.addAllFields)
        Me.Controls.Add(Me.testLookupQuery)
        Me.Controls.Add(Me.saveDefs)
        Me.Controls.Add(Me.saveDefsAs)
        Me.Controls.Add(Me.loadDefs)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.LTable)
        Me.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DBSheetCreateForm"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "DB Sheet creation"
        CType(Me.DBSheetCols, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DBSheetCols As Windows.Forms.DataGridView
    Friend WithEvents Password As Windows.Forms.TextBox
    Public WithEvents Database As Windows.Forms.ComboBox
    Public WithEvents LDatabase As Windows.Forms.Label
    Public WithEvents Label1 As Windows.Forms.Label
    Public WithEvents LQuery As Windows.Forms.Label
#End Region
End Class