<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
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
	Public WithEvents ForTableKey As System.Windows.Forms.ComboBox
	Public WithEvents Table As System.Windows.Forms.ComboBox
	Public WithEvents Column As System.Windows.Forms.ComboBox
	Public WithEvents ForTable As System.Windows.Forms.ComboBox
	Public WithEvents ForTableLookup As System.Windows.Forms.ComboBox
	Public WithEvents IsForeign As System.Windows.Forms.CheckBox
	Public WithEvents addToDBsheetCols As System.Windows.Forms.Button
	Public WithEvents removeDBsheetCols As System.Windows.Forms.Button
	Public WithEvents clearAllFields As System.Windows.Forms.Button
	Public WithEvents outerJoin As System.Windows.Forms.CheckBox
	Public WithEvents LookupQuery As System.Windows.Forms.TextBox
	Public WithEvents regenLookupQueries As System.Windows.Forms.Button
	Public WithEvents IsPrimary As System.Windows.Forms.CheckBox
    Public WithEvents moveUp As System.Windows.Forms.Button
    Public WithEvents moveDown As System.Windows.Forms.Button
	Public WithEvents addAllFields As System.Windows.Forms.Button
	Public WithEvents testLookupQuery As System.Windows.Forms.Button
	Public WithEvents ForDatabase As System.Windows.Forms.ComboBox
	Public WithEvents Sorting As System.Windows.Forms.ComboBox
    Public WithEvents saveDefs As System.Windows.Forms.Button
    Public WithEvents saveDefsAs As System.Windows.Forms.Button
    Public WithEvents loadDefs As System.Windows.Forms.Button
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents LTable As System.Windows.Forms.Label
    Public WithEvents LColumn As System.Windows.Forms.Label
    Public WithEvents LForTableKey As System.Windows.Forms.Label
    Public WithEvents LForTableLookup As System.Windows.Forms.Label
    Public WithEvents LLookupQuery As System.Windows.Forms.Label
    Public WithEvents LForDatabase As System.Windows.Forms.Label
    Public WithEvents LForTable As System.Windows.Forms.Label
    Public WithEvents Label28 As System.Windows.Forms.Label

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
        Me.ForTableKey = New System.Windows.Forms.ComboBox()
        Me.Table = New System.Windows.Forms.ComboBox()
        Me.Column = New System.Windows.Forms.ComboBox()
        Me.ForTable = New System.Windows.Forms.ComboBox()
        Me.ForTableLookup = New System.Windows.Forms.ComboBox()
        Me.IsForeign = New System.Windows.Forms.CheckBox()
        Me.addToDBsheetCols = New System.Windows.Forms.Button()
        Me.removeDBsheetCols = New System.Windows.Forms.Button()
        Me.clearAllFields = New System.Windows.Forms.Button()
        Me.outerJoin = New System.Windows.Forms.CheckBox()
        Me.LookupQuery = New System.Windows.Forms.TextBox()
        Me.regenLookupQueries = New System.Windows.Forms.Button()
        Me.IsPrimary = New System.Windows.Forms.CheckBox()
        Me.moveUp = New System.Windows.Forms.Button()
        Me.moveDown = New System.Windows.Forms.Button()
        Me.addAllFields = New System.Windows.Forms.Button()
        Me.testLookupQuery = New System.Windows.Forms.Button()
        Me.ForDatabase = New System.Windows.Forms.ComboBox()
        Me.Sorting = New System.Windows.Forms.ComboBox()
        Me.saveDefs = New System.Windows.Forms.Button()
        Me.saveDefsAs = New System.Windows.Forms.Button()
        Me.loadDefs = New System.Windows.Forms.Button()
        Me.Password = New System.Windows.Forms.TextBox()
        Me.Database = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.LTable = New System.Windows.Forms.Label()
        Me.LColumn = New System.Windows.Forms.Label()
        Me.LForTableKey = New System.Windows.Forms.Label()
        Me.LForTableLookup = New System.Windows.Forms.Label()
        Me.LLookupQuery = New System.Windows.Forms.Label()
        Me.LForDatabase = New System.Windows.Forms.Label()
        Me.LForTable = New System.Windows.Forms.Label()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.DBSheetCols = New System.Windows.Forms.DataGridView()
        Me.LDatabase = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.DBSheetCols, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'createQuery
        '
        Me.createQuery.AllowDrop = True
        Me.createQuery.BackColor = System.Drawing.SystemColors.Control
        Me.createQuery.ForeColor = System.Drawing.SystemColors.ControlText
        Me.createQuery.Location = New System.Drawing.Point(754, 8)
        Me.createQuery.Name = "createQuery"
        Me.createQuery.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.createQuery.Size = New System.Drawing.Size(133, 24)
        Me.createQuery.TabIndex = 36
        Me.createQuery.TabStop = False
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
        Me.Query.Location = New System.Drawing.Point(756, 37)
        Me.Query.MaxLength = 0
        Me.Query.Multiline = True
        Me.Query.Name = "Query"
        Me.Query.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Query.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.Query.Size = New System.Drawing.Size(445, 451)
        Me.Query.TabIndex = 35
        Me.Query.Text = "__"
        Me.ToolTipMain.SetToolTip(Me.Query, "Query: If needed, modify the query for the DBSheet data being displayed. Attentio" &
        "n: create DBSheet query will destroy all custom information here !!")
        '
        'testQuery
        '
        Me.testQuery.AllowDrop = True
        Me.testQuery.BackColor = System.Drawing.SystemColors.Control
        Me.testQuery.ForeColor = System.Drawing.SystemColors.ControlText
        Me.testQuery.Location = New System.Drawing.Point(894, 8)
        Me.testQuery.Name = "testQuery"
        Me.testQuery.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.testQuery.Size = New System.Drawing.Size(121, 24)
        Me.testQuery.TabIndex = 34
        Me.testQuery.TabStop = False
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
        Me.WhereClause.Location = New System.Drawing.Point(756, 512)
        Me.WhereClause.MaxLength = 0
        Me.WhereClause.Multiline = True
        Me.WhereClause.Name = "WhereClause"
        Me.WhereClause.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.WhereClause.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.WhereClause.Size = New System.Drawing.Size(445, 64)
        Me.WhereClause.TabIndex = 33
        Me.ToolTipMain.SetToolTip(Me.WhereClause, "Where Parameter Clause: Restrict the data displayed with the Where part of an SQL" &
        " Select statement (enter without ""WHERE"" !).")
        '
        'ForTableKey
        '
        Me.ForTableKey.AllowDrop = True
        Me.ForTableKey.BackColor = System.Drawing.SystemColors.Window
        Me.ForTableKey.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ForTableKey.ForeColor = System.Drawing.SystemColors.WindowText
        Me.ForTableKey.Location = New System.Drawing.Point(352, 94)
        Me.ForTableKey.Name = "ForTableKey"
        Me.ForTableKey.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ForTableKey.Size = New System.Drawing.Size(189, 21)
        Me.ForTableKey.Sorted = True
        Me.ForTableKey.TabIndex = 24
        Me.ToolTipMain.SetToolTip(Me.ForTableKey, "id of the lookup information in the foreign table")
        '
        'Table
        '
        Me.Table.AllowDrop = True
        Me.Table.BackColor = System.Drawing.SystemColors.Window
        Me.Table.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Table.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Table.Location = New System.Drawing.Point(5, 52)
        Me.Table.Name = "Table"
        Me.Table.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Table.Size = New System.Drawing.Size(156, 21)
        Me.Table.Sorted = True
        Me.Table.TabIndex = 23
        Me.ToolTipMain.SetToolTip(Me.Table, "Main Table of DBSheet: on this table the created DBSheet definition will allow ed" &
        "iting data")
        '
        'Column
        '
        Me.Column.AllowDrop = True
        Me.Column.BackColor = System.Drawing.SystemColors.Window
        Me.Column.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Column.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Column.Location = New System.Drawing.Point(168, 52)
        Me.Column.Name = "Column"
        Me.Column.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Column.Size = New System.Drawing.Size(232, 21)
        Me.Column.Sorted = True
        Me.Column.TabIndex = 22
        Me.ToolTipMain.SetToolTip(Me.Column, "column to be added to the table (can be foreign id, see right)")
        '
        'ForTable
        '
        Me.ForTable.AllowDrop = True
        Me.ForTable.BackColor = System.Drawing.SystemColors.Window
        Me.ForTable.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ForTable.ForeColor = System.Drawing.SystemColors.WindowText
        Me.ForTable.Location = New System.Drawing.Point(168, 94)
        Me.ForTable.Name = "ForTable"
        Me.ForTable.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ForTable.Size = New System.Drawing.Size(178, 21)
        Me.ForTable.Sorted = True
        Me.ForTable.TabIndex = 21
        Me.ToolTipMain.SetToolTip(Me.ForTable, "foreign table of the lookup information")
        '
        'ForTableLookup
        '
        Me.ForTableLookup.AllowDrop = True
        Me.ForTableLookup.BackColor = System.Drawing.SystemColors.Window
        Me.ForTableLookup.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ForTableLookup.ForeColor = System.Drawing.SystemColors.WindowText
        Me.ForTableLookup.Location = New System.Drawing.Point(547, 94)
        Me.ForTableLookup.Name = "ForTableLookup"
        Me.ForTableLookup.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ForTableLookup.Size = New System.Drawing.Size(197, 21)
        Me.ForTableLookup.Sorted = True
        Me.ForTableLookup.TabIndex = 20
        Me.ToolTipMain.SetToolTip(Me.ForTableLookup, "name of column in foreign table that gives the user-readable information of the l" &
        "ookup")
        '
        'IsForeign
        '
        Me.IsForeign.AllowDrop = True
        Me.IsForeign.BackColor = System.Drawing.SystemColors.Control
        Me.IsForeign.ForeColor = System.Drawing.SystemColors.ControlText
        Me.IsForeign.Location = New System.Drawing.Point(514, 54)
        Me.IsForeign.Name = "IsForeign"
        Me.IsForeign.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.IsForeign.Size = New System.Drawing.Size(92, 16)
        Me.IsForeign.TabIndex = 19
        Me.IsForeign.Text = " is foreign?"
        Me.ToolTipMain.SetToolTip(Me.IsForeign, "is the column a foreign id needing a lookup, then define the foreign lookup infor" &
        "mation in the 4 fields below")
        Me.IsForeign.UseVisualStyleBackColor = False
        '
        'addToDBsheetCols
        '
        Me.addToDBsheetCols.AllowDrop = True
        Me.addToDBsheetCols.BackColor = System.Drawing.SystemColors.Control
        Me.addToDBsheetCols.ForeColor = System.Drawing.SystemColors.ControlText
        Me.addToDBsheetCols.Location = New System.Drawing.Point(5, 218)
        Me.addToDBsheetCols.Name = "addToDBsheetCols"
        Me.addToDBsheetCols.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.addToDBsheetCols.Size = New System.Drawing.Size(103, 24)
        Me.addToDBsheetCols.TabIndex = 18
        Me.addToDBsheetCols.TabStop = False
        Me.addToDBsheetCols.Text = "&add to DBsheet"
        Me.addToDBsheetCols.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolTipMain.SetToolTip(Me.addToDBsheetCols, "add above column definitions to DBsheet (alternate use: abort editing of current " &
        "field)")
        Me.addToDBsheetCols.UseVisualStyleBackColor = False
        '
        'removeDBsheetCols
        '
        Me.removeDBsheetCols.AllowDrop = True
        Me.removeDBsheetCols.BackColor = System.Drawing.SystemColors.Control
        Me.removeDBsheetCols.ForeColor = System.Drawing.SystemColors.ControlText
        Me.removeDBsheetCols.Location = New System.Drawing.Point(110, 218)
        Me.removeDBsheetCols.Name = "removeDBsheetCols"
        Me.removeDBsheetCols.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.removeDBsheetCols.Size = New System.Drawing.Size(118, 24)
        Me.removeDBsheetCols.TabIndex = 17
        Me.removeDBsheetCols.TabStop = False
        Me.removeDBsheetCols.Text = "&remove from DBsheet"
        Me.removeDBsheetCols.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolTipMain.SetToolTip(Me.removeDBsheetCols, "delete selected (or last) column definition from DBSheet")
        Me.removeDBsheetCols.UseVisualStyleBackColor = False
        '
        'clearAllFields
        '
        Me.clearAllFields.AllowDrop = True
        Me.clearAllFields.BackColor = System.Drawing.SystemColors.Control
        Me.clearAllFields.ForeColor = System.Drawing.SystemColors.ControlText
        Me.clearAllFields.Location = New System.Drawing.Point(648, 218)
        Me.clearAllFields.Name = "clearAllFields"
        Me.clearAllFields.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.clearAllFields.Size = New System.Drawing.Size(96, 24)
        Me.clearAllFields.TabIndex = 16
        Me.clearAllFields.TabStop = False
        Me.clearAllFields.Text = "&clear all Fields"
        Me.clearAllFields.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolTipMain.SetToolTip(Me.clearAllFields, "clear all column definitions in DBsheet")
        Me.clearAllFields.UseVisualStyleBackColor = False
        '
        'outerJoin
        '
        Me.outerJoin.AllowDrop = True
        Me.outerJoin.BackColor = System.Drawing.SystemColors.Control
        Me.outerJoin.ForeColor = System.Drawing.SystemColors.ControlText
        Me.outerJoin.Location = New System.Drawing.Point(352, 126)
        Me.outerJoin.Name = "outerJoin"
        Me.outerJoin.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.outerJoin.Size = New System.Drawing.Size(89, 15)
        Me.outerJoin.TabIndex = 15
        Me.outerJoin.Text = " is outer join?"
        Me.ToolTipMain.SetToolTip(Me.outerJoin, "create outer join for selected foreign column lookup (allows null and nonexisting" &
        " values in foreign id)")
        Me.outerJoin.UseVisualStyleBackColor = False
        '
        'LookupQuery
        '
        Me.LookupQuery.AcceptsReturn = True
        Me.LookupQuery.AllowDrop = True
        Me.LookupQuery.BackColor = System.Drawing.SystemColors.Window
        Me.LookupQuery.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.LookupQuery.ForeColor = System.Drawing.SystemColors.WindowText
        Me.LookupQuery.Location = New System.Drawing.Point(5, 147)
        Me.LookupQuery.MaxLength = 0
        Me.LookupQuery.Multiline = True
        Me.LookupQuery.Name = "LookupQuery"
        Me.LookupQuery.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LookupQuery.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.LookupQuery.Size = New System.Drawing.Size(739, 64)
        Me.LookupQuery.TabIndex = 14
        Me.ToolTipMain.SetToolTip(Me.LookupQuery, "lookup query or values (separated by ""||"") here that restrict the possible values" &
        " of selected column")
        '
        'regenLookupQueries
        '
        Me.regenLookupQueries.AllowDrop = True
        Me.regenLookupQueries.BackColor = System.Drawing.SystemColors.Control
        Me.regenLookupQueries.ForeColor = System.Drawing.SystemColors.ControlText
        Me.regenLookupQueries.Location = New System.Drawing.Point(395, 218)
        Me.regenLookupQueries.Name = "regenLookupQueries"
        Me.regenLookupQueries.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.regenLookupQueries.Size = New System.Drawing.Size(163, 24)
        Me.regenLookupQueries.TabIndex = 13
        Me.regenLookupQueries.TabStop = False
        Me.regenLookupQueries.Text = "re&generate all lookup queries"
        Me.regenLookupQueries.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolTipMain.SetToolTip(Me.regenLookupQueries, "regenerate restrictions for foreign table queries")
        Me.regenLookupQueries.UseVisualStyleBackColor = False
        '
        'IsPrimary
        '
        Me.IsPrimary.AllowDrop = True
        Me.IsPrimary.BackColor = System.Drawing.SystemColors.Control
        Me.IsPrimary.ForeColor = System.Drawing.SystemColors.ControlText
        Me.IsPrimary.Location = New System.Drawing.Point(414, 53)
        Me.IsPrimary.Name = "IsPrimary"
        Me.IsPrimary.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.IsPrimary.Size = New System.Drawing.Size(76, 17)
        Me.IsPrimary.TabIndex = 12
        Me.IsPrimary.Text = " is primary key?"
        Me.ToolTipMain.SetToolTip(Me.IsPrimary, "define column as primary key (unique selection must be possible with combination " &
        "of all primary key columns)")
        Me.IsPrimary.UseVisualStyleBackColor = False
        '
        'moveUp
        '
        Me.moveUp.AllowDrop = True
        Me.moveUp.BackColor = System.Drawing.SystemColors.Control
        Me.moveUp.Font = New System.Drawing.Font("Arial", 11.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.moveUp.ForeColor = System.Drawing.SystemColors.ControlText
        Me.moveUp.Location = New System.Drawing.Point(288, 218)
        Me.moveUp.Name = "moveUp"
        Me.moveUp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.moveUp.Size = New System.Drawing.Size(25, 24)
        Me.moveUp.TabIndex = 10
        Me.moveUp.TabStop = False
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
        Me.moveDown.Location = New System.Drawing.Point(319, 218)
        Me.moveDown.Name = "moveDown"
        Me.moveDown.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.moveDown.Size = New System.Drawing.Size(24, 24)
        Me.moveDown.TabIndex = 9
        Me.moveDown.TabStop = False
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
        Me.addAllFields.Location = New System.Drawing.Point(564, 218)
        Me.addAllFields.Name = "addAllFields"
        Me.addAllFields.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.addAllFields.Size = New System.Drawing.Size(78, 24)
        Me.addAllFields.TabIndex = 8
        Me.addAllFields.TabStop = False
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
        Me.testLookupQuery.Location = New System.Drawing.Point(588, 120)
        Me.testLookupQuery.Name = "testLookupQuery"
        Me.testLookupQuery.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.testLookupQuery.Size = New System.Drawing.Size(156, 25)
        Me.testLookupQuery.TabIndex = 7
        Me.testLookupQuery.TabStop = False
        Me.testLookupQuery.Text = "test &Lookup Query"
        Me.testLookupQuery.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolTipMain.SetToolTip(Me.testLookupQuery, "test Lookup query in a new Excel Sheet...")
        Me.testLookupQuery.UseVisualStyleBackColor = False
        '
        'ForDatabase
        '
        Me.ForDatabase.AllowDrop = True
        Me.ForDatabase.BackColor = System.Drawing.SystemColors.Window
        Me.ForDatabase.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ForDatabase.ForeColor = System.Drawing.SystemColors.WindowText
        Me.ForDatabase.Location = New System.Drawing.Point(5, 94)
        Me.ForDatabase.Name = "ForDatabase"
        Me.ForDatabase.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ForDatabase.Size = New System.Drawing.Size(156, 21)
        Me.ForDatabase.Sorted = True
        Me.ForDatabase.TabIndex = 6
        Me.ToolTipMain.SetToolTip(Me.ForDatabase, "choose Database to select foreign tables from.")
        '
        'Sorting
        '
        Me.Sorting.AllowDrop = True
        Me.Sorting.BackColor = System.Drawing.SystemColors.Window
        Me.Sorting.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Sorting.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Sorting.Location = New System.Drawing.Point(667, 52)
        Me.Sorting.Name = "Sorting"
        Me.Sorting.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Sorting.Size = New System.Drawing.Size(77, 21)
        Me.Sorting.TabIndex = 5
        Me.ToolTipMain.SetToolTip(Me.Sorting, "shall DBsheet data be sorted by this column?")
        '
        'saveDefs
        '
        Me.saveDefs.AllowDrop = True
        Me.saveDefs.BackColor = System.Drawing.SystemColors.Control
        Me.saveDefs.ForeColor = System.Drawing.SystemColors.ControlText
        Me.saveDefs.Location = New System.Drawing.Point(244, 8)
        Me.saveDefs.Name = "saveDefs"
        Me.saveDefs.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.saveDefs.Size = New System.Drawing.Size(109, 23)
        Me.saveDefs.TabIndex = 2
        Me.saveDefs.TabStop = False
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
        Me.saveDefsAs.Location = New System.Drawing.Point(111, 8)
        Me.saveDefsAs.Name = "saveDefsAs"
        Me.saveDefsAs.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.saveDefsAs.Size = New System.Drawing.Size(130, 23)
        Me.saveDefsAs.TabIndex = 1
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
        Me.loadDefs.Size = New System.Drawing.Size(103, 23)
        Me.loadDefs.TabIndex = 0
        Me.loadDefs.TabStop = False
        Me.loadDefs.Text = "&load DBSheet def"
        Me.loadDefs.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolTipMain.SetToolTip(Me.loadDefs, "load DB Sheet definitions into current context")
        Me.loadDefs.UseVisualStyleBackColor = False
        '
        'Password
        '
        Me.Password.Location = New System.Drawing.Point(644, 11)
        Me.Password.Name = "Password"
        Me.Password.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.Password.Size = New System.Drawing.Size(100, 20)
        Me.Password.TabIndex = 39
        Me.ToolTipMain.SetToolTip(Me.Password, "enter the password for the user required to access schema information (given in D" &
        "BSheetConnString)")
        '
        'Database
        '
        Me.Database.AllowDrop = True
        Me.Database.BackColor = System.Drawing.SystemColors.Window
        Me.Database.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Database.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Database.Location = New System.Drawing.Point(451, 10)
        Me.Database.Name = "Database"
        Me.Database.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Database.Size = New System.Drawing.Size(146, 21)
        Me.Database.Sorted = True
        Me.Database.TabIndex = 40
        Me.ToolTipMain.SetToolTip(Me.Database, "choose Database to select foreign tables from.")
        '
        'Label3
        '
        Me.Label3.AllowDrop = True
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(753, 493)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(131, 13)
        Me.Label3.TabIndex = 37
        Me.Label3.Text = "Where Parameter Clause:"
        '
        'LTable
        '
        Me.LTable.AllowDrop = True
        Me.LTable.AutoSize = True
        Me.LTable.BackColor = System.Drawing.Color.Transparent
        Me.LTable.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LTable.Location = New System.Drawing.Point(2, 36)
        Me.LTable.Name = "LTable"
        Me.LTable.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LTable.Size = New System.Drawing.Size(33, 13)
        Me.LTable.TabIndex = 32
        Me.LTable.Text = "Table"
        '
        'LColumn
        '
        Me.LColumn.AllowDrop = True
        Me.LColumn.AutoSize = True
        Me.LColumn.BackColor = System.Drawing.Color.Transparent
        Me.LColumn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LColumn.Location = New System.Drawing.Point(165, 36)
        Me.LColumn.Name = "LColumn"
        Me.LColumn.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LColumn.Size = New System.Drawing.Size(29, 13)
        Me.LColumn.TabIndex = 31
        Me.LColumn.Text = "Field"
        '
        'LForTableKey
        '
        Me.LForTableKey.AllowDrop = True
        Me.LForTableKey.AutoSize = True
        Me.LForTableKey.BackColor = System.Drawing.Color.Transparent
        Me.LForTableKey.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LForTableKey.Location = New System.Drawing.Point(349, 78)
        Me.LForTableKey.Name = "LForTableKey"
        Me.LForTableKey.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LForTableKey.Size = New System.Drawing.Size(118, 13)
        Me.LForTableKey.TabIndex = 30
        Me.LForTableKey.Text = "Foreign Table Key Field"
        '
        'LForTableLookup
        '
        Me.LForTableLookup.AllowDrop = True
        Me.LForTableLookup.AutoSize = True
        Me.LForTableLookup.BackColor = System.Drawing.Color.Transparent
        Me.LForTableLookup.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LForTableLookup.Location = New System.Drawing.Point(544, 78)
        Me.LForTableLookup.Name = "LForTableLookup"
        Me.LForTableLookup.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LForTableLookup.Size = New System.Drawing.Size(134, 13)
        Me.LForTableLookup.TabIndex = 29
        Me.LForTableLookup.Text = "Foreign Table Lookup Field"
        '
        'LLookupQuery
        '
        Me.LLookupQuery.AllowDrop = True
        Me.LLookupQuery.AutoSize = True
        Me.LLookupQuery.BackColor = System.Drawing.Color.Transparent
        Me.LLookupQuery.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LLookupQuery.Location = New System.Drawing.Point(2, 126)
        Me.LLookupQuery.Name = "LLookupQuery"
        Me.LLookupQuery.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LLookupQuery.Size = New System.Drawing.Size(78, 13)
        Me.LLookupQuery.TabIndex = 28
        Me.LLookupQuery.Text = "Lookup Query:"
        '
        'LForDatabase
        '
        Me.LForDatabase.AllowDrop = True
        Me.LForDatabase.AutoSize = True
        Me.LForDatabase.BackColor = System.Drawing.Color.Transparent
        Me.LForDatabase.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LForDatabase.Location = New System.Drawing.Point(2, 78)
        Me.LForDatabase.Name = "LForDatabase"
        Me.LForDatabase.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LForDatabase.Size = New System.Drawing.Size(92, 13)
        Me.LForDatabase.TabIndex = 27
        Me.LForDatabase.Text = "Foreign Database"
        '
        'LForTable
        '
        Me.LForTable.AllowDrop = True
        Me.LForTable.AutoSize = True
        Me.LForTable.BackColor = System.Drawing.Color.Transparent
        Me.LForTable.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LForTable.Location = New System.Drawing.Point(165, 78)
        Me.LForTable.Name = "LForTable"
        Me.LForTable.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LForTable.Size = New System.Drawing.Size(72, 13)
        Me.LForTable.TabIndex = 26
        Me.LForTable.Text = "Foreign Table"
        '
        'Label28
        '
        Me.Label28.AllowDrop = True
        Me.Label28.AutoSize = True
        Me.Label28.BackColor = System.Drawing.Color.Transparent
        Me.Label28.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label28.Location = New System.Drawing.Point(630, 55)
        Me.Label28.Name = "Label28"
        Me.Label28.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label28.Size = New System.Drawing.Size(31, 13)
        Me.Label28.TabIndex = 25
        Me.Label28.Text = "Sort:"
        '
        'DBSheetCols
        '
        Me.DBSheetCols.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DBSheetCols.Location = New System.Drawing.Point(5, 249)
        Me.DBSheetCols.Name = "DBSheetCols"
        Me.DBSheetCols.Size = New System.Drawing.Size(739, 327)
        Me.DBSheetCols.TabIndex = 38
        '
        'LDatabase
        '
        Me.LDatabase.AllowDrop = True
        Me.LDatabase.AutoSize = True
        Me.LDatabase.BackColor = System.Drawing.Color.Transparent
        Me.LDatabase.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LDatabase.Location = New System.Drawing.Point(392, 13)
        Me.LDatabase.Name = "LDatabase"
        Me.LDatabase.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LDatabase.Size = New System.Drawing.Size(53, 13)
        Me.LDatabase.TabIndex = 41
        Me.LDatabase.Text = "Database"
        '
        'Label1
        '
        Me.Label1.AllowDrop = True
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(611, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(27, 13)
        Me.Label1.TabIndex = 42
        Me.Label1.Text = "Pwd"
        '
        'DBSheetCreateForm
        '
        Me.AllowDrop = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(1208, 588)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.LDatabase)
        Me.Controls.Add(Me.Database)
        Me.Controls.Add(Me.Password)
        Me.Controls.Add(Me.DBSheetCols)
        Me.Controls.Add(Me.createQuery)
        Me.Controls.Add(Me.Query)
        Me.Controls.Add(Me.testQuery)
        Me.Controls.Add(Me.WhereClause)
        Me.Controls.Add(Me.ForTableKey)
        Me.Controls.Add(Me.Table)
        Me.Controls.Add(Me.Column)
        Me.Controls.Add(Me.ForTable)
        Me.Controls.Add(Me.ForTableLookup)
        Me.Controls.Add(Me.IsForeign)
        Me.Controls.Add(Me.addToDBsheetCols)
        Me.Controls.Add(Me.removeDBsheetCols)
        Me.Controls.Add(Me.clearAllFields)
        Me.Controls.Add(Me.outerJoin)
        Me.Controls.Add(Me.LookupQuery)
        Me.Controls.Add(Me.regenLookupQueries)
        Me.Controls.Add(Me.IsPrimary)
        Me.Controls.Add(Me.moveUp)
        Me.Controls.Add(Me.moveDown)
        Me.Controls.Add(Me.addAllFields)
        Me.Controls.Add(Me.testLookupQuery)
        Me.Controls.Add(Me.ForDatabase)
        Me.Controls.Add(Me.Sorting)
        Me.Controls.Add(Me.saveDefs)
        Me.Controls.Add(Me.saveDefsAs)
        Me.Controls.Add(Me.loadDefs)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.LTable)
        Me.Controls.Add(Me.LColumn)
        Me.Controls.Add(Me.LForTableKey)
        Me.Controls.Add(Me.LForTableLookup)
        Me.Controls.Add(Me.LLookupQuery)
        Me.Controls.Add(Me.LForDatabase)
        Me.Controls.Add(Me.LForTable)
        Me.Controls.Add(Me.Label28)
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
#End Region
End Class