<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EditDBModifDef
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EditDBModifDef))
        Me.OKBtn = New System.Windows.Forms.Button()
        Me.CancelBtn = New System.Windows.Forms.Button()
        Me.PosIndex = New System.Windows.Forms.Label()
        Me.EditBox = New System.Windows.Forms.RichTextBox()
        Me.doDBMOnSave = New System.Windows.Forms.CheckBox()
        Me.DBFskip = New System.Windows.Forms.CheckBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.availSettingsLB = New System.Windows.Forms.ComboBox()
        Me.availSettLbl = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'OKBtn
        '
        Me.OKBtn.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.OKBtn.Location = New System.Drawing.Point(536, 381)
        Me.OKBtn.Name = "OKBtn"
        Me.OKBtn.Size = New System.Drawing.Size(55, 23)
        Me.OKBtn.TabIndex = 5
        Me.OKBtn.Text = "OK"
        Me.ToolTip1.SetToolTip(Me.OKBtn, "apply changes done in this dialog")
        Me.OKBtn.UseVisualStyleBackColor = True
        '
        'CancelBtn
        '
        Me.CancelBtn.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CancelBtn.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CancelBtn.Location = New System.Drawing.Point(597, 381)
        Me.CancelBtn.Name = "CancelBtn"
        Me.CancelBtn.Size = New System.Drawing.Size(55, 23)
        Me.CancelBtn.TabIndex = 0
        Me.CancelBtn.Text = "Cancel"
        Me.ToolTip1.SetToolTip(Me.CancelBtn, "discard changes done in this dialog")
        Me.CancelBtn.UseVisualStyleBackColor = True
        '
        'PosIndex
        '
        Me.PosIndex.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PosIndex.AutoSize = True
        Me.PosIndex.Location = New System.Drawing.Point(12, 381)
        Me.PosIndex.Name = "PosIndex"
        Me.PosIndex.Size = New System.Drawing.Size(0, 13)
        Me.PosIndex.TabIndex = 3
        '
        'EditBox
        '
        Me.EditBox.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.EditBox.Location = New System.Drawing.Point(12, 12)
        Me.EditBox.Name = "EditBox"
        Me.EditBox.Size = New System.Drawing.Size(640, 363)
        Me.EditBox.TabIndex = 1
        Me.EditBox.Text = ""
        Me.ToolTip1.SetToolTip(Me.EditBox, "Edit the DB Modifier Definition CustomXML here.")
        '
        'doDBMOnSave
        '
        Me.doDBMOnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.doDBMOnSave.AutoSize = True
        Me.doDBMOnSave.Location = New System.Drawing.Point(382, 385)
        Me.doDBMOnSave.Name = "doDBMOnSave"
        Me.doDBMOnSave.Size = New System.Drawing.Size(139, 17)
        Me.doDBMOnSave.TabIndex = 3
        Me.doDBMOnSave.Text = "do DBModifiers on save"
        Me.ToolTip1.SetToolTip(Me.doDBMOnSave, "sets CustomProperty doDBMOnSave in order to execute DB Modifiers marked for execu" &
        "tion automatically without asking.")
        Me.doDBMOnSave.UseVisualStyleBackColor = True
        '
        'DBFskip
        '
        Me.DBFskip.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DBFskip.AutoSize = True
        Me.DBFskip.Location = New System.Drawing.Point(219, 385)
        Me.DBFskip.Name = "DBFskip"
        Me.DBFskip.Size = New System.Drawing.Size(157, 17)
        Me.DBFskip.TabIndex = 2
        Me.DBFskip.Text = "skip DBFunctions on open?"
        Me.ToolTip1.SetToolTip(Me.DBFskip, "sets the CustomProperty skipDBFunc to true/false in order to not automatically re" &
        "fresh DB functions on Workbook opening.")
        Me.DBFskip.UseVisualStyleBackColor = True
        '
        'availSettingsLB
        '
        Me.availSettingsLB.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.availSettingsLB.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.availSettingsLB.FormattingEnabled = True
        Me.availSettingsLB.Location = New System.Drawing.Point(107, 408)
        Me.availSettingsLB.MaxDropDownItems = 30
        Me.availSettingsLB.Name = "availSettingsLB"
        Me.availSettingsLB.Size = New System.Drawing.Size(323, 21)
        Me.availSettingsLB.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.availSettingsLB, "Select available Setting to insert into above configuration file")
        '
        'availSettLbl
        '
        Me.availSettLbl.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.availSettLbl.AutoSize = True
        Me.availSettLbl.Location = New System.Drawing.Point(5, 411)
        Me.availSettLbl.Name = "availSettLbl"
        Me.availSettLbl.Size = New System.Drawing.Size(96, 13)
        Me.availSettLbl.TabIndex = 0
        Me.availSettLbl.Text = "available Setttings:"
        '
        'EditDBModifDef
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.CancelBtn
        Me.ClientSize = New System.Drawing.Size(664, 464)
        Me.ControlBox = False
        Me.Controls.Add(Me.availSettLbl)
        Me.Controls.Add(Me.DBFskip)
        Me.Controls.Add(Me.doDBMOnSave)
        Me.Controls.Add(Me.EditBox)
        Me.Controls.Add(Me.PosIndex)
        Me.Controls.Add(Me.CancelBtn)
        Me.Controls.Add(Me.OKBtn)
        Me.Controls.Add(Me.availSettingsLB)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(680, 480)
        Me.Name = "EditDBModifDef"
        Me.Text = "Edit DBModifier Definitions (CustomXML)"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OKBtn As Windows.Forms.Button
    Friend WithEvents CancelBtn As Windows.Forms.Button
    Friend WithEvents PosIndex As Windows.Forms.Label
    Friend WithEvents EditBox As Windows.Forms.RichTextBox
    Friend WithEvents doDBMOnSave As Windows.Forms.CheckBox
    Friend WithEvents DBFskip As Windows.Forms.CheckBox
    Friend WithEvents ToolTip1 As Windows.Forms.ToolTip
    Friend WithEvents availSettingsLB As Windows.Forms.ComboBox
    Friend WithEvents availSettLbl As Windows.Forms.Label
End Class
