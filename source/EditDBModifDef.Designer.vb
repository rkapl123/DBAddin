﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EditDBModifDef))
        Me.OKBtn = New System.Windows.Forms.Button()
        Me.CancelBtn = New System.Windows.Forms.Button()
        Me.PosIndex = New System.Windows.Forms.Label()
        Me.EditBox = New System.Windows.Forms.RichTextBox()
        Me.doDBMOnSave = New System.Windows.Forms.CheckBox()
        Me.DBFskip = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'OKBtn
        '
        Me.OKBtn.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.OKBtn.Location = New System.Drawing.Point(664, 415)
        Me.OKBtn.Name = "OKBtn"
        Me.OKBtn.Size = New System.Drawing.Size(55, 23)
        Me.OKBtn.TabIndex = 1
        Me.OKBtn.Text = "OK"
        Me.OKBtn.UseVisualStyleBackColor = True
        '
        'CancelBtn
        '
        Me.CancelBtn.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CancelBtn.Location = New System.Drawing.Point(725, 415)
        Me.CancelBtn.Name = "CancelBtn"
        Me.CancelBtn.Size = New System.Drawing.Size(55, 23)
        Me.CancelBtn.TabIndex = 2
        Me.CancelBtn.Text = "Cancel"
        Me.CancelBtn.UseVisualStyleBackColor = True
        '
        'PosIndex
        '
        Me.PosIndex.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PosIndex.AutoSize = True
        Me.PosIndex.Location = New System.Drawing.Point(13, 415)
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
        Me.EditBox.Size = New System.Drawing.Size(768, 397)
        Me.EditBox.TabIndex = 4
        Me.EditBox.Text = ""
        '
        'doDBMOnSave
        '
        Me.doDBMOnSave.AutoSize = True
        Me.doDBMOnSave.Location = New System.Drawing.Point(519, 419)
        Me.doDBMOnSave.Name = "doDBMOnSave"
        Me.doDBMOnSave.Size = New System.Drawing.Size(139, 17)
        Me.doDBMOnSave.TabIndex = 5
        Me.doDBMOnSave.Text = "do DBModifiers on save"
        Me.doDBMOnSave.UseVisualStyleBackColor = True
        '
        'DBFskip
        '
        Me.DBFskip.AutoSize = True
        Me.DBFskip.Location = New System.Drawing.Point(356, 419)
        Me.DBFskip.Name = "DBFskip"
        Me.DBFskip.Size = New System.Drawing.Size(157, 17)
        Me.DBFskip.TabIndex = 6
        Me.DBFskip.Text = "skip DBFunctions on open?"
        Me.DBFskip.UseVisualStyleBackColor = True
        '
        'EditDBModifDef
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.DBFskip)
        Me.Controls.Add(Me.doDBMOnSave)
        Me.Controls.Add(Me.EditBox)
        Me.Controls.Add(Me.PosIndex)
        Me.Controls.Add(Me.CancelBtn)
        Me.Controls.Add(Me.OKBtn)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
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
End Class