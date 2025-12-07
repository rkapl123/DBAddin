Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports Microsoft.Vbe.Interop
Imports System.Windows.Forms
Imports System.Collections.Generic


''' <summary>Dialog for creating DB Modifier configurations</summary>
Public Class DBModifCreate
    ''' <summary>on loading of Form catch Data Errors produced when filling DBSeqenceDataGrid here</summary>
    Private DBSeqStepValidationErrors As String = ""
    ''' <summary>only catch errors until Form is displayed</summary>
    Private DBSeqStepValidationErrorsShown As Boolean = False

    ''' <summary>check for required fields before closing</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Dim NameValidationResult As String = ""
        ' Check for valid range name
        If Me.DBModifName.Text <> "" Then
            Dim checkName As String = Me.Tag + Me.DBModifName.Text
            If checkName.Length() > 255 Then
                NameValidationResult = "more than 255 characters long (including " + Me.Tag + ") !"
            ElseIf IsNumeric(Strings.Left(checkName, 1)) Then
                NameValidationResult = "starts with a number !"
            Else
                For i As Integer = 0 To checkName.Length - 1
                    If Not Char.IsLetterOrDigit(checkName.Chars(i)) And checkName.Chars(i) <> "_" Then
                        NameValidationResult = "contains non-alphanumeric character: " + checkName.Chars(i).ToString() + " !"
                        Exit For
                    End If
                Next
            End If
        End If
        Dim primKeys As Integer = 0
        ' besides valid range name, also check for requirements: 
        ' mandatory fields filled (visible Table-name, Primary keys And Database), NameValidation above OK, no Double invocation for execOnSave in DB Sequences And sequence parts and only one primary key for AutoInc Flag
        ' Beware: All If/ElseIf branches have to contain a validation error message, because the dialog stays open in this case. Only the Else branch closes the dialog.
        If NameValidationResult <> "" Then
            UserMsg("Invalid DBModifier name '" + Me.DBModifName.Text + "', Error: " + NameValidationResult, Me.Tag + " Validation")
        ElseIf Me.Tablename.Text = "" And Me.Tablename.Visible Then
            UserMsg("Field Table-name is required, please fill in!", Me.Tag + " Validation")
        ElseIf Me.PrimaryKeys.Visible AndAlso Not Integer.TryParse(Me.PrimaryKeys.Text, primKeys) Then
            UserMsg("Field Primary Keys is required and has to be an integer number, please fill in accordingly!", Me.Tag + " Validation")
        ElseIf Me.Database.Text = "" And Me.Database.Visible Then
            UserMsg("Field Database is required, please fill in!", Me.Tag + " Validation")
        ElseIf Me.Tag = "DBMapper" AndAlso Me.AutoIncFlag.Checked AndAlso primKeys > 1 Then
            UserMsg("Only one primary key is allowed when Auto Incrementing is enabled!", Me.Tag + " Validation")
        Else
            ' check for double invocation because of execOnSave both being set on current DB Modifier ...
            If Me.execOnSave.Checked And DBModifDefColl.ContainsKey("DBSeqnce") Then
                Dim MyDBModifName As String = Me.Tag + Me.DBModifName.Text
                ' and on DB Sequence that contains the current DB Mapper or DB Action:
                If Me.Tag <> "DBSeqnce" Then
                    For Each DBModifierCheck As DBSeqnce In DBModifDefColl("DBSeqnce").Values
                        ' check for Sequences that have execOnSave set...
                        If DBModifierCheck.execOnSave Then
                            ' ...if they contain the current DBAction/DBMapper
                            For Each sequenceParam As String In DBModifierCheck.getSequenceSteps
                                Dim definition() As String = Split(sequenceParam, ":")
                                If MyDBModifName = definition(1) Then
                                    Dim DBModifTargetAddress As String = "(Target Address could not be found...)"
                                    If DBModifDefColl(definition(0)).ContainsKey(definition(1)) Then DBModifTargetAddress = DBModifDefColl(definition(0)).Item(definition(1)).getTargetRangeAddress()
                                    Dim foundDBModifName As String = IIf(DBModifierCheck.getName = "DBSeqnce", "Unnamed DBSequence", DBModifierCheck.getName)
                                    UserMsg(Me.Tag + Me.DBModifName.Text + " in " + DBModifTargetAddress + " will be executed twice on saving, because it is part of '" + foundDBModifName + "', which is also executed on saving." + vbCrLf + IIf(Me.execOnSave.Checked, "Disabling 'Execute on save' now, you ", "You") + " can re-enable after disabling it on '" + foundDBModifName + "'", Me.Tag + " Validation")
                                    Me.execOnSave.Checked = False
                                End If
                            Next
                        End If
                    Next
                    If Not Me.execOnSave.Checked Then Exit Sub
                Else ' or on any DB Modifier being contained in current DB Sequence:
                    For i As Integer = 0 To Me.DBSeqenceDataGrid.Rows().Count - 2
                        Dim definition() As String = Split(Me.DBSeqenceDataGrid.Rows(i).Cells(0).Value, ":")
                        If (definition(0) = "DBAction" Or definition(0) = "DBMapper") AndAlso DBModifDefColl(definition(0)).ContainsKey(definition(1)) AndAlso DBModifDefColl(definition(0)).Item(definition(1)).execOnSave Then
                            Dim foundDBModifName As String = IIf(definition(1) = "", "Unnamed " + definition(0), definition(1))
                            Dim DBModifTargetAddress As String = DBModifDefColl(definition(0)).Item(definition(1)).getTargetRangeAddress()
                            UserMsg(foundDBModifName + " in " + DBModifTargetAddress + " will be executed twice on saving, because it is part of this DBSequence, which is also executed on saving." + vbCrLf + IIf(Me.execOnSave.Checked, "Disabling 'Execute on save' now, you ", "You") + " can re-enable after disabling it on '" + foundDBModifName + "'", Me.Tag + " Validation")
                            Me.execOnSave.Checked = False
                        End If
                    Next
                    If Not Me.execOnSave.Checked Then Exit Sub
                End If
            End If
            ' checks for DBSequences if refresh is done inside of transaction brackets (would lead to deadlocks)
            If Me.Tag = "DBSeqnce" Then
                Dim TransactionOpened As Boolean = False
                For i As Integer = 0 To Me.DBSeqenceDataGrid.Rows().Count - 2
                    Dim definition() As String = Split(Me.DBSeqenceDataGrid.Rows(i).Cells(0).Value, ":")
                    If (definition(0) = "DBBegin") Then
                        TransactionOpened = True
                    ElseIf (definition(0) = "DBCommitRollback") Then
                        TransactionOpened = False
                    End If
                    If (TransactionOpened And Strings.Left(definition(0), 7) = "Refresh") Then
                        UserMsg("You placed a " + definition(0) + " inside of a transaction, this might lead to deadlocks as DB functions use a different connection than DB Modifiers." + vbCrLf + "If the DB function done in the refresh doesn't query any data being modified inside the transaction, you may ignore this warning.", "DBSeqnce Validation", MsgBoxStyle.Exclamation)
                    End If
                Next
            End If
            ' validate parameter ranges for DBAction
            If Me.Tag = "DBAction" And Me.paramRangesStr.Text <> "" Then
                For Each paramRange In Split(Me.paramRangesStr.Text, ",")
                    Try
                        checkAndReturnRange(paramRange)
                    Catch ex As Exception
                        UserMsg(ex.Message, "DBAction Validation")
                        Exit Sub
                    End Try
                Next
            End If
            ' check if a new created definition already exists with the same name
            Dim CustomXmlParts As Object = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
            If Me.DBModifName.Tag = "" Then
                ' either in the DBModifier definitions (CustomXML) ..
                If CustomXmlParts.Count > 0 AndAlso Not IsNothing(CustomXmlParts(1).SelectSingleNode("/ns0:root/ns0:" + Me.Tag + "[@Name='" + Me.DBModifName.Text + "']")) Then
                    UserMsg("There is already a " + Me.Tag + " definition named '" + Me.DBModifName.Text + "' in the DBModif Definitions of the current workbook", Me.Tag + "Validation")
                    Exit Sub
                End If
                ' or as a name (DBMapper/DBAction)
                Dim checkExists As Excel.Name = Nothing
                Dim actWbNames As Excel.Names = Nothing
                Try : actWbNames = ExcelDnaUtil.Application.ActiveWorkbook.Names : Catch ex As Exception : End Try
                Try : checkExists = actWbNames.Item(Me.Tag + Me.DBModifName.Text) : Catch ex As Exception : End Try
                If Not IsNothing(checkExists) And Me.Tag <> "DBSeqnce" Then
                    UserMsg("DBModifier range name '" + Me.Tag + Me.DBModifName.Text + "' already exists!", Me.Tag + "Validation")
                    Exit Sub
                End If
            End If
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End If
    End Sub

    ''' <summary>ignore all done changes in dialog</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    ''' <summary>in case of (actually impossible) data errors in DBSequence DataGridView row entries, catch and log them here</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DBSeqenceDataGrid_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DBSeqenceDataGrid.DataError
        If Not DBSeqStepValidationErrorsShown Then
            DBSeqStepValidationErrors += "Error in row " + (e.RowIndex + 1).ToString() + ",content: " + Me.DBSeqenceDataGrid.Rows(e.RowIndex).Cells(0).Value + vbCrLf
        End If
    End Sub

    ''' <summary>the DBMapper and DBAction Target Range Address is displayed as a hyperlink, simulate this link here</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TargetRangeAddress_Click(sender As Object, e As EventArgs) Handles TargetRangeAddress.Click
        If Me.TargetRangeAddress.Text = "" Then Exit Sub
        ' only get TargetRangeAddress up to bracket (possibly contained named range formula)
        Dim clickAddress = Me.TargetRangeAddress.Text.Substring(0, IIf(Me.TargetRangeAddress.Text.IndexOf("(") < 0, Me.TargetRangeAddress.Text.Length, Me.TargetRangeAddress.Text.IndexOf("(")))
        Dim rangePart() As String = Split(clickAddress, "!")
        Try
            ExcelDnaUtil.Application.Worksheets(rangePart(0)).Select()
            ExcelDnaUtil.Application.Range(rangePart(1)).Select()
        Catch ex As Exception
            UserMsg("Couldn't select " + clickAddress + ":" + ex.Message)
        End Try
    End Sub

    ''' <summary>move row up in DataGridView of DB Sequence</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub MoveRowUp_Click(sender As Object, e As EventArgs) Handles MoveRowUp.Click
        If IsNothing(DBSeqenceDataGrid.CurrentRow) Then Return
        Dim rw As DataGridViewRow = DBSeqenceDataGrid.CurrentRow
        Dim selIndex As Integer = DBSeqenceDataGrid.CurrentRow.Index
        ' avoid moving up of first row
        If selIndex = 0 Then Return
        DBSeqenceDataGrid.Rows.RemoveAt(selIndex)
        DBSeqenceDataGrid.Rows.Insert(selIndex - 1, rw)
        DBSeqenceDataGrid.Rows(selIndex - 1).Cells(0).Selected = True
    End Sub

    ''' <summary>move row down in DataGridView of DB Sequence</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub MoveRowDown_Click(sender As Object, e As EventArgs) Handles MoveRowDown.Click
        If IsNothing(DBSeqenceDataGrid.CurrentRow) Then Return
        Dim rw As DataGridViewRow = DBSeqenceDataGrid.CurrentRow
        Dim selIndex As Integer = DBSeqenceDataGrid.CurrentRow.Index
        ' avoid moving down of last row, DBSeqenceDataGrid.Rows.Count is 1 more than the actual inserted rows because of the "new" row, selIndex is 0 based
        If selIndex = DBSeqenceDataGrid.Rows.Count - 2 Then Return
        DBSeqenceDataGrid.Rows.RemoveAt(selIndex)
        DBSeqenceDataGrid.Rows.Insert(selIndex + 1, rw)
        DBSeqenceDataGrid.Rows(selIndex + 1).Cells(0).Selected = True
    End Sub

    ''' <summary>prepare context menus to be displayed after right mouse click</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DBSeqenceDataGrid_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DBSeqenceDataGrid.CellMouseDown
        If e.Button = System.Windows.Forms.MouseButtons.Right Then
            DBSeqenceDataGrid.ContextMenuStrip = MoveMenu
        End If
    End Sub

    ''' <summary>Shown Event to display Data Errors when adding DBSequence Grid elements</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DBModifCreate_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        If Me.Tag = "DBAction" Then
            ' add all visible names to context menu on paramRangesStr text box, first workbook level names
            Dim submenu As New ToolStripMenuItem With {
                .Text = ExcelDnaUtil.Application.ActiveWorkbook.Name
            }
            For Each aName As Excel.Name In ExcelDnaUtil.Application.ActiveWorkbook.Names()
                If aName.Visible And InStr(aName.Name, "!") = 0 Then submenu.DropDownItems.Add(aName.Name, Nothing, AddressOf paramRangesContextMenuHandler)
            Next
            If submenu.DropDownItems.Count > 0 Then paramRangeMenu.Items.Add(submenu)
            ' then for each worksheet the worksheet level names
            For Each ws As Excel.Worksheet In ExcelDnaUtil.Application.ActiveWorkbook.Worksheets()
                submenu = New ToolStripMenuItem With {
                    .Text = ws.Name
                }
                For Each aName As Excel.Name In ws.Names()
                    If aName.Visible Then submenu.DropDownItems.Add(aName.Name, Nothing, AddressOf paramRangesContextMenuHandler)
                Next
                If submenu.DropDownItems.Count > 0 Then paramRangeMenu.Items.Add(submenu)
            Next
        End If

        ' on creating the form and DBSequenceDataGrid in DBModif.createDBModif, Data Errors produced when filling DBSeqenceDataGrid are caught by 
        ' DBModifCreate.DBSeqenceDataGrid_DataError event procedure and stored in DBSeqStepValidationErrors. 
        ' If any errors have been caught, display these in alternate text-form RepairDBSeqnce along with instructions on how to repair them
        If DBSeqStepValidationErrors <> "" Then
            ' first get allowed values from filled DataGridView DataSource
            Dim cb As DataGridViewComboBoxColumn = DBSeqenceDataGrid.Columns(0)
            Dim ds As List(Of String) = cb.DataSource()
            Dim allowedValues As String = ""
            For Each def As String In ds
                allowedValues += def + vbCrLf
            Next
            ' then display allowed values along with error messages and instruction on how to repair.
            Me.RepairDBSeqnce.Text = DBSeqStepValidationErrors + vbCrLf + "Allowed Entries are:" + vbCrLf + allowedValues + vbCrLf + "Repair existing definitions below and remove all above incl. this line to fix it by clicking OK:" + vbCrLf + Me.RepairDBSeqnce.Text
            Me.RepairDBSeqnce.Show()
            Me.RepairDBSeqnce.Width = Me.DBSeqenceDataGrid.Width
            Me.RepairDBSeqnce.Height = 325
            Me.RepairDBSeqnce.Top = Me.DatabaseLabel.Top
            Me.DBSeqenceDataGrid.Hide()
            ' go into "repaired" mode (indicating rewriting DBSequence Steps in DBModif.createDBModif)
            Me.Tag = "repaired"
            UserMsg("Defined DBSequence steps did not match allowed values." + vbCrLf + "Please follow the instructions in text-box to fix it...", "DBSequence definition Insert error")
        End If
        DBSeqStepValidationErrorsShown = True
    End Sub

    ''' <summary>Create Command-button Click event</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CreateCB_Click(sender As Object, e As EventArgs) Handles CreateCB.Click
        ' create a command-button for the current DBmodification?
        Dim cbshp As Excel.OLEObject = Nothing
        Dim cb As Forms.CommandButton
        Try
            cbshp = ExcelDnaUtil.Application.ActiveSheet.OLEObjects.Add(ClassType:="Forms.CommandButton.1", Link:=False, DisplayAsIcon:=False, Left:=600, Top:=70, Width:=120, Height:=24)
            cb = cbshp.Object
        Catch ex As Exception
            UserMsg("Can't create command button: " + ex.Message, "CommandButton create Error")
            Try : cbshp.Delete() : Catch ex2 As Exception : End Try
            Exit Sub
        End Try
        Dim cbName As String = Me.Tag + Me.DBModifName.Text
        Try
            cb.Name = cbName
            cb.Caption = IIf(Me.DBModifName.Text = "", "Unnamed " + Me.Tag, Me.Tag + Me.DBModifName.Text)
        Catch ex As Exception
            cbshp.Delete()
            ' known failure when setting the cb name if there already exists a button with that name
            If ex.Message.Contains("0x8002802C") Then
                UserMsg("Can't name the new command button '" + cbName + "' as there already exists a button with that name: " + ex.Message, "CommandButton create Error")
            Else
                UserMsg("Can't name command button '" + cbName + "': " + ex.Message, "CommandButton create Error")
            End If
            Exit Sub
        End Try
        If Len(cbName) > excelNameLengthLimit Then
            cbshp.Delete()
            UserMsg("CommandButton code-names cannot be longer than " + CStr(excelNameLengthLimit) + " characters ! '" + cbName + "': ", "CommandButton create Error")
            Exit Sub
        End If
        ' fail to assign a handler? remove command-button (otherwise it gets hard to edit an existing DBModification with a different name).
        Try
            AddInEvents.colCommandButtons.Add(New CommandbuttonClickHandler With {.cb = cb})
        Catch ex As Exception
            UserMsg("Error assigning DB modifier commandbutton '" + cbName + "': " + ex.Message, "CommandButton create Error")
            cbshp.Delete()
        End Try
    End Sub

    ''' <summary>trigger to enable/disable all parametrized settings</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub parametrized_Click(sender As Object, e As EventArgs) Handles parametrized.Click
        setDBActionParametrizedGUI()
    End Sub

    ''' <summary>setDBActionParametrizedGUI: actual enabling of parametrized settings (done separately to be able to set this on form startup)</summary>
    Public Sub setDBActionParametrizedGUI()
        If Me.parametrized.Checked Then
            Me.paramRangesStr.Enabled = True
            Me.convertAsDate.Enabled = True
            Me.convertAsString.Enabled = True
            Me.paramEnclosing.Enabled = True
            Me.continueIfRowEmpty.Enabled = True
        Else
            Me.paramRangesStr.Enabled = False
            Me.convertAsDate.Enabled = False
            Me.convertAsString.Enabled = False
            Me.paramEnclosing.Enabled = False
            Me.continueIfRowEmpty.Enabled = False
        End If

    End Sub
    ''' <summary>enables the context menu on paramRangesStr text box</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub paramRangesStr_MouseDown(sender As Object, e As MouseEventArgs) Handles paramRangesStr.MouseDown
        If e.Button = System.Windows.Forms.MouseButtons.Right Then
            paramRangesStr.ContextMenuStrip = paramRangeMenu
        End If
    End Sub

    ''' <summary>event handler for context menu on paramRangesStr text box</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Sub paramRangesContextMenuHandler(sender As Object, e As EventArgs)
        paramRangesStr.Text += IIf(Len(paramRangesStr.Text) > 0, ",", "") + sender.Text
        paramRangesStr.SelectionStart = Len(paramRangesStr.Text)
    End Sub

End Class