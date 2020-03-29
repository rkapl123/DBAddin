Imports System.Xml
Imports System.Windows.Forms
Imports ExcelDna.Integration
Imports Microsoft.Office.Core
Imports System.IO
Imports System.Diagnostics
Imports System.Configuration

''' <summary>Dialog used to display and edit the CustomXMLPart utilized for storing the DBModif definitions</summary>
Public Class EditDBModifDef
    Private CustomXmlParts As Object

    ''' <summary>store the displayed/edited textbox content back into the custom xml definition, indluding validation feedback</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub OKBtn_Click(sender As Object, e As EventArgs) Handles OKBtn.Click
        If Me.DBFskip.Visible Then
            ' Make a StringWriter to reformat the indented XML.
            Using sw As New System.IO.StringWriter()
                ' Make a XmlTextWriter to (un)format the XML.
                Using xml_writer As New XmlTextWriter(sw)
                    ' revert indentation...
                    xml_writer.Formatting = Formatting.None
                    Dim doc As New XmlDocument()
                    Try
                        Dim schemaString As String = My.Resources.DefinitionFiles.DBModifDef
                        Dim schemadoc As XmlReader = XmlReader.Create(New StringReader(schemaString))
                        doc.Schemas.Add("DBModifDef", schemadoc)
                        Dim eventHandler As Schema.ValidationEventHandler = New Schema.ValidationEventHandler(AddressOf myValidationEventHandler)
                        doc.LoadXml(Me.EditBox.Text)
                        doc.Validate(eventHandler)
                    Catch ex As Exception
                        ErrorMsg("Problems with parsing changed definition: " & ex.Message, "Edit DB Modifier Definitions XML")
                        Exit Sub
                    End Try
                    doc.WriteTo(xml_writer)
                    xml_writer.Flush()
                    ' store the result in CustomXmlParts
                    CustomXmlParts(1).Delete
                    CustomXmlParts.Add(sw.ToString)
                End Using
            End Using
            ' add/change the tickboxes doDBMOnSave and DBFskip
            If Not Me.DBFskip.CheckState = CheckState.Indeterminate Then
                Try
                    Try : ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties("DBFskip").Delete : Catch ex As Exception : End Try
                    ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties.Add(Name:="DBFskip", LinkToContent:=False, Type:=MsoDocProperties.msoPropertyTypeBoolean, Value:=Me.DBFskip.Checked)
                Catch ex As Exception
                    ErrorMsg("Error when adding DBFskip to Workbook:" + ex.Message, "Edit DB Modifier Definitions XML")
                    Exit Sub
                End Try
            End If
            If Not Me.doDBMOnSave.CheckState = CheckState.Indeterminate Then
                Try
                    Try : ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties("doDBMOnSave").Delete : Catch ex As Exception : End Try
                    ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties.Add(Name:="doDBMOnSave", LinkToContent:=False, Type:=MsoDocProperties.msoPropertyTypeBoolean, Value:=Me.doDBMOnSave.Checked)
                Catch ex As Exception
                    ErrorMsg("Error when adding doDBMOnSave to Workbook:" + ex.Message, "Edit DB Modifier Definitions XML")
                    Exit Sub
                End Try
            End If
        Else
            ' save the users app settings...
            Dim doc As New XmlDocument()
            Try
                doc.LoadXml(Me.EditBox.Text)
            Catch ex As Exception
                ErrorMsg("Problems with parsing changed app settings: " & ex.Message, "Edit DB Addin Settings")
                Exit Sub
            End Try
            Try
                File.WriteAllText(xllPath & ".config", Me.EditBox.Text, System.Text.Encoding.Default)
            Catch ex As Exception
                ErrorMsg("Couldn't write DB Addin Usersettings from " & xllPath & ":" & ex.Message, "Edit DB Addin Settings")
                Exit Sub
            End Try
            Try
                ConfigurationManager.RefreshSection("appSettings")
                Dim testSetting As String = ConfigurationManager.AppSettings("DefaultEnvironment")
            Catch ex As Exception
                ErrorMsg("Problem with AppSettings:" & ex.Message, "Edit DB Addin Settings")
                Exit Sub
            End Try
        End If
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Sub myValidationEventHandler(sender As Object, e As Schema.ValidationEventArgs)
        If e.Severity = Schema.XmlSeverityType.Error Or e.Severity = Schema.XmlSeverityType.Warning Then Throw New Exception(e.Message)
    End Sub

    ''' <summary>no change was made to definition</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CancelBtn_Click(sender As Object, e As EventArgs) Handles CancelBtn.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private xllPath As String = ""

    ''' <summary>put the costom xml definition in the edit box for display/editing</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub EditDBModifDef_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        If Me.DBFskip.Visible Then
            CustomXmlParts = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
            If CustomXmlParts.Count = 0 Then
                ErrorMsg("No DB Modifier Definition (CustomXMLPart) contained in Workbook " & ExcelDnaUtil.Application.ActiveWorkbook.Name, "Edit DB Modifier Definitions XML")
                Me.DialogResult = DialogResult.Cancel
                Me.Close()
            Else
                ' Make a StringWriter to hold the result.
                Using sw As New System.IO.StringWriter()
                    ' Make a XmlTextWriter to format the XML.
                    Using xml_writer As New XmlTextWriter(sw)
                        xml_writer.Formatting = Formatting.Indented
                        Dim doc As New XmlDocument()
                        doc.LoadXml(CustomXmlParts(1).XML)
                        doc.WriteTo(xml_writer)
                        xml_writer.Flush()
                        ' Display the result.
                        Me.EditBox.Text = sw.ToString
                    End Using
                End Using
            End If
            Try
                Me.DBFskip.Checked = ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties("DBFskip").Value
            Catch ex As Exception
                Me.DBFskip.CheckState = CheckState.Indeterminate
            End Try
            Try
                Me.doDBMOnSave.Checked = ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties("doDBMOnSave").Value
            Catch ex As Exception
                Me.doDBMOnSave.CheckState = CheckState.Indeterminate
            End Try
        Else
            Me.OKBtn.Text = "Save"
            Me.ToolTip1.SetToolTip(OKBtn, "save DB Addin User settings to DBAddin.xll.config")
            ' gets user settings and display them
            ' get path of xll:
            For Each tModule As ProcessModule In Process.GetCurrentProcess().Modules
                Dim sModule As String = tModule.FileName
                If sModule.ToUpper.Contains("DBADDIN") Then
                    xllPath = tModule.FileName
                    Exit For
                End If
            Next
            ' read setting from xllPath + ".config"
            Dim appSettings As String
            Try
                appSettings = File.ReadAllText(xllPath & ".config", System.Text.Encoding.Default)
            Catch ex As Exception
                ErrorMsg("Couldn't read DB Addin Usersettings from " & xllPath & ":" & ex.Message, "Edit DB Addin Settings")
                Exit Sub
            End Try
            Me.EditBox.Text = appSettings
        End If
    End Sub

    ''' <summary>show the current line and column for easier detection of problems in xml document</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub EditBox_SelectionChanged(sender As Object, e As EventArgs) Handles EditBox.SelectionChanged
        Me.PosIndex.Text = "Line: " & (Me.EditBox.GetLineFromCharIndex(Me.EditBox.SelectionStart) + 1).ToString & ", Column: " & (Me.EditBox.SelectionStart - Me.EditBox.GetFirstCharIndexOfCurrentLine + 1).ToString
    End Sub

End Class