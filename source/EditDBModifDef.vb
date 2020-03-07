Imports System.Xml
Imports System.Windows.Forms
Imports ExcelDna.Integration
Imports Microsoft.Office.Core

''' <summary>Dialog used to display and edit the CustomXMLPart utilized for storing the DBModif definitions</summary>
Public Class EditDBModifDef
    Private CustomXmlParts As Object

    ''' <summary>store the displayed/edited textbox content back into the custom xml definition, indluding validation feedback</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub OKBtn_Click(sender As Object, e As EventArgs) Handles OKBtn.Click
        ' Make a StringWriter to reformat the indented XML.
        Using sw As New System.IO.StringWriter()
            ' Make a XmlTextWriter to (un)format the XML.
            Using xml_writer As New XmlTextWriter(sw)
                ' revert indentation...
                xml_writer.Formatting = Formatting.None
                Dim doc As New XmlDocument()
                Try
                    doc.LoadXml(Me.EditBox.Text)
                Catch ex As Exception
                    MsgBox("Problems with parsing changed definition: " & ex.Message, MsgBoxStyle.Critical)
                    Exit Sub
                End Try
                doc.WriteTo(xml_writer)
                xml_writer.Flush()
                ' store the result in CustomXmlParts
                CustomXmlParts(1).Delete
                CustomXmlParts.Add(sw.ToString())
            End Using
        End Using
        ' add/change the tickboxes doDBMOnSave and DBFskip
        If Not Me.DBFskip.CheckState = CheckState.Indeterminate Then
            Try
                Try : ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties("DBFskip").Delete : Catch ex As Exception : End Try
                ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties.Add(Name:="DBFskip", LinkToContent:=False, Type:=MsoDocProperties.msoPropertyTypeBoolean, Value:=Me.DBFskip.Checked)
            Catch ex As Exception
                MsgBox("Error when adding DBFskip to Workbook:" + ex.Message, MsgBoxStyle.Critical)
            End Try
        End If
        If Not Me.doDBMOnSave.CheckState = CheckState.Indeterminate Then
            Try
                Try : ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties("doDBMOnSave").Delete : Catch ex As Exception : End Try
                ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties.Add(Name:="doDBMOnSave", LinkToContent:=False, Type:=MsoDocProperties.msoPropertyTypeBoolean, Value:=Me.doDBMOnSave.Checked)
            Catch ex As Exception
                MsgBox("Error when adding doDBMOnSave to Workbook:" + ex.Message, MsgBoxStyle.Critical)
            End Try
        End If
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    ''' <summary>no change was made to definition</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CancelBtn_Click(sender As Object, e As EventArgs) Handles CancelBtn.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    ''' <summary>put the costom xml definition in the edit box for display/editing</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub EditDBModifDef_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        CustomXmlParts = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
        If CustomXmlParts.Count = 0 Then
            MsgBox("No DB Modifier Definition (CustomXMLPart) contained in Workbook " & ExcelDnaUtil.Application.ActiveWorkbook.Name)
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
                    Me.EditBox.Text = sw.ToString()
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
    End Sub

    ''' <summary>show the current line and column for easier detection of problems in xml document</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub EditBox_SelectionChanged(sender As Object, e As EventArgs) Handles EditBox.SelectionChanged
        Me.PosIndex.Text = "Line: " & (Me.EditBox.GetLineFromCharIndex(Me.EditBox.SelectionStart) + 1).ToString() & ", Column: " & (Me.EditBox.SelectionStart - Me.EditBox.GetFirstCharIndexOfCurrentLine + 1).ToString()
    End Sub

End Class