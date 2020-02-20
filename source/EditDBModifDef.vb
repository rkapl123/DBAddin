Imports System.Xml
Imports System.Windows.Forms
Imports ExcelDna.Integration
Imports Microsoft.Office.Core

''' <summary>Dialog used to display and edit the CustomXMLPart utilized for storing the DBModif definitions</summary>
Public Class EditDBModifDef
    Private xmlPartEnumerator As System.Collections.IEnumerator
    Private xmlParts As CustomXMLPart

    Private Sub OKBtn_Click(sender As Object, e As EventArgs) Handles OKBtn.Click
        xmlParts.LoadXML(Me.EditBox.Text)
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub CancelBtn_Click(sender As Object, e As EventArgs) Handles CancelBtn.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub EditDBModifDef_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        xmlPartEnumerator = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef").GetEnumerator
        xmlPartEnumerator.Reset()
        If Not xmlPartEnumerator.MoveNext() Then
            MsgBox("No DBModifDef CustomXMLPart contained in Workbook " & ExcelDnaUtil.Application.ActiveWorkbook.Name)
            Me.DialogResult = DialogResult.Cancel
            Me.Close()
        Else
            xmlParts = DirectCast(xmlPartEnumerator, CustomXMLPart).Current
            ' Use an XmlTextWriter to format.
            ' Make a StringWriter to hold the result.
            Using sw As New System.IO.StringWriter()
                ' Make the XmlTextWriter to format the XML.
                Using xml_writer As New XmlTextWriter(sw)
                    xml_writer.Formatting = Formatting.Indented
                    Dim doc As New XmlDocument()
                    doc.LoadXml(xmlParts.XML)
                    doc.WriteTo(xml_writer)
                    xml_writer.Flush()
                    ' Display the result.
                    Me.EditBox.Text = sw.ToString()
                End Using
            End Using
        End If
    End Sub

End Class