Imports System.Xml
Imports System.Windows.Forms
Imports ExcelDna.Integration

''' <summary>Dialog used to display and edit the CustomXMLPart utilized for storing the DBModif definitions</summary>
Public Class EditDBModifDef
    Private CustomXmlParts As Object

    Private Sub OKBtn_Click(sender As Object, e As EventArgs) Handles OKBtn.Click
        Using sw As New System.IO.StringWriter()
            ' Make the XmlTextWriter to format the XML.
            Using xml_writer As New XmlTextWriter(sw)
                ' revert indentation...
                xml_writer.Formatting = Formatting.None
                Dim doc As New XmlDocument()
                Try
                    doc.LoadXml(Me.EditBox.Text)
                Catch ex As Exception
                    MsgBox("Problems with parsing changed definition: " & ex.Message)
                    Exit Sub
                End Try
                doc.WriteTo(xml_writer)
                xml_writer.Flush()
                ' store the result.
                CustomXmlParts(1).Delete
                CustomXmlParts.Add(sw.ToString())
            End Using
        End Using
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub CancelBtn_Click(sender As Object, e As EventArgs) Handles CancelBtn.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub EditDBModifDef_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        CustomXmlParts = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
        If CustomXmlParts.Count = 0 Then
            MsgBox("No DBModifDef CustomXMLPart contained in Workbook " & ExcelDnaUtil.Application.ActiveWorkbook.Name)
            Me.DialogResult = DialogResult.Cancel
            Me.Close()
        Else
            ' Use an XmlTextWriter to format.
            ' Make a StringWriter to hold the result.
            Using sw As New System.IO.StringWriter()
                ' Make the XmlTextWriter to format the XML.
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
    End Sub

    Private Sub EditBox_Click(sender As Object, e As EventArgs) Handles EditBox.Click
        Me.PosIndex.Text = "Line: " & (Me.EditBox.GetLineFromCharIndex(Me.EditBox.SelectionStart) + 1).ToString() & ", Column: " & (Me.EditBox.SelectionStart - Me.EditBox.GetFirstCharIndexOfCurrentLine + 1).ToString()
    End Sub

    Private Sub EditBox_KeyDown(sender As Object, e As KeyEventArgs) Handles EditBox.KeyDown
        Me.PosIndex.Text = "Line: " & (Me.EditBox.GetLineFromCharIndex(Me.EditBox.SelectionStart) + 1).ToString() & ", Column: " & (Me.EditBox.SelectionStart - Me.EditBox.GetFirstCharIndexOfCurrentLine + 1).ToString()
    End Sub

End Class