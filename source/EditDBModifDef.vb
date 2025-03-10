﻿Imports ExcelDna.Integration
Imports Microsoft.Office.Core
Imports System.Diagnostics
Imports System.IO
Imports System.Windows.Forms
Imports System.Xml

''' <summary>Dialog used to display and edit the CustomXMLPart utilized for storing the DBModif definitions, and the three parts of DBAddin settings (Addin level, user and central)</summary>
Public Class EditDBModifDef
    ''' <summary>the edited CustomXmlParts for the DBModif definitions</summary>
    Private CustomXmlParts As Object
    ''' <summary>the settings path for user or central setting (for re-saving after modification)</summary>
    Private settingsPath As String = ""

    ''' <summary>put the custom xml definition in the edit box for display/editing</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub EditDBModifDef_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        ' depending on visibility of DBFskip and doDBMOnSave (custom properties available only on Workbook) 
        ' show Workbook level DBModif Definitions
        If Me.DBFskip.Visible Then
            CustomXmlParts = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
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
            Dim availableSettings As String() = Split(My.Resources.SchemaFiles.SettingsDBModif, vbLf) ' avoid dependency on vbCrLf being in the VS settings of Edit/Advanced/set End of Line Sequence
            For Each settingLine As String In availableSettings
                Me.availSettingsLB.Items.Add(settingLine.Replace(vbCr, "")) ' remove vbCr, if compiled with End of Line Sequence vbCrLf
            Next
        Else
            Dim availableSettings As String() = Split(My.Resources.SchemaFiles.Settings, vbLf)
            For Each settingLine As String In availableSettings
                Me.availSettingsLB.Items.Add(settingLine.Replace(vbCr, "")) ' remove vbCr, if compiled with End of Line Sequence vbCrLf
            Next
            ' show DBAddin settings (user/central/addin level): set the appropriate config xml into EditBox (depending on Me.Tag)
            ' get DBAddin user settings and display them
            ' find path of xll:
            For Each tModule As ProcessModule In Process.GetCurrentProcess().Modules
                Dim sModule As String = tModule.FileName
                If sModule.ToUpper.Contains("DBADDIN") Then
                    settingsPath = tModule.FileName
                    Exit For
                End If
            Next
            ' read setting from xll path + ".config": addin level settings
            Me.Text = "DBAddin.xll.config settings in " + settingsPath
            Dim settingsStr As String
            Try
                settingsPath += ".config"
                settingsStr = File.ReadAllText(settingsPath, System.Text.Encoding.Default)
            Catch ex As Exception
                UserMsg("Couldn't read DBAddin.xll.config settings from " + settingsPath + ":" + ex.Message, "Edit DB Addin Settings")
                Exit Sub
            End Try
            ' if central or user settings were chosen, overwrite settingsStr
            If Me.Tag = "central" Or Me.Tag = "user" Then
                ' get central settings filename from DBAddin.xll.config appSettings file attribute or
                ' get user settings filename from DBAddin.xll.config UserSettings configSource attribute 
                Dim doc As New XmlDocument()
                Dim xpathStr As String = If(Me.Tag = "central", "/configuration/appSettings/@file", "/configuration/UserSettings/@configSource")
                doc.LoadXml(settingsStr)
                If Not IsNothing(doc.SelectSingleNode(xpathStr)) Then
                    Dim settingfilename As String = doc.SelectSingleNode(xpathStr).Value
                    ' no path given in central filename: assume it is in same directory
                    If InStr(settingfilename, "\") = 0 Then settingfilename = Replace(settingsPath, "DBaddin.xll.config", "") + settingfilename
                    ' and read central settings
                    settingsPath = settingfilename
                    Try
                        settingsStr = File.ReadAllText(settingsPath, System.Text.Encoding.Default)
                    Catch ex As Exception
                        UserMsg("Couldn't read DB Addin " + Me.Tag + " settings from " + settingsPath + ":" + ex.Message, "Edit DB Addin Settings")
                        Exit Sub
                    End Try
                    Me.Text = Me.Tag + " settings in " + settingsPath
                Else
                    UserMsg("No attribute available as filename reference to " + Me.Tag + " settings (searched xpath: " + xpathStr + ") !", "Edit DB Addin Settings")
                    Exit Sub
                End If
            End If
            Me.OKBtn.Text = "Save"
            Me.ToolTip1.SetToolTip(OKBtn, "save " + Me.Text)
            Me.EditBox.Text = settingsStr
        End If
    End Sub

    ''' <summary>store the displayed/edited text-box content back into the custom xml definition, including validation feedback</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub OKBtn_Click(sender As Object, e As EventArgs) Handles OKBtn.Click
        ' CustomXmlParts of current workbook are saved if dialog is not used as settings display (there DBFskip.Visible would be false as it only applies to a workbook)
        If Me.DBFskip.Visible Then
            ' Make a StringWriter to reformat the indented XML.
            Using sw As New System.IO.StringWriter()
                ' Make a XmlTextWriter to (un)format the XML.
                Using xml_writer As New XmlTextWriter(sw)
                    ' revert indentation...
                    xml_writer.Formatting = Formatting.None
                    Dim doc As New XmlDocument()
                    Try
                        ' validate definition XML
                        Dim schemaString As String = My.Resources.SchemaFiles.DBModifDef
                        Dim schemadoc As XmlReader = XmlReader.Create(New StringReader(schemaString))
                        doc.Schemas.Add("DBModifDef", schemadoc)
                        Dim eventHandler As New Schema.ValidationEventHandler(AddressOf myValidationEventHandler)
                        doc.LoadXml(Me.EditBox.Text)
                        doc.Validate(eventHandler)
                    Catch ex As Exception
                        UserMsg("Problems with parsing changed definition: " + ex.Message, "Edit DB Modifier Definitions XML")
                        Exit Sub
                    End Try
                    doc.WriteTo(xml_writer)
                    xml_writer.Flush()
                    ' store the result in CustomXmlParts
                    CustomXmlParts(1).Delete
                    CustomXmlParts.Add(sw.ToString())
                End Using
            End Using
            ' add/change the tick-boxes doDBMOnSave and DBFskip
            If Not Me.DBFskip.CheckState = CheckState.Indeterminate Then
                Try
                    Try : ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties("DBFskip").Delete : Catch ex As Exception : End Try
                    ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties.Add(Name:="DBFskip", LinkToContent:=False, Type:=MsoDocProperties.msoPropertyTypeBoolean, Value:=Me.DBFskip.Checked)
                Catch ex As Exception
                    UserMsg("Error when adding DBFskip to Workbook:" + ex.Message, "Edit DB Modifier Definitions XML")
                    Exit Sub
                End Try
            End If
            If Not Me.doDBMOnSave.CheckState = CheckState.Indeterminate Then
                Try
                    Try : ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties("doDBMOnSave").Delete : Catch ex As Exception : End Try
                    ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties.Add(Name:="doDBMOnSave", LinkToContent:=False, Type:=MsoDocProperties.msoPropertyTypeBoolean, Value:=Me.doDBMOnSave.Checked)
                Catch ex As Exception
                    UserMsg("Error when adding doDBMOnSave to Workbook:" + ex.Message, "Edit DB Modifier Definitions XML")
                    Exit Sub
                End Try
            End If
        Else
            ' save Addin, users or central settings...
            Dim doc As New XmlDocument()
            Try
                ' validate settings
                Dim schemaString As String = My.Resources.SchemaFiles.DotNetConfig20
                If Me.Tag = "central" Then schemaString = My.Resources.SchemaFiles.DBAddinCentral
                If Me.Tag = "user" Then schemaString = My.Resources.SchemaFiles.DBAddinUser
                Dim schemadoc As XmlReader = XmlReader.Create(New StringReader(schemaString))
                doc.Schemas.Add("", schemadoc)
                Dim eventHandler As New Schema.ValidationEventHandler(AddressOf myValidationEventHandler)
                doc.LoadXml(Me.EditBox.Text)
                doc.Validate(eventHandler)
            Catch ex As Exception
                UserMsg("Problems with parsing changed " + Me.Tag + " settings: " + ex.Message, "Edit DB Addin Settings")
                Exit Sub
            End Try
            Try
                File.WriteAllText(settingsPath, Me.EditBox.Text, System.Text.Encoding.UTF8)
            Catch ex As Exception
                UserMsg("Couldn't write " + Me.Tag + " settings into " + settingsPath + ": " + ex.Message, "Edit DB Addin Settings")
                Exit Sub
            End Try
        End If
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    ''' <summary>validation handler for XML schema (user app settings and DBModif Def) checking</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Sub myValidationEventHandler(sender As Object, e As Schema.ValidationEventArgs)
        ' simply pass back Errors and Warnings as an exception
        If e.Severity = Schema.XmlSeverityType.Error Or e.Severity = Schema.XmlSeverityType.Warning Then Throw New Exception(e.Message)
    End Sub

    ''' <summary>no change was made to definition</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CancelBtn_Click(sender As Object, e As EventArgs) Handles CancelBtn.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    ''' <summary>show the current line and column for easier detection of problems in xml document</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub EditBox_SelectionChanged(sender As Object, e As EventArgs) Handles EditBox.SelectionChanged
        Me.PosIndex.Text = "Line: " + (Me.EditBox.GetLineFromCharIndex(Me.EditBox.SelectionStart) + 1).ToString() + ", Column: " + (Me.EditBox.SelectionStart - Me.EditBox.GetFirstCharIndexOfCurrentLine + 1).ToString()
    End Sub

    ''' <summary>adds the selected setting to the settings (at the end)</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub availSettingsLB_SelectedIndexChanged(sender As Object, e As EventArgs) Handles availSettingsLB.SelectedIndexChanged
        Dim curLineBegin = 0
        Dim settingKey As String = availSettingsLB.Text
        If InStr(settingKey, "+ env") > 0 Then
            Dim envToBeSet As String = InputBox("Environment to use:", "Edit DB Addin Settings", env())
            settingKey = settingKey.Replace(" + env", envToBeSet)
        End If
        Dim lines() As String
        If Me.DBFskip.Visible Then
            ' duplicate "</root>" at the end ...
            Me.EditBox.Text = Me.EditBox.Text + vbCrLf + Me.EditBox.Lines(Me.EditBox.Lines.Length - 1)
            ' replace the but-last line with the third last
            lines = Me.EditBox.Lines
            lines(Me.EditBox.Lines.Length - 2) = lines(Me.EditBox.Lines.Length - 3)
            ' ... and replace the third last line with the new setting
            lines(Me.EditBox.Lines.Length - 3) = "    <" + settingKey + "></" + settingKey + ">"
        Else
            Me.EditBox.SelectAll()
            Me.EditBox.SelectionBackColor = Me.EditBox.BackColor
            For Each editBoxLine In Me.EditBox.Lines
                If InStr(editBoxLine, "<add key=""" + settingKey + """") > 0 Then
                    Me.EditBox.SelectionStart = curLineBegin
                    Me.EditBox.SelectionLength = editBoxLine.Length + 1
                    Me.EditBox.SelectionBackColor = Drawing.Color.Yellow
                    Me.EditBox.ScrollToCaret()
                    UserMsg("Setting " + settingKey + " already exists in " + Me.Tag + " settings", "Edit DB Addin Settings")
                    Exit Sub
                End If
                curLineBegin += editBoxLine.Length + 1
            Next
            ' duplicate "</UserSettings>" at the end ...
            Me.EditBox.Text = Me.EditBox.Text + vbCrLf + Me.EditBox.Lines(Me.EditBox.Lines.Length - 1)
            ' ... and replace the penultimate line with the new setting
            lines = Me.EditBox.Lines
            lines(Me.EditBox.Lines.Length - 2) = "    <add key=""" + settingKey + """ value=""""/>"
        End If
        Me.EditBox.Lines = lines
        Me.EditBox.SelectionStart = Me.EditBox.Text.Length
        Me.EditBox.ScrollToCaret()
    End Sub

    ''' <summary>override paste key combinations to avoid pasting rich text into edit box</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub EditBox_KeyDown(sender As Object, e As KeyEventArgs) Handles EditBox.KeyDown
        If (e.Control And e.KeyCode = Keys.V) Then
            Me.EditBox.Paste(DataFormats.GetFormat("Text"))
            e.Handled = True
        End If
    End Sub

End Class