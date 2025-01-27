Imports ExcelDna.Integration
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.Common
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Windows.Forms


''' <summary>global helper functions for DBModifiers</summary>
Public Module DBModifHelper
    ''' <summary>DBModif definition collections of DBmodif types (key of top level dictionary) with values being collections of DBModifierNames (key of contained dictionaries) and DBModifiers (value of contained dictionaries))</summary>
    Public DBModifDefColl As Dictionary(Of String, Dictionary(Of String, DBModif))
    ''' <summary>main db connection for DB modifiers</summary>
    Public idbcnn As System.Data.IDbConnection
    ''' <summary>avoid entering Application.SheetChange Event handler during listfetch/setquery</summary>
    Public preventChangeWhileFetching As Boolean = False
    ''' <summary>indicates an error in execution of DBModifiers, used for commit/rollback and for non-interactive message return</summary>
    Public hadError As Boolean
    ''' <summary>used to work around the fact that when started by Application.Run, Formulas are sometimes returned as local</summary>
    Public listSepLocal As String = ExcelDnaUtil.Application.International(Excel.XlApplicationInternational.xlListSeparator)
    ''' <summary>common transaction, needed for DBSequence and all other DB Modifiers</summary>
    Public trans As DbTransaction = Nothing
    ''' <summary>opening quote, e.g. [ for SQL Server</summary>
    Public openingQuote As String
    ''' <summary>closing quote, e.g. ] for SQL Server</summary>
    Public closingQuote As String
    ''' <summary>Replacement for closing quote, if needed, e.g. SQL Server requires ] to be replaced by ]]</summary>
    Public closingQuoteReplacement As String
    ''' <summary>Mapping of field names to parameter names (Param + number)</summary>
    Public FieldParamMap As Dictionary(Of String, String)

    ''' <summary>cast .NET data type to ADO.NET DbType</summary>
    ''' <param name="t">given .NET data type</param>
    ''' <returns>ADO.NET DbType</returns>
    Public Function TypeToDbType(t As Type, columnName As String, schemaDataTypeCollection As Collection) As DbType
        ' use the provider specific type information if it exists
        If schemaDataTypeCollection.Contains(columnName) Then
            Select Case schemaDataTypeCollection(columnName)
                Case "char" : TypeToDbType = DbType.AnsiStringFixedLength
                Case "nchar" : TypeToDbType = DbType.StringFixedLength
                Case "varchar" : TypeToDbType = DbType.AnsiString
                Case "nvarchar" : TypeToDbType = DbType.String
                Case "uniqueidentifier" : TypeToDbType = DbType.Guid
                Case "binary" : TypeToDbType = DbType.Binary
                Case "datetime2" : TypeToDbType = DbType.DateTime2
                Case "time" : TypeToDbType = DbType.Time
                Case Else
                    Try
                        TypeToDbType = DirectCast([Enum].Parse(GetType(DbType), t.Name), DbType)
                    Catch ex As Exception
                        TypeToDbType = DbType.Object
                    End Try
            End Select
            Exit Function
        End If
        Try
            TypeToDbType = DirectCast([Enum].Parse(GetType(DbType), t.Name), DbType)
            ' for most string types AnsiString is better
            If TypeToDbType = DbType.String Then TypeToDbType = DbType.AnsiString
        Catch ex As Exception
            TypeToDbType = DbType.Object
        End Try
    End Function

    ''' <summary>opens a database connection</summary>
    ''' <param name="env">number of the environment as given in the settings</param>
    ''' <param name="database">database to replace database selection parameter in connection string of environment</param>
    ''' <returns>True on success</returns>
    Public Function openIdbConnection(env As Integer, database As String) As Boolean
        openIdbConnection = False

        Dim theConnString As String = fetchSetting("ConstConnString" + env.ToString(), "")
        If theConnString = "" Then
            UserMsg("No connection string given for environment: " + env.ToString() + ", please correct and rerun.", "Open Connection Error")
            Exit Function
        End If
        Dim dbidentifier As String = fetchSetting("DBidentifierCCS" + env.ToString(), "")
        If dbidentifier = "" Then
            UserMsg("No DB identifier given for environment: " + env.ToString() + ", please correct and rerun.", "Open Connection Error")
            Exit Function
        End If

        ' change the database in the connection string
        theConnString = Change(theConnString, dbidentifier, database, ";")
        ' need to change/set the connection timeout in the connection string as the property is readonly then...
        If InStr(theConnString, "Connection Timeout=") > 0 Then
            theConnString = Change(theConnString, "Connection Timeout=", CnnTimeout.ToString(), ";")
        ElseIf InStr(theConnString, "Connect Timeout=") > 0 Then
            theConnString = Change(theConnString, "Connect Timeout=", CnnTimeout.ToString(), ";")
        Else
            theConnString += ";Connection Timeout=" + CnnTimeout.ToString()
        End If

        Try
            If Left(theConnString.ToUpper, 5) = "ODBC;" Then
                ' change to ODBC driver setting, if SQLOLEDB
                theConnString = Replace(theConnString, fetchSetting("ConnStringSearch" + env.ToString(), "provider=SQLOLEDB"), fetchSetting("ConnStringReplace" + env.ToString(), "driver=SQL SERVER"))
                ' remove "ODBC;"
                theConnString = Right(theConnString, theConnString.Length - 5)
                idbcnn = New OdbcConnection(theConnString)
            ElseIf InStr(theConnString.ToLower, "provider=sqloledb") Or InStr(theConnString.ToLower, "driver=sql server") Then
                ' remove provider=SQLOLEDB; (or whatever is in ConnStringSearch<>) for sql server as this is not allowed for ado.net (e.g. from a connection string for MS Query/Office)
                theConnString = Replace(theConnString, fetchSetting("ConnStringSearch" + env.ToString(), "provider=SQLOLEDB") + ";", "")
                idbcnn = New SqlConnection(theConnString)
            ElseIf InStr(theConnString.ToLower, "oledb") Then
                idbcnn = New OleDbConnection(theConnString)
            Else
                ' try with odbc
                idbcnn = New OdbcConnection(theConnString)
            End If
        Catch ex As Exception
            UserMsg("Error creating connection object: " + ex.Message + ", connection string: " + theConnString, "Open Connection Error")
            idbcnn = Nothing
            ExcelDnaUtil.Application.StatusBar = False
            Exit Function
        End Try

        LogInfo("open connection with " + theConnString)
        ExcelDnaUtil.Application.StatusBar = "Trying " + CnnTimeout.ToString() + " sec. with connection string: " + theConnString
        Try
            idbcnn.Open()
            openIdbConnection = True
        Catch ex As Exception
            UserMsg("Error connecting to DB: " + ex.Message + ", connection string: " + theConnString, "Open Connection Error")
            idbcnn = Nothing
        End Try
        ExcelDnaUtil.Application.StatusBar = False
    End Function

    ''' <summary>in case there is a defined DBMapper underlying the DBListFetch/DBSetQuery target area then change the extent of it (oldRange) to the new area given in theRange</summary>
    ''' <param name="theRange">new extent after refresh of DBListFetch/DBSetQuery function</param>
    ''' <param name="oldRange">extent before refresh of DBListFetch/DBSetQuery function</param>
    Public Sub resizeDBMapperRange(theRange As Excel.Range, oldRange As Excel.Range)
        Dim actWbNames As Excel.Names
        Try : actWbNames = ExcelDnaUtil.Application.ActiveWorkbook.Names : Catch ex As Exception
            LogWarn("Exception when trying to get the active workbook names for getting DBModifier definitions: " + ex.Message + ", this might be either due to errors in the VBA-IDE (missing references) or due to opening this workbook from an MS-Office hyperlink, starting up Excel (timing issue). Switch to another workbook and back to fix.")
            Exit Sub
        End Try
        ' only do this for the active workbook...
        If theRange.Parent.Parent Is ExcelDnaUtil.Application.ActiveWorkbook Then
            ' getDBModifNameFromRange gets any DBModifName (starting with DBMapper/DBAction...) intersecting theRange, so we can reassign it to the changed range with this...
            Dim dbMapperRangeName As String = getDBModifNameFromRange(theRange)
            ' only allow resizing of dbMapperRange if it was EXACTLY matching the FORMER target range of the DB Function
            If Left(dbMapperRangeName, 8) = "DBMapper" AndAlso oldRange.Address = actWbNames.Item(dbMapperRangeName).RefersToRange.Address Then
                ' (re)assign db mapper range name to the passed (changed) DBListFetch/DBSetQuery function target range
                Try : theRange.Name = dbMapperRangeName
                Catch ex As Exception
                    Throw New Exception("Error when assigning name '" + dbMapperRangeName + "' to DBListFetch/DBSetQuery target range: " + ex.Message)
                End Try
                ' pass the associated DBMapper the new target range
                Try
                    Dim extendedMapper As DBMapper = DBModifDefColl("DBMapper").Item(dbMapperRangeName)
                    extendedMapper.setTargetRange(theRange)
                    extendedMapper.previousCUDLength = theRange.Rows.Count
                Catch ex As Exception
                    Throw New Exception("Error passing new Range to the associated DBMapper object when extending '" + dbMapperRangeName + "' to DBListFetch/DBSetQuery target range: " + ex.Message)
                End Try
            End If
        End If
    End Sub

    ''' <summary>creates a DBModif at the current active cell or edits an existing one defined in targetDefName (after being called in defined range or from ribbon + Ctrl + Shift)</summary>
    ''' <param name="createdDBModifType"></param>
    ''' <param name="targetDefName"></param>
    Public Sub createDBModif(createdDBModifType As String, Optional targetDefName As String = "")
        Dim actWb As Excel.Workbook = Nothing
        Try : actWb = ExcelDnaUtil.Application.ActiveWorkbook : Catch ex As Exception
            UserMsg("Exception when trying to get the active workbook for creating DB Modifier: " + ex.Message + ", this might be either due to errors in the VBA-IDE (missing references) or due to opening this workbook from an MS-Office hyperlink, starting up Excel (timing issue). Switch to another workbook and back to fix.")
            Exit Sub
        End Try
        If IsNothing(actWb) Then Exit Sub
        Dim actWbNames As Excel.Names = Nothing
        Try : actWbNames = actWb.Names : Catch ex As Exception
            UserMsg("Exception when trying to get the active workbook names for creating DB Modifier: " + ex.Message + ", this might be either due to errors in the VBA-IDE (missing references) or due to opening this workbook from an MS-Office hyperlink, starting up Excel (timing issue). Switch to another workbook and back to fix.")
            Exit Sub
        End Try
        Dim existingDBModif As DBModif = Nothing
        Dim existingDefName As String = targetDefName

        ' fetch parameters if there is an existing definition...
        If DBModifDefColl.ContainsKey(createdDBModifType) AndAlso DBModifDefColl(createdDBModifType).ContainsKey(existingDefName) Then
            existingDBModif = DBModifDefColl(createdDBModifType).Item(existingDefName)
            ' reset the target range to a potentially changed area
            If createdDBModifType <> "DBSeqnce" Then
                Dim existingDefRange As Excel.Range = Nothing
                Try
                    existingDefRange = ExcelDnaUtil.Application.Range(existingDefName)
                Catch ex As Exception
                    ' if target name relates to an invalid (offset) formula, getting a range fails  ...
                    If InStr(actWbNames.Item(existingDefName).RefersTo, "OFFSET(") > 0 Then
                        UserMsg("Offset formula that '" + existingDefName + "' refers to, did not return a valid range." + vbCrLf + "Please check the offset formula to return a valid range !", "DBModifier Definitions Error")
                        ExcelDnaUtil.Application.Dialogs(Excel.XlBuiltInDialog.xlDialogNameManager).Show()
                        Exit Sub
                    End If
                End Try
                existingDBModif.setTargetRange(existingDefRange)
            End If
        End If

        ' prepare DBModifier Create Dialog
        Dim theDBModifCreateDlg As New DBModifCreate()
        With theDBModifCreateDlg
            ' store DBModification type in tag for validation purposes...
            .Tag = createdDBModifType
            .envSel.DataSource = environdefs
            .envSel.SelectedIndex = -1
            .DBModifName.Text = Replace(existingDefName, createdDBModifType, "")
            .DBModifName.Tag = existingDefName
            .RepairDBSeqnce.Hide()
            .NameLabel.Text = IIf(createdDBModifType = "DBSeqnce", "DBSequence", createdDBModifType) + " Name:"
            .Text = "Edit " + IIf(createdDBModifType = "DBSeqnce", "DBSequence", createdDBModifType) + " definition"
            If createdDBModifType <> "DBMapper" Then
                .TablenameLabel.Hide()
                .PrimaryKeysLabel.Hide()
                .AdditionalStoredProcLabel.Hide()
                .IgnoreColumnsLabel.Hide()
                .Tablename.Hide()
                .PrimaryKeys.Hide()
                .insertIfMissing.Hide()
                .addStoredProc.Hide()
                .IgnoreColumns.Hide()
                .CUDflags.Hide()
                .AutoIncFlag.Hide()
                .IgnoreDataErrors.Hide()
            End If
            If createdDBModifType = "DBAction" Then
                .paramRangesStr.Top = .Tablename.Top
                .TablenameLabel.Show()
                .TablenameLabel.Text = "Parameter Range Names:"
                .paramRangesStr.Left = .Tablename.Left
                .paramEnclosing.Top = .PrimaryKeys.Top
                .PrimaryKeysLabel.Show()
                .PrimaryKeysLabel.Text = "Parameter enclosing char:"
                .paramEnclosing.Left = .PrimaryKeys.Left
                .convertAsDate.Top = .IgnoreColumns.Top
                .IgnoreColumnsLabel.Text = "Cols num params date:"
                .IgnoreColumnsLabel.Show()
                .convertAsDate.Left = .IgnoreColumns.Left
                .convertAsString.Top = .addStoredProc.Top
                .AdditionalStoredProcLabel.Show()
                .AdditionalStoredProcLabel.Text = "Cols num params string:"
                .convertAsString.Left = .addStoredProc.Left
                .parametrized.Top = .CUDflags.Top
                .parametrized.Left = .CUDflags.Left
                .continueIfRowEmpty.Top = .IgnoreDataErrors.Top
                .continueIfRowEmpty.Left = .IgnoreDataErrors.Left
            Else
                .TablenameLabel.Text = "Tablename:"
                .PrimaryKeysLabel.Text = "Primary keys count:"
                .IgnoreColumnsLabel.Text = "Ignore columns:"
                .AdditionalStoredProcLabel.Text = "Additional stored procedure:"
                .parametrized.Hide()
                .paramRangesStr.Hide()
                .paramEnclosing.Hide()
                .convertAsDate.Hide()
                .convertAsString.Hide()
                .continueIfRowEmpty.Hide()
            End If
            If createdDBModifType = "DBSeqnce" Then
                theDBModifCreateDlg.FormBorderStyle = FormBorderStyle.Sizable
                ' hide controls irrelevant for DBSeqnce
                .TargetRangeAddress.Hide()
                .envSel.Hide()
                .EnvironmentLabel.Hide()
                .Database.Hide()
                .DatabaseLabel.Hide()
                .DBSeqenceDataGrid.Top = 55
                .DBSeqenceDataGrid.Height = 320
                .execOnSave.Top = .CreateCB.Top
                .AskForExecute.Top = .CreateCB.Top
                .execOnSave.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
                .AskForExecute.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
                ' fill Data grid-view for DBSequence
                Dim cb As New DataGridViewComboBoxColumn With {
                        .HeaderText = "Sequence Step",
                        .ReadOnly = False
                    }
                cb.ValueType() = GetType(String)
                Dim ds As New List(Of String)

                ' first add the DBMapper and DBAction definitions available in the Workbook
                For Each DBModiftype As String In DBModifDefColl.Keys
                    ' avoid DB Sequences (might be - indirectly - self referencing, leading to endless recursion)
                    If DBModiftype <> "DBSeqnce" Then
                        For Each nodeName As String In DBModifDefColl(DBModiftype).Keys
                            ds.Add(DBModiftype + ":" + nodeName)
                        Next
                    End If
                Next

                ' then add DBRefresh items for allowing refreshing DBFunctions (DBListFetch and DBSetQuery) during a Sequence
                Dim searchCell As Excel.Range
                For Each ws As Excel.Worksheet In actWb.Worksheets
                    ExcelDnaUtil.Application.Statusbar = "Looking for DBFunctions in " + ws.Name + " for adding possibility to DB Sequence"
                    For Each theFunc As String In {"DBListFetch(", "DBSetQuery(", "DBRowFetch("}
                        searchCell = ws.Cells.Find(What:=theFunc, After:=ws.Range("A1"), LookIn:=Excel.XlFindLookIn.xlFormulas, LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlNext, MatchCase:=False)
                        Dim firstFoundAddress As String = ""
                        If searchCell IsNot Nothing Then firstFoundAddress = searchCell.Address
                        While searchCell IsNot Nothing
                            Dim underlyingName As String = getUnderlyingDBNameFromRange(searchCell)
                            ds.Add("Refresh " + theFunc + searchCell.Parent.Name + "!" + searchCell.Address + "):" + underlyingName)
                            searchCell = ws.Cells.FindNext(searchCell)
                            If searchCell.Address = firstFoundAddress Then Exit While
                        End While
                    Next
                    ' reset the cell find dialog....
                    searchCell = Nothing
                    searchCell = ws.Cells.Find(What:="", After:=ws.Range("A1"), LookIn:=Excel.XlFindLookIn.xlFormulas, LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlNext, MatchCase:=False)
                Next
                ExcelDnaUtil.Application.Statusbar = False
                ' at last add special items DBBeginTrans and DBCommitTrans for setting DB Transaction brackets
                ds.Add("DBBegin:Begins DB Transaction")
                ds.Add("DBCommitRollback:Commits or Rolls back DB Transaction")
                ' and bind the dataset to the combo-box
                cb.DataSource() = ds
                .DBSeqenceDataGrid.Columns.Add(cb)
                .DBSeqenceDataGrid.Columns(0).Width = 400
            Else
                theDBModifCreateDlg.FormBorderStyle = FormBorderStyle.FixedDialog
                theDBModifCreateDlg.MinimumSize = New Drawing.Size(width:=490, height:=290)
                theDBModifCreateDlg.Size = New Drawing.Size(width:=490, height:=290)
                ' hide controls irrelevant for DBMapper and DBAction
                .DBSeqenceDataGrid.Hide()
            End If

            ' delegate filling of dialog fields to created DBModif object
            If existingDBModif IsNot Nothing Then existingDBModif.setDBModifCreateFields(theDBModifCreateDlg)
            ' reflect parametrized settings of DBAction in GUI
            theDBModifCreateDlg.setDBActionParametrizedGUI()

            ' display dialog for parameters
            If theDBModifCreateDlg.ShowDialog() = DialogResult.Cancel Then Exit Sub

            ' only for DBMapper or DBAction: change or add target range name, for DBAction check template placeholders
            If createdDBModifType <> "DBSeqnce" Then
                Dim targetRange As Excel.Range
                If existingDBModif Is Nothing Then
                    targetRange = ExcelDnaUtil.Application.Selection
                Else
                    targetRange = existingDBModif.getTargetRange()
                End If

                If existingDefName = "" Then
                    Try
                        actWbNames.Add(Name:=createdDBModifType + .DBModifName.Text, RefersTo:=targetRange)
                    Catch ex As Exception
                        UserMsg("Error when assigning range name '" + createdDBModifType + .DBModifName.Text + "' to active cell: " + ex.Message, "DBModifier Creation Error")
                        Exit Sub
                    End Try
                Else
                    ' rename named range...
                    actWbNames.Item(existingDefName).Name = createdDBModifType + .DBModifName.Text
                End If

                ' cross check with template parameter placeholders
                If createdDBModifType = "DBAction" And theDBModifCreateDlg.paramRangesStr.Text <> "" Then
                    Dim templateSQL As String = ""
                    For Each aCell As Excel.Range In targetRange
                        templateSQL += aCell.Value
                    Next
                    Dim paramEnclosing As String = IIf(theDBModifCreateDlg.paramEnclosing.Text = "", "!", theDBModifCreateDlg.paramEnclosing.Text)
                    Dim paramNum As Integer = 0
                    For Each paramRange In Split(theDBModifCreateDlg.paramRangesStr.Text, ",")
                        paramNum += 1 : Dim placeHolder As String = paramEnclosing + paramNum.ToString() + paramEnclosing
                        If InStr(templateSQL, placeHolder) = 0 Then
                            UserMsg("Didn't find a corresponding placeholder (" + placeHolder + ") in DBAction template SQL for parameter " + paramNum.ToString() + ", this might be an error!", "DBAction Validation", MsgBoxStyle.Exclamation)
                        End If
                        templateSQL = templateSQL.Replace(placeHolder, "match" + paramNum.ToString())
                    Next
                    If templateSQL Like "*" + paramEnclosing + "*" + paramEnclosing + "*" Then
                        UserMsg("found placeholders (" + paramEnclosing + "*" + paramEnclosing + ") not covered by parameters in DBAction template SQL (" + templateSQL + "), this might be an error!", "DBAction Validation", MsgBoxStyle.Exclamation)
                    End If
                End If
            End If

            Dim CustomXmlParts As Object = actWb.CustomXMLParts.SelectByNamespace("DBModifDef")
            If CustomXmlParts.Count = 0 Then
                ' in case no CustomXmlPart in Namespace DBModifDef exists in the workbook, add one
                actWb.CustomXMLParts.Add("<root xmlns=""DBModifDef""></root>")
                CustomXmlParts = actWb.CustomXMLParts.SelectByNamespace("DBModifDef")
            End If

            ' remove old node in case of renaming DBModifier
            ' Elements have names of DBModif types, attribute Name is given name (<DBMapper Name=existingDefName>)
            If Not IsNothing(CustomXmlParts(1).SelectSingleNode("/ns0:root/ns0:" + createdDBModifType + "[@Name='" + Replace(existingDefName, createdDBModifType, "") + "']")) Then
                CustomXmlParts(1).SelectSingleNode("/ns0:root/ns0:" + createdDBModifType + "[@Name='" + Replace(existingDefName, createdDBModifType, "") + "']").Delete
            End If

            ' NamespaceURI:="DBModifDef" is required to avoid adding a xmlns attribute to each element.
            CustomXmlParts(1).SelectSingleNode("/ns0:root").AppendChildNode(createdDBModifType, NamespaceURI:="DBModifDef")
            ' new appended elements are last, get it to append further child elements
            Dim dbModifNode As CustomXMLNode = CustomXmlParts(1).SelectSingleNode("/ns0:root").LastChild
            ' append the detailed settings to the definition element
            dbModifNode.AppendChildNode("Name", NodeType:=MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, NodeValue:= .DBModifName.Text)
            dbModifNode.AppendChildNode("execOnSave", NamespaceURI:="DBModifDef", NodeValue:= .execOnSave.Checked.ToString())
            dbModifNode.AppendChildNode("askBeforeExecute", NamespaceURI:="DBModifDef", NodeValue:= .AskForExecute.Checked.ToString())
            If createdDBModifType = "DBMapper" Then
                dbModifNode.AppendChildNode("env", NamespaceURI:="DBModifDef", NodeValue:=(.envSel.SelectedIndex + 1).ToString()) ' if not selected, set environment to 0 (default anyway)
                dbModifNode.AppendChildNode("database", NamespaceURI:="DBModifDef", NodeValue:= .Database.Text)
                dbModifNode.AppendChildNode("tableName", NamespaceURI:="DBModifDef", NodeValue:= .Tablename.Text)
                dbModifNode.AppendChildNode("primKeysStr", NamespaceURI:="DBModifDef", NodeValue:= .PrimaryKeys.Text)
                dbModifNode.AppendChildNode("insertIfMissing", NamespaceURI:="DBModifDef", NodeValue:= .insertIfMissing.Checked.ToString())
                dbModifNode.AppendChildNode("executeAdditionalProc", NamespaceURI:="DBModifDef", NodeValue:= .addStoredProc.Text)
                dbModifNode.AppendChildNode("ignoreColumns", NamespaceURI:="DBModifDef", NodeValue:= .IgnoreColumns.Text)
                dbModifNode.AppendChildNode("CUDFlags", NamespaceURI:="DBModifDef", NodeValue:= .CUDflags.Checked.ToString())
                dbModifNode.AppendChildNode("AutoIncFlag", NamespaceURI:="DBModifDef", NodeValue:= .AutoIncFlag.Checked.ToString())
                dbModifNode.AppendChildNode("IgnoreDataErrors", NamespaceURI:="DBModifDef", NodeValue:= .IgnoreDataErrors.Checked.ToString())
            ElseIf createdDBModifType = "DBAction" Then
                dbModifNode.AppendChildNode("env", NamespaceURI:="DBModifDef", NodeValue:=(.envSel.SelectedIndex + 1).ToString())
                dbModifNode.AppendChildNode("database", NamespaceURI:="DBModifDef", NodeValue:= .Database.Text)
                dbModifNode.AppendChildNode("parametrized", NamespaceURI:="DBModifDef", NodeValue:= .parametrized.Checked.ToString())
                dbModifNode.AppendChildNode("continueIfRowEmpty", NamespaceURI:="DBModifDef", NodeValue:= .continueIfRowEmpty.Checked.ToString())
                If .paramRangesStr.Text <> "" Then dbModifNode.AppendChildNode("paramRangesStr", NamespaceURI:="DBModifDef", NodeValue:= .paramRangesStr.Text)
                If .paramEnclosing.Text <> "" Then dbModifNode.AppendChildNode("paramEnclosing", NamespaceURI:="DBModifDef", NodeValue:= .paramEnclosing.Text)
                If .convertAsDate.Text <> "" Then dbModifNode.AppendChildNode("convertAsDate", NamespaceURI:="DBModifDef", NodeValue:= .convertAsDate.Text)
                If .convertAsString.Text <> "" Then dbModifNode.AppendChildNode("convertAsString", NamespaceURI:="DBModifDef", NodeValue:= .convertAsString.Text)
            ElseIf createdDBModifType = "DBSeqnce" Then
                ' "repaired" mode (indicating rewriting DBSequence Steps)
                If .Tag = "repaired" Then
                    Dim repairedSequence() As String = Split(.RepairDBSeqnce.Text, vbCrLf)
                    For i As Integer = 0 To UBound(repairedSequence)
                        dbModifNode.AppendChildNode("seqStep", NamespaceURI:="DBModifDef", NodeValue:=repairedSequence(i))
                    Next
                Else
                    For i As Integer = 0 To .DBSeqenceDataGrid.Rows().Count - 2
                        dbModifNode.AppendChildNode("seqStep", NamespaceURI:="DBModifDef", NodeValue:= .DBSeqenceDataGrid.Rows(i).Cells(0).Value)
                    Next
                End If
            End If
            ' any features added directly to DBModif definition in XML need to be re-added now
            If existingDBModif IsNot Nothing Then existingDBModif.addHiddenFeatureDefs(dbModifNode)
            ' refresh mapper definitions to reflect changes immediately...
            getDBModifDefinitions(actWb)
            ' extend Data-range for new DBMappers immediately after definition...
            If createdDBModifType = "DBMapper" Then
                DirectCast(DBModifDefColl("DBMapper").Item(createdDBModifType + .DBModifName.Text), DBMapper).extendDataRange()
            End If

        End With
    End Sub

    ''' <summary>check one param range input (name) and return the range if successful</summary>
    ''' <param name="paramRange">name of parameter range</param>
    ''' <returns>the range of the name</returns>
    Public Function checkAndReturnRange(paramRange As String) As Excel.Range
        Dim actWbNames As Excel.Names
        Try : actWbNames = ExcelDnaUtil.Application.ActiveWorkbook.Names : Catch ex As Exception
            Throw New Exception("Exception when trying to get the active workbook names for executeTemplateSQL: " + ex.Message + ", this might be either due to errors in the VBA-IDE (missing references) or due to opening this workbook from an MS-Office hyperlink, starting up Excel (timing issue). Switch to another workbook and back to fix.")
        End Try
        ' either get the range from a workbook based name or current sheet name (no ! in name)
        If InStr(paramRange, "!") = 0 Then
            If Not existsName(paramRange) Then
                Throw New Exception("Name '" + paramRange + "' doesn't exist as a workbook name (you need to qualify names defined in worksheets with sheet_name!range_name).")
            Else
                checkAndReturnRange = actWbNames.Item(paramRange).RefersToRange
            End If
        Else
            ' .. or from a worksheet based name from a sheet
            Dim wsNameParts() As String = Split(paramRange, "!")
            Dim sheetName As String = wsNameParts(0).Replace("'", "")
            Dim nameSheet = ExcelDnaUtil.Application.ActiveWorkbook.Worksheets(sheetName)
            If existsSheet(sheetName, ExcelDnaUtil.Application.ActiveWorkbook) Then
                If nameSheet Is ExcelDnaUtil.Application.ActiveSheet Then
                    ' different access to names from current sheet, these are in actWbNames with full qualification
                    If existsName(paramRange) Then
                        checkAndReturnRange = actWbNames.Item(paramRange).RefersToRange
                    Else
                        Throw New Exception("Name '" + paramRange + "' is not defined in current worksheet")
                    End If
                Else
                    If existsNameInSheet(wsNameParts(1), nameSheet) Then
                        checkAndReturnRange = getRangeFromNameInSheet(wsNameParts(1), nameSheet)
                    Else
                        Throw New Exception("Name '" + paramRange + "' is not defined in worksheet '" + sheetName + "'")
                    End If
                End If
            Else
                Throw New Exception("Sheet '" + sheetName + "' referred to in '" + paramRange + "' does not exist in active workbook")
            End If
        End If
    End Function

    ''' <summary>gets defined names for DBModifier (DBMapper/DBAction/DBSeqnce) invocation in the passed workbook and updates Ribbon with it</summary>
    ''' <param name="actWb">the passed workbook</param>
    ''' <param name="onlyCheck">only used for UserMsg dialogs (check vs. get)</param>
    Public Sub getDBModifDefinitions(actWb As Excel.Workbook, Optional onlyCheck As Boolean = False)

        ' load DBModifier definitions (objects) into Global collection DBModifDefColl
        LogInfo("reading DBModifier Definitions for Workbook: " + actWb.Name)
        Try
            DBModifDefColl.Clear()
            Dim CustomXmlParts As Object = actWb.CustomXMLParts.SelectByNamespace("DBModifDef")
            If CustomXmlParts.Count = 1 Then
                Dim actWbNames As Excel.Names
                Try : actWbNames = actWb.Names : Catch ex As Exception
                    UserMsg("Exception when trying to get the active workbook names for getting DBModifier definitions: " + ex.Message + ", this might be either due to errors in the VBA-IDE (missing references) or due to opening this workbook from an MS-Office hyperlink, starting up Excel (timing issue). Switch to another workbook and back to fix.")
                    Exit Sub
                End Try

                ' read DBModifier definitions from CustomXMLParts
                For Each customXMLNodeDef As CustomXMLNode In CustomXmlParts(1).SelectSingleNode("/ns0:root").ChildNodes
                    Dim DBModiftype As String = Left(customXMLNodeDef.BaseName, 8)
                    If DBModiftype = "DBSeqnce" Or DBModiftype = "DBMapper" Or DBModiftype = "DBAction" Then
                        Dim nodeName As String
                        If customXMLNodeDef.Attributes.Count > 0 Then
                            nodeName = DBModiftype + customXMLNodeDef.Attributes(1).Text
                        Else
                            nodeName = customXMLNodeDef.BaseName
                        End If
                        LogInfo("reading DBModifier Definition for " + nodeName)
                        Dim targetRange As Excel.Range = Nothing
                        ' for DBMappers and DBActions the data of the DBModification is stored in Ranges, so check for those and get the Range
                        If DBModiftype = "DBMapper" Or DBModiftype = "DBAction" Then
                            For Each rangename As Excel.Name In actWbNames
                                Dim rangenameName As String = Replace(rangename.Name, rangename.Parent.Name + "!", "")
                                If rangenameName = nodeName Then
                                    If InStr(rangename.RefersTo, "#REF!") > 0 Then
                                        UserMsg(DBModiftype + " definitions range " + rangename.Name + " contains #REF!", "DBModifier Definitions Error")
                                        Exit For
                                    End If
                                    ' might fail if target name relates to an invalid (offset) formula ...
                                    Try
                                        targetRange = rangename.RefersToRange
                                    Catch ex As Exception
                                        If InStr(rangename.RefersTo, "OFFSET(") > 0 Then
                                            UserMsg("Offset formula that '" + nodeName + "' refers to, did not return a valid range." + vbCrLf + "Please check the offset formula to return a valid range !", "DBModifier Definitions Error")
                                            ExcelDnaUtil.Application.Dialogs(Excel.XlBuiltInDialog.xlDialogNameManager).Show()
                                            GoTo EndOuterLoop
                                        End If
                                    End Try
                                    Exit For
                                End If
                            Next
                            If targetRange Is Nothing Then
                                Dim answer As MsgBoxResult = QuestionMsg("Required target range named '" + nodeName + "' cannot be found for this " + DBModiftype + " definition." + vbCrLf + "Should the target range name and definition be removed (If you still need the " + DBModiftype + ", (re)create the target range with this name again)?", , "DBModifier Definitions Error", MsgBoxStyle.Critical)
                                If answer = MsgBoxResult.Ok Then
                                    ' remove name, in case it still exists
                                    Try : actWbNames.Item(nodeName).Delete() : Catch ex As Exception : End Try
                                    ' remove node
                                    If Not IsNothing(CustomXmlParts(1).SelectSingleNode("/ns0:root/ns0:" + DBModiftype + "[@Name='" + Replace(nodeName, DBModiftype, "") + "']")) Then
                                        Try : CustomXmlParts(1).SelectSingleNode("/ns0:root/ns0:" + DBModiftype + "[@Name='" + Replace(nodeName, DBModiftype, "") + "']").Delete : Catch ex As Exception
                                            UserMsg("Error removing node in DBModif definitions: " + ex.Message)
                                        End Try
                                    End If
                                End If
                                Continue For
                            End If
                        End If
                        ' finally create the DBModif Object ...
                        Dim newDBModif As DBModif = Nothing
                        ' fill parameters into CustomXMLPart:
                        If DBModiftype = "DBMapper" Then
                            newDBModif = New DBMapper(customXMLNodeDef, targetRange)
                        ElseIf DBModiftype = "DBAction" Then
                            newDBModif = New DBAction(customXMLNodeDef, targetRange)
                        ElseIf DBModiftype = "DBSeqnce" Then
                            newDBModif = New DBSeqnce(customXMLNodeDef)
                        Else
                            UserMsg("Not supported DBModifier type: " + DBModiftype, "DBModifier Definitions Error")
                        End If
                        ' ... and add it to the collection DBModifDefColl
                        Dim defColl As Dictionary(Of String, DBModif) ' definition lookup collection for DBModifiername -> object
                        If newDBModif IsNot Nothing Then
                            If Not DBModifDefColl.ContainsKey(DBModiftype) Then
                                ' add to new DBModiftype "menu"
                                defColl = New Dictionary(Of String, DBModif) From {
                                        {nodeName, newDBModif}
                                    }
                                DBModifDefColl.Add(DBModiftype, defColl)
                            Else
                                ' add definition to existing DBModiftype "menu"
                                defColl = DBModifDefColl(DBModiftype)
                                If defColl.ContainsKey(nodeName) Then
                                    UserMsg("DBModifier " + nodeName + " added twice, this potentially indicates legacy definitions that were modified!" + vbCrLf + "To fix, convert all other definitions in the same way and then remove the legacy definitions by editing the raw DB Modif definitions.", IIf(onlyCheck, "check", "get") + " DBModif Definitions")
                                Else
                                    defColl.Add(nodeName, newDBModif)
                                End If
                            End If
                        End If
                    End If
EndOuterLoop:
                Next
            ElseIf CustomXmlParts.Count > 1 Then
                UserMsg("Multiple CustomXmlParts for DBModifDef existing!", IIf(onlyCheck, "check", "get") + " DBModif Definitions")
            End If
            theRibbon.Invalidate()
        Catch ex As Exception
            UserMsg("Exception in getting DB Modifier Definitions: " + ex.Message, "DBModifier Definitions Error")
        End Try
    End Sub

    ''' <summary>correct quotes in field name</summary>
    ''' <param name="fieldname">field name to correct</param>
    ''' <returns>quote corrected field name</returns>
    Public Function CorrectQuotes(fieldname As String) As String
        CorrectQuotes = Replace(fieldname, closingQuote, closingQuoteReplacement)
    End Function

    ''' <summary>gets DB Modification Name (DBMapper or DBAction) from theRange</summary>
    ''' <param name="theRange"></param>
    ''' <returns>the retrieved name as a string (not name object !)</returns>
    Public Function getDBModifNameFromRange(theRange As Excel.Range) As String
        Dim nm As Excel.Name
        Dim rng, testRng As Excel.Range
        Dim theWbNames As Excel.Names

        getDBModifNameFromRange = ""
        If theRange Is Nothing Then Exit Function
        Try : theWbNames = theRange.Parent.Parent.Names : Catch ex As Exception
            UserMsg("Exception getting the range's parent workbook names: " + ex.Message + ", this might be either due to errors in the VBA-IDE (missing references) or due to opening this workbook from an MS-Office hyperlink, starting up Excel (timing issue). Switch to another workbook and back to fix.")
            Exit Function
        End Try
        Try
            ' try all names in workbook
            For Each nm In theWbNames
                rng = Nothing
                ' test whether range referring to that name (if it is a real range)...
                Try : rng = nm.RefersToRange : Catch ex As Exception : End Try
                If rng IsNot Nothing Then
                    testRng = Nothing
                    ' ...intersects with the passed range
                    Try : testRng = ExcelDnaUtil.Application.Intersect(theRange, rng) : Catch ex As Exception : End Try
                    If testRng IsNot Nothing And (InStr(1, nm.Name, "DBMapper") >= 1 Or InStr(1, nm.Name, "DBAction") >= 1) Then
                        ' and pass back the name if it does and is a DBMapper or a DBAction
                        getDBModifNameFromRange = nm.Name
                        Exit Function
                    End If
                End If
            Next
        Catch ex As Exception
            UserMsg("Exception: " + ex.Message, "get DBModif Name From Range")
        End Try
    End Function

    ''' <summary>To check for errors in passed range obj, makes use of the fact that Range.Value never passes Integer Values back except for Errors</summary>
    ''' <param name="rangeval">Range.Value to be checked for errors</param>
    ''' <remarks>https://xldennis.wordpress.com/2006/11/22/dealing-with-cverr-values-in-net-%E2%80%93-part-i-the-problem/ and https://xldennis.wordpress.com/2006/11/29/dealing-with-cverr-values-in-net-part-ii-solutions/</remarks>
    ''' <returns>true if error</returns>
    Public Function IsXLCVErr(rangeval As Object) As Boolean
        Return TypeOf (rangeval) Is Int32
    End Function

    ''' <summary>to convert the error number to text</summary>
    ''' <param name="whichError">integer error number</param>
    ''' <returns>text of error</returns>
    Public Function CVErrText(whichError As Integer) As String
        Select Case whichError
            Case -2146826281 : Return "#Div0!"
            Case -2146826245 : Return "#GettingData"
            Case -2146826246 : Return "#N/A"
            Case -2146826259 : Return "#Name"
            Case -2146826288 : Return "#Null!"
            Case -2146826252 : Return "#Num!"
            Case -2146826265 : Return "#Ref!"
            Case -2146826273 : Return "#Value!"
            Case Else : Return "unknown error !!"
        End Select
    End Function

    ''' <summary>execute given DBModifier, used for VBA call by Application.Run</summary>
    ''' <param name="DBModifName">Full name of DB Modifier, including type at beginning</param>
    ''' <param name="headLess">if set to true, DBAddin will avoid to issue messages and return messages in exceptions which are returned (headless)</param>
    ''' <returns>empty string on success, error message otherwise</returns>
    <ExcelCommand(Name:="executeDBModif")>
    Public Function executeDBModif(DBModifName As String, Optional headLess As Boolean = False) As String
        hadError = False : nonInteractive = headLess
        nonInteractiveErrMsgs = "" ' reset non-interactive messages
        Dim DBModiftype As String = Left(DBModifName, 8)
        If DBModiftype = "DBSeqnce" Or DBModiftype = "DBMapper" Or DBModiftype = "DBAction" Then
            If Not DBModifDefColl(DBModiftype).ContainsKey(DBModifName) Then
                If DBModifDefColl(DBModiftype).Count = 0 Then
                    nonInteractive = False
                    Return "No DBModifier contained in Workbook at all!"
                End If
                Dim DBModifavailable As String = ""
                For Each DBMtype As String In {"DBMapper", "DBAction", "DBSeqnce"}
                    For Each DBMkey As String In DBModifDefColl(DBMtype).Keys
                        DBModifavailable += "," + DBMkey
                    Next
                Next
                nonInteractive = False
                Return "DB Modifier '" + DBModifName + "' not existing, available: " + DBModifavailable
            End If
            LogInfo("Doing DBModifier '" + DBModifName + "' ...")
            Try
                DBModifDefColl(DBModiftype).Item(DBModifName).doDBModif()
            Catch ex As Exception
                nonInteractive = False
                Return "DB Modifier '" + DBModifName + "' doDBModif had following error(s): " + ex.Message
            End Try
            nonInteractive = False
            If hadError Then Return nonInteractiveErrMsgs
        ElseIf DBModiftype = "Refresh " Then
            ' DBModifName for DBfunction refresh is "Refresh Sheet-name!Address" where Sheet-name!Address is a cell containing the DBfunction 
            Dim RangeParts() As String = Split(Mid(DBModifName, 9), "!")
            If RangeParts.Length = 2 And RangeParts(0) <> "" And RangeParts(1) <> "" Then
                Dim SheetName = Replace(RangeParts(0), "'", "") ' for sheet-names with blanks surrounding quotations are needed, remove them here
                Dim Address = RangeParts(1)
                Dim srcExtent As String = ""
                Try : srcExtent = getUnderlyingDBNameFromRange(ExcelDnaUtil.Application.Worksheets(SheetName).Range(Address)) : Catch ex As Exception : End Try
                If srcExtent = "" Then Return "No valid address found in " + DBModifName + " (Sheet-name: " + SheetName + ", Address: " + Address + ")"
                Dim aDBModifier As New DBModifDummy()
                aDBModifier.executeRefresh(srcExtent)
            Else
                Return "No Worksheet/Address could be parsed from " + DBModifName
            End If
        Else
            nonInteractive = False
            Return "No valid type (" + DBModiftype + ") in passed DB Modifier '" + DBModifName + "', DB Modifier name must start with 'DBSeqnce', 'DBMapper' Or 'DBAction' !"
        End If
        Return "" ' no error, no message
    End Function

    ''' <summary>set given execution parameter, used for VBA call by Application.Run</summary>
    ''' <param name="Param">execution parameter, like "selectedEnvironment" (zero based here!) or "CnnTimeout"</param>
    ''' <param name="Value">execution parameter value</param>
    <ExcelCommand(Name:="setExecutionParam")>
    Public Sub setExecutionParam(Param As String, Value As Object)
        Try
            If Param = "headLess" Then
                nonInteractive = Value
                nonInteractiveErrMsgs = "" ' reset non-interactive messages
            ElseIf Param = "selectedEnvironment" Then
                SettingsTools.selectedEnvironment = Value
                theRibbon.InvalidateControl("envDropDown")
            ElseIf Param = "ConstConnString" Then
                SettingsTools.ConstConnString = Value
            ElseIf Param = "CnnTimeout" Then
                SettingsTools.CnnTimeout = Value
            ElseIf Param = "CmdTimeout" Then
                SettingsTools.CmdTimeout = Value
            ElseIf Param = "preventRefreshFlag" Then
                Functions.preventRefreshFlag = Value
                theRibbon.InvalidateControl("preventRefresh")
            Else
                UserMsg("parameter " + Param + " not supported by setExecutionParams")
                Exit Sub
            End If
        Catch ex As Exception
            UserMsg("setting parameter " + Param + " with value " + CStr(Value) + " resulted in error " + ex.Message)
        End Try
    End Sub

    ''' <summary>get given execution parameter or setting parameter found by fetchSetting, used for VBA call by Application.Run</summary>
    ''' <param name="Param">execution parameter, like "selectedEnvironment" (zero based here!), "env()" or "CnnTimeout"</param>
    ''' <returns>execution or setting parameter value</returns>
    <ExcelCommand(Name:="getExecutionParam")>
    Public Function getExecutionParam(Param As String) As Object
        If Param = "selectedEnvironment" Then
            Return SettingsTools.selectedEnvironment
        ElseIf Param = "env()" Then
            Return SettingsTools.env()
        ElseIf Param = "ConstConnString" Then
            Return SettingsTools.ConstConnString
        ElseIf Param = "CnnTimeout" Then
            Return SettingsTools.CnnTimeout
        ElseIf Param = "CmdTimeout" Then
            Return SettingsTools.CmdTimeout
        ElseIf Param = "preventRefreshFlag" Then
            Return Functions.preventRefreshFlag
        Else
            Return fetchSetting(Param, "parameter " + Param + " neither supported by getExecutionParam nor found with fetchSetting(Param)")
        End If
    End Function

    ''' <summary>marks a row in a DBMapper for deletion, used as a ExcelCommand to have a keyboard shortcut</summary>
    <ExcelCommand(Name:="deleteRow")>
    Public Sub deleteRow()
        Dim targetName As String = getDBModifNameFromRange(ExcelDnaUtil.Application.Selection)
        If Left(targetName, 8) = "DBMapper" Then DirectCast(DBModifDefColl("DBMapper").Item(targetName), DBMapper).insertCUDMarks(ExcelDnaUtil.Application.Selection, deleteFlag:=True)
    End Sub

    ''' <summary>inserts a row in a DBMapper, used as a ExcelCommand to have a keyboard shortcut</summary>
    <ExcelCommand(Name:="insertRow")>
    Public Sub insertRow()
        Dim targetName As String = getDBModifNameFromRange(ExcelDnaUtil.Application.Selection)
        If Left(targetName, 8) = "DBMapper" Then
            ' get the target range for the DBMapper to get the ListObject
            Dim insertTarget As Excel.Range = DirectCast(DBModifDefColl("DBMapper").Item(targetName), DBMapper).getTargetRange
            ' calculate insert row from selection and top row of insert target
            Dim insertRow As Integer = ExcelDnaUtil.Application.Selection.Row - insertTarget.Row
            ' just add a row to the ListObject, the rest (shifting down existing CUD Marks and adding "i") is being taken care of the Application_SheetChange event procedure and the insertCUDMarks method
            insertTarget.ListObject.ListRows.Add(insertRow)
        End If
    End Sub

    ''' <summary>get the range from a worksheet name in the given sheet</summary>
    ''' <param name="theName">string name of range name</param>
    ''' <param name="theWs">given sheet</param>
    ''' <returns></returns>
    Public Function getRangeFromNameInSheet(ByRef theName As String, theWs As Excel.Worksheet) As Excel.Range
        For Each aName As Excel.Name In theWs.Names()
            If aName.Name = theWs.Name + "!" + theName Then
                getRangeFromNameInSheet = aName.RefersToRange
                Exit Function
            End If
        Next
        getRangeFromNameInSheet = Nothing
    End Function

End Module
