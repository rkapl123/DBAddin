Imports ADODB
Imports Microsoft.Office.Interop.Excel

''' <summary>Contains Mapper function saveRangeToDB for storing/updating tabular excel data and some helper functions</summary>
Public Module Mapper

    ''' <summary>main db connection For mapper</summary>
    Public dbcnn As ADODB.Connection

    ''' <summary>dump data given in dataRange to a database table specified by tableName and connID
    ''' this database table can have multiple primary columns specified by primKeysStr
    ''' assumption: layout of dataRange is
    ''' primKey1Val,primKey2Val,..,primKeyNVal,DataCol1Val,DataCol2Val,..,DataColNVal</summary>
    ''' <param name="DataRange">Excel Range, where Data is taken from</param>
    Public Sub saveRangeToDB(DataRange As Object, DBMapperName As String, Optional WbIsSaving As Boolean = False)
        Dim tableName As String = ""                         ' Database Table, where Data is to be stored
        Dim primKeysStr As String = ""                       ' String containing primary Key names for updating table data, comma separated
        Dim database As String = ""                          ' Database to store to
        Dim env As Integer = DBAddin.selectedEnvironment + 1 ' Environment where connection id should be taken from (if not existing, take from selectedEnvironment)
        Dim insertIfMissing As Boolean = False               ' if set, then insert row into table if primary key is missing there. Default = False (only update)
        Dim executeAdditionalProc As String = ""             ' additional stored procedure to be executed after saving
        Dim ignoreColumns As String = ""                     ' columns to be ignored (helper columns)
        Dim storeDBMapOnSave As Boolean = False              ' should DBMap be saved on Excel Saving? (default no)

        Dim rst As ADODB.Recordset
        Dim checkrst As ADODB.Recordset
        Dim checkTypes() As CheckTypeFld = Nothing
        Dim primKeys() As String

        Dim i As Integer
        Dim rowNum As Long, colNum As Long

        ' extend DataRange if it is only one cell ...
        If DataRange.Rows.Count = 1 And DataRange.Columns.Count = 1 Then
            Dim rowEnd = DataRange.End(XlDirection.xlDown).Row
            Dim colEnd = DataRange.End(XlDirection.xlToRight).Column
            DataRange = DataRange.Parent.Range(DataRange, DataRange.Parent.Cells(rowEnd, colEnd))
        End If
        If IsNothing(DataRange.Cells(1, 1).Comment.Text) Then
            LogError("No definition comment found for DBMapper definition in " + DBMapperName)
        End If

        ' set up parameters from comment text
        If Not getParametersFromText(DataRange.Cells(1, 1).Comment.Text, env, tableName, primKeysStr, database, insertIfMissing, executeAdditionalProc, ignoreColumns, storeDBMapOnSave) Then Exit Sub
        If WbIsSaving And Not storeDBMapOnSave Then Exit Sub
        primKeys = Split(primKeysStr, ",")
        ignoreColumns = LCase(ignoreColumns) + "," ' lowercase and add comma for better retrieval

        'now create/get a connection (dbcnn) for env(ironment)
        LogInfo("saveRangeToDB Mapper: open connection...")
        If Not openConnection(env, database) Then Exit Sub

        'checkrst is opened to get information about table schema (field types)
        checkrst = New ADODB.Recordset
        rst = New ADODB.Recordset
        Try
            checkrst.Open(tableName, dbcnn, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly, CommandTypeEnum.adCmdTableDirect)
        Catch ex As Exception
            LogWarn("Table: " & tableName & " caused error: " & ex.Message & " in sheet " & DataRange.Parent.name)
            checkrst.Close()
            GoTo cleanup
        End Try

        ' to find the record to be updated, get types for primKeyCompound to build WHERE Clause with it
        For i = 0 To UBound(primKeys)
            ReDim Preserve checkTypes(i)

            If checkIsNumeric(checkrst.Fields(primKeys(i)).Type) Then
                checkTypes(i) = CheckTypeFld.checkIsNumericFld
            ElseIf checkIsDate(checkrst.Fields(primKeys(i)).Type) Then
                checkTypes(i) = CheckTypeFld.checkIsDateFld
            ElseIf checkIsTime(checkrst.Fields(primKeys(i)).Type) Then
                checkTypes(i) = CheckTypeFld.checkIsTimeFld
            Else
                checkTypes(i) = CheckTypeFld.checkIsStringFld
            End If
        Next
        ' check if all column names (except ignored) of DBMapper Range exist in table
        colNum = 1
        Do
            Dim fieldname As String = DataRange.Cells(1, colNum).Value
            ' only if not ignored...
            If InStr(1, LCase(fieldname) + ",", ignoreColumns) = 0 Then
                Try
                    Dim testExist As String = checkrst.Fields(fieldname).Name
                Catch ex As Exception
                    DataRange.Parent.Activate
                    DataRange.Cells(1, colNum).Select
                    LogWarn("Field '" & fieldname & "' does not exist in Table '" & tableName & "' and is not in ignoreColumns, Error in sheet " & DataRange.Parent.name)
                    GoTo cleanup
                End Try
            End If
            colNum += 1
        Loop Until colNum > DataRange.Columns.Count
        checkrst.Close()

        rowNum = 2
        dbcnn.CursorLocation = CursorLocationEnum.adUseServer

        Dim finishLoop As Boolean
        ' walk through rows
        Do
            Dim primKeyCompound As String = " WHERE "

            For i = 0 To UBound(primKeys)
                Dim primKeyValue
                primKeyValue = DataRange.Cells(rowNum, i + 1).Value
                primKeyCompound = primKeyCompound & primKeys(i) & " = " & dbFormatType(primKeyValue, checkTypes(i)) & IIf(i = UBound(primKeys), "", " AND ")
                If IsError(primKeyValue) Then
                    LogError("Error in primary key value, cell (" & rowNum & "," & i + 1 & ") in sheet " & DataRange.Parent.name & ", & row " & rowNum)
                    GoTo nextRow
                End If
                If primKeyValue.ToString().Length = 0 Then
                    LogError("Empty primary key value, cell (" & rowNum & "," & i + 1 & ") in sheet " & DataRange.Parent.name & ", & row " & rowNum)
                    GoTo nextRow
                End If
            Next
            hostApp.StatusBar = "Inserting/Updating data, " & primKeyCompound & " in table " & tableName

            Try
                rst.Open("SELECT * FROM " & tableName & primKeyCompound, dbcnn, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)
            Catch ex As Exception
                LogWarn("Problem getting recordset, Error: " & ex.Message & " in sheet " & DataRange.Parent.name & ", & row " & rowNum)
                rst.Close()
                GoTo cleanup
            End Try

            If rst.EOF Then
                If insertIfMissing Then
                    rst.AddNew()
                    For i = 0 To UBound(primKeys)
                        rst.Fields(primKeys(i)).Value = IIf(DataRange.Cells(rowNum, i + 1).ToString().Length = 0, vbNull, DataRange.Cells(rowNum, i + 1).Value)
                    Next
                Else
                    DataRange.Parent.Activate
                    DataRange.Cells(rowNum, i + 1).Select
                    LogWarn("Problem getting recordset " & primKeyCompound & " from table '" & tableName & "', insertIfMissing = " & insertIfMissing & " in sheet " & DataRange.Parent.name & ", & row " & rowNum)
                    rst.Close()
                    GoTo cleanup
                End If
            End If

            ' walk through columns and fill fields
            colNum = UBound(primKeys) + 1
            Do
                Dim fieldname As String = DataRange.Cells(1, colNum).Value
                If InStr(1, ignoreColumns, LCase(fieldname) + ",") = 0 Then
                    Try
                        rst.Fields(fieldname).Value = IIf(DataRange.Cells(rowNum, colNum).ToString().Length = 0, vbNull, DataRange.Cells(rowNum, colNum).Value)
                    Catch ex As Exception
                        DataRange.Parent.Activate
                        DataRange.Cells(rowNum, colNum).Select
                        LogError("General Error: " & ex.Message & " with Table: " & tableName & ", Field: " & fieldname & ", in sheet " & DataRange.Parent.name & ", & row " & rowNum & ", col: " & colNum)
                        rst.Close()
                        GoTo cleanup
                    End Try
                End If
                colNum += 1
            Loop Until colNum > DataRange.Columns.Count

            ' now do the update/insert in the DB
            Try
                rst.Update()
            Catch ex As Exception
                DataRange.Parent.Activate
                DataRange.Rows(rowNum).Select
                LogWarn("Table: " & rst.Source & ", Error: " & ex.Message & " in sheet " & DataRange.Parent.name & ", & row " & rowNum)
                rst.Close()
                GoTo cleanup
            End Try
            rst.Close()
nextRow:
            Try
                finishLoop = IIf(DataRange.Cells(rowNum + 1, 1).ToString().Length = 0, True, False)
            Catch ex As Exception
                LogError("Error in primary column: Cells(" & rowNum + 1 & ",1)" & ex.Message)
                'finishLoop = True ' commented to allow erroneous data...
            End Try
            rowNum += 1
        Loop Until rowNum > DataRange.Rows.Count Or finishLoop

        If executeAdditionalProc.Length > 0 Then
            Try
                dbcnn.Execute(executeAdditionalProc)
            Catch ex As Exception
                LogError("Error in executing additional stored procedure:" & ex.Message)
                GoTo cleanup
            End Try
        End If
cleanup:
        dbcnn = Nothing
        hostApp.StatusBar = False
    End Sub

    ''' <summary>formats theVal to fit the type of column theHead having data type dataType</summary>
    ''' <param name="theVal"></param>
    ''' <param name="dataType"></param>
    ''' <returns>the formatted value</returns>
    Private Function dbFormatType(ByVal theVal As Object, dataType As CheckTypeFld) As String

        If dataType = CheckTypeFld.checkIsNumericFld Then ' only decimal points allowed in numeric data
            dbFormatType = Replace(CStr(theVal), ",", ".")
        ElseIf dataType = CheckTypeFld.checkIsDateFld Then
            dbFormatType = "'" & Format$(theVal, "YYYYMMDD") & "'" ' standard SQL Date formatting
        ElseIf dataType = CheckTypeFld.checkIsTimeFld Then
            dbFormatType = "'" & Format$(theVal, "YYYYMMDD HH:MM:SS") & "'" ' standard SQL Date/time formatting
        ElseIf TypeName(theVal) = "Boolean" Then
            dbFormatType = IIf(theVal, 1, 0)
        ElseIf dataType = CheckTypeFld.checkIsStringFld Then ' quote Strings
            theVal = Replace(theVal, "'", "''") ' quote quotes inside Strings
            dbFormatType = "'" & theVal & "'"
        Else
            LogError("Error: unknown data type '" & dataType & "' given in Mapper.dbFormatType !!")
            dbFormatType = String.Empty
        End If
    End Function

    ''' <summary>opens a database connection</summary>
    ''' <param name="env">number of the environment as given in the settings</param>
    ''' <param name="database">database to replace database selection parameter in connection string of environment</param>
    ''' <returns>True on success</returns>
    Public Function openConnection(env As Integer, database As String) As Boolean
        openConnection = False

        Dim theConnString As String = fetchSetting("ConstConnString" & env, String.Empty)
        If theConnString = String.Empty Then
            LogError("No Connectionstring given for environment: " & env & ", please correct and rerun.")
            Exit Function
        End If
        Dim dbidentifier As String = fetchSetting("DBidentifierCCS" & env, String.Empty)
        If dbidentifier = String.Empty Then
            LogError("No DB identifier given for environment: " & env & ", please correct and rerun.")
            Exit Function
        End If
        dbcnn = New Connection
        theConnString = Change(theConnString, dbidentifier, database, ";")
        hostApp.StatusBar = "Trying " & DBAddin.CnnTimeout & " sec. with connstring: " & theConnString
        Try
            dbcnn.ConnectionString = theConnString
            dbcnn.ConnectionTimeout = DBAddin.CnnTimeout
            dbcnn.CommandTimeout = DBAddin.CmdTimeout
            dbcnn.Open()
        Catch ex As Exception
            Dim exitMe As Boolean : exitMe = True
            LogWarn("openConnection: Error connecting to DB: " & Err.Description & ", connection string: " & theConnString, exitMe)
            If dbcnn.State = ADODB.ObjectStateEnum.adStateOpen Then dbcnn.Close()
            dbcnn = Nothing
        End Try
        hostApp.StatusBar = String.Empty
        openConnection = True
    End Function

    ''' <summary>gets defined named ranges for DBMapper invocation in the current workbook and updates Ribbon with it</summary>
    Sub getDBMapperDefinitions()
        ' load DBMapper definitions
        DBAddin.DBMapperDefColl = New Dictionary(Of String, Dictionary(Of String, Range))
        For Each namedrange As Name In hostApp.ActiveWorkbook.Names
            Dim cleanname As String = Replace(namedrange.Name, namedrange.Parent.Name & "!", "")
            If Left(cleanname, 8) = "DBMapper" Then
                If InStr(namedrange.RefersTo, "#REF!") > 0 Then LogError("DBMapper definitions range " + namedrange.Parent.name + "!" + namedrange.Name + " contains #REF!")
                Dim nodeName As String = Replace(Replace(namedrange.Name, "DBMapper", ""), namedrange.Parent.Name & "!", "")
                If nodeName = "" Then nodeName = "UnnamedDBMapper"

                Dim i As Integer = namedrange.RefersToRange.Parent.Index
                Dim defColl As Dictionary(Of String, Range)
                If Not DBMapperDefColl.ContainsKey("ID" + i.ToString()) Then
                    ' add to new sheet "menu"
                    defColl = New Dictionary(Of String, Range)
                    defColl.Add(nodeName, namedrange.RefersToRange)
                    DBMapperDefColl.Add("ID" + i.ToString(), defColl)
                Else
                    ' add definition to existing sheet "menu"
                    defColl = DBMapperDefColl("ID" + i.ToString())
                    defColl.Add(nodeName, namedrange.RefersToRange)
                End If
            End If
            If DBMapperDefColl.Count >= 15 Then LogError("Not more than 15 sheets with DBMapper definitions possible, ignoring definitions in sheet " + namedrange.Parent.Name)
        Next
        DBAddin.theRibbon.Invalidate()
    End Sub

    ''' <summary>saves defined DBMaps (depending on configuration)</summary>
    Sub saveDBMaps(Wb As Workbook)
        ' save all DBmaps on saving except Readonly is recommended on this workbook
        Dim DBmapSheet As String
        If Not Wb.ReadOnlyRecommended Then
            For Each DBmapSheet In DBMapperDefColl.Keys
                For Each dbmapdefkey In DBMapperDefColl(DBmapSheet).Keys
                    saveRangeToDB(DBMapperDefColl(DBmapSheet).Item(dbmapdefkey), dbmapdefkey, True)
                Next
            Next
        End If
    End Sub

    ''' <summary>get parameters from passed paramText</summary>
    ''' <param name="paramText">paramtext to be parsed, currently only saveRangeToDB params</param>
    ''' <param name="env">Environment (integer) where connection id should be taken from (if not existing, take from selectedEnvironment)</param>
    ''' <param name="tableName">Database Table, where Data is to be stored</param>
    ''' <param name="primKeysStr">String containing primary Key names for updating table data, comma separated</param>
    ''' <param name="database">Database to store to</param>
    ''' <param name="insertIfMissing">if set, then insert row into table if primary key is missing there. Default = False (only update)</param>
    ''' <param name="executeAdditionalProc">additional stored procedure to be executed after saving</param>
    ''' <param name="ignoreColumns">columns to be ignored (helper columns)</param>
    ''' <param name="storeDBMapOnSave">should DBMap be saved on Excel Saving? (default no)</param>
    Function getParametersFromText(paramText As String, ByRef env As Integer, ByRef tableName As String, ByRef primKeysStr As String, ByRef database As String, ByRef insertIfMissing As Boolean, ByRef executeAdditionalProc As String, ByRef ignoreColumns As String, ByRef storeDBMapOnSave As Boolean) As Boolean
        Dim saveRangeParams() As String = functionSplit(paramText, ",", """", "saveRangeToDB", "(", ")")
        If IsNothing(saveRangeParams) Then Return False
        If saveRangeParams.Length < 4 Then
            LogError("At least environment (can be empty), Tablename, primary keys and database have to be provided as saveRangeToDB parameters !")
            Return False
        End If
        If saveRangeParams(0) <> "" Then env = Convert.ToInt16(saveRangeParams(0))
        tableName = saveRangeParams(1).Replace("""", "").Trim ' remove all quotes and trim right and left
        If tableName = "" Then
            LogError("No Tablename given in DBMapper comment!")
            Return False
        End If
        primKeysStr = saveRangeParams(2).Replace("""", "").Trim
        If primKeysStr = "" Then
            LogError("No primary keys given in DBMapper comment!")
            Return False
        End If
        database = saveRangeParams(3).Replace("""", "").Trim
        If database = "" Then
            LogError("No database given in DBMapper comment!")
            Return False
        End If
        insertIfMissing = False
        If saveRangeParams.Length > 4 Then
            If saveRangeParams(4) <> "" Then insertIfMissing = Convert.ToBoolean(saveRangeParams(4))
        End If
        If saveRangeParams.Length > 5 Then
            If saveRangeParams(5) <> "" Then executeAdditionalProc = saveRangeParams(5).Replace("""", "").Trim
        End If
        If saveRangeParams.Length > 6 Then
            If saveRangeParams(6) <> "" Then ignoreColumns = saveRangeParams(6).Replace("""", "").Trim
        End If
        storeDBMapOnSave = False
        If saveRangeParams.Length > 7 Then
            If saveRangeParams(7) <> "" Then storeDBMapOnSave = Convert.ToBoolean(saveRangeParams(7))
        End If
        Return True
    End Function

    ''' <summary>creates a DBMap at the current active cell</summary>
    Sub createDBMapper()
        Dim theDBMapperCreateDlg As DBMapperCreate = Nothing
        Dim activeCellName As String = ""
        ' try regular defined name
        Try : activeCellName = hostApp.Selection.Name.Name : Catch ex As Exception : End Try
        ' try potential name to ListObject (parts), only possible if manually defined !
        If activeCellName = "" And Not IsNothing(hostApp.Selection.ListObject) Then
            For Each listObjectCol In hostApp.Selection.ListObject.ListColumns
                Dim aName As Name
                For Each aName In hostApp.ActiveWorkbook.Names
                    If aName.RefersTo = "=" & hostApp.Selection.ListObject.Name & "[[#Headers],[" & hostApp.Selection.Value & "]]" Then
                        activeCellName = aName.Name
                        Exit For
                    End If
                Next
                If activeCellName <> "" Then Exit For
            Next
        End If
        If Not IsNothing(hostApp.ActiveCell.Comment) Then
            ' fetch parameters if existing comment and DBMapper definition...
            Dim tableName As String = ""                         ' Database Table, where Data is to be stored
            Dim primKeysStr As String = ""                       ' String containing primary Key names for updating table data, comma separated
            Dim database As String = ""                          ' Database to store to
            Dim env As Integer = -1                              ' Environment where connection id should be taken from
            Dim insertIfMissing As Boolean = False               ' if set, then insert row into table if primary key is missing there. Default = False (only update)
            Dim executeAdditionalProc As String = ""             ' additional stored procedure to be executed after saving
            Dim ignoreColumns As String = ""                     ' columns to be ignored (helper columns)
            Dim storeDBMapOnSave As Boolean = False              ' should DBMap be saved on Excel Saving? (default no)

            If Not getParametersFromText(hostApp.ActiveCell.Comment.Text, env, tableName, primKeysStr, database, insertIfMissing, executeAdditionalProc, ignoreColumns, storeDBMapOnSave) Then Exit Sub
            theDBMapperCreateDlg = New DBMapperCreate()
            If InStr(1, activeCellName, "DBMapper") > 0 Then theDBMapperCreateDlg.DBMapperName.Text = Replace(activeCellName, "DBMapper", "")
            theDBMapperCreateDlg.envSel.DataSource = environdefs
            theDBMapperCreateDlg.envSel.SelectedIndex = env
            theDBMapperCreateDlg.Tablename.Text = tableName
            theDBMapperCreateDlg.PrimaryKeys.Text = primKeysStr
            theDBMapperCreateDlg.Database.Text = database
            theDBMapperCreateDlg.insertIfMissing.Checked = insertIfMissing
            theDBMapperCreateDlg.addStoredProc.Text = executeAdditionalProc
            theDBMapperCreateDlg.IgnoreColumns.Text = ignoreColumns
            theDBMapperCreateDlg.storeDBMapOnSave.Checked = storeDBMapOnSave
        End If
        ' no DBMapper definitions found...
        If IsNothing(theDBMapperCreateDlg) Then
            theDBMapperCreateDlg = New DBMapperCreate()
            theDBMapperCreateDlg.envSel.DataSource = environdefs
            theDBMapperCreateDlg.envSel.SelectedIndex = -1
        End If

        ' display dialog for parameters
        If theDBMapperCreateDlg.ShowDialog() = System.Windows.Forms.DialogResult.Cancel Then Exit Sub
        ' set name
        If InStr(1, activeCellName, "DBMapper") > 0 Then   ' fetch parameters if existing comment and DBMapper definition...
            Try : hostApp.ActiveWorkbook.Names(activeCellName).Delete
            Catch ex As Exception : LogError("Error when removing name '" + activeCellName + "' from active cell: " & ex.Message)
            End Try
        End If
        Try : hostApp.ActiveCell.Name = "DBMapper" + theDBMapperCreateDlg.DBMapperName.Text
        Catch ex As Exception : LogError("Error when assigning name 'DBMapper" + theDBMapperCreateDlg.DBMapperName.Text + "' to active cell: " & ex.Message)
        End Try
        ' set parameters in comment text
        Try : hostApp.ActiveCell.ClearComments() : Catch ex As Exception : End Try
        Dim paramText As String = "saveRangeToDB(" +
            IIf(theDBMapperCreateDlg.envSel.SelectedIndex = -1, "", theDBMapperCreateDlg.envSel.SelectedIndex.ToString()) + ",""" +
            theDBMapperCreateDlg.Tablename.Text + """,""" +
            theDBMapperCreateDlg.PrimaryKeys.Text + """,""" +
            theDBMapperCreateDlg.Database.Text + """," +
            theDBMapperCreateDlg.insertIfMissing.Checked.ToString() + ",""" +
            IIf(Len(theDBMapperCreateDlg.addStoredProc.Text) = 0, "", theDBMapperCreateDlg.addStoredProc.Text) + """,""" +
            IIf(Len(theDBMapperCreateDlg.IgnoreColumns.Text) = 0, "", theDBMapperCreateDlg.IgnoreColumns.Text) + """," +
            theDBMapperCreateDlg.storeDBMapOnSave.Checked.ToString() + ")"
        Try
            hostApp.ActiveCell.AddComment()
            hostApp.ActiveCell.Comment.Text(Text:=paramText)
            hostApp.ActiveCell.Comment.Shape.TextFrame.Characters.Font.Bold = False
        Catch ex As Exception : LogError("Error when adding comments with DBMapper parameters to active cell: " & ex.Message) : End Try
        ' refresh mapper definitions to reflect changes immediately...
        getDBMapperDefinitions()
    End Sub

End Module

Public Enum CheckTypeFld
    checkIsNumericFld = 0
    checkIsDateFld = 1
    checkIsTimeFld = 2
    checkIsStringFld = 3
End Enum
