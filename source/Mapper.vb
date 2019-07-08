Imports ADODB
Imports Microsoft.Office.Interop.Excel

''' <summary>Contains Mapper function saveRangeToDB for storing/updating tabular excel data and some helper functions</summary>
Public Module Mapper
    ''' <summary>main db connection For mapper</summary>
    Public dbcnn As ADODB.Connection
    ''' <summary>should the Mapper instance be used interactively, thus giving error messages to the user
    ''' or used in automation mode (invoked by VBA, needs to be explicitly set), collecting error messages for later retrieval</summary>
    Public isInteractive As Boolean

    ''' <summary>dump data given in dataRange to a database table specified by tableName and connID
    ''' this database table can have multiple primary columns specified by primKeysStr
    ''' assumption: layout of dataRange is
    ''' primKey1Val,primKey2Val,..,primKeyNVal,DataCol1Val,DataCol2Val,..,DataColNVal</summary>
    ''' <param name="DataRange">Excel Range, where Data is taken from</param>
    Public Sub saveRangeToDB(DataRange As Object, DBMapperName As String)
        Dim tableName As String = String.Empty           ' Database Table, where Data is to be stored
        Dim primKeysStr As String = String.Empty         ' String containing primary Key names for updating table data, comma separated
        Dim database As String = String.Empty            ' Database to store to
        Dim env As Integer = DBAddin.selectedEnvironment + 1 ' Environment where connection id should be taken from (if not existing, take from selectedEnvironment)
        Dim insertIfMissing As Boolean = False           ' if set, then insert row into table if primary key is missing there. Default = False (only update)
        Dim executeAdditionalProc As String = ""         ' additional stored procedure to be executed after saving

        Dim rst As ADODB.Recordset
        Dim checkrst As ADODB.Recordset
        Dim checkTypes() As CheckTypeFld = Nothing
        Dim primKeys() As String
        Dim i As Integer, headerRow As Integer
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
        ' set up parameters
        Dim saveRangeParams() As String = functionSplit(DataRange.Cells(1, 1).Comment.Text, ",", """", "saveRangeToDB", "(", ")")
        If IsNothing(saveRangeParams) Then Exit Sub
        If saveRangeParams.Length < 4 Then
            LogError("At least environment (can be empty), Tablename, primary keys and database have to be provided as saveRangeToDB parameters !")
            Exit Sub
        End If
        If saveRangeParams(0) <> "" Then env = Convert.ToInt16(saveRangeParams(0))
        tableName = saveRangeParams(1).Replace("""", "").Trim ' remove all quotes and trim right and left
        If tableName = "" Then
            LogError("No Tablename given in DBMapper comment!")
            Exit Sub
        End If
        primKeysStr = saveRangeParams(2).Replace("""", "").Trim
        If primKeysStr = "" Then
            LogError("No primary keys given in DBMapper comment!")
            Exit Sub
        End If
        database = saveRangeParams(3).Replace("""", "").Trim
        If database = "" Then
            LogError("No database given in DBMapper comment!")
            Exit Sub
        End If
        If saveRangeParams.Length > 4 Then
            If saveRangeParams(4) <> "" Then insertIfMissing = Convert.ToBoolean(saveRangeParams(4))
        End If
        If saveRangeParams.Length > 5 Then
            If saveRangeParams(5) <> "" Then executeAdditionalProc = saveRangeParams(5).Replace("""", "").Trim
        End If
        On Error GoTo saveRangeToDB_Err
        primKeys = Split(primKeysStr, ",")

        'now create/get a connection (dbcnn) for env(ironment)
        LogInfo("saveRangeToDB Mapper: open connection...")
        If Not openConnection(env, database) Then Exit Sub

        'checkrst is opened to get information about table schema (field types)
        checkrst = New ADODB.Recordset
        rst = New ADODB.Recordset
        On Error Resume Next
        checkrst.Open(tableName, dbcnn, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly, CommandTypeEnum.adCmdTableDirect)

        If Err.Number <> 0 Then
            LogWarn("Table: " & tableName & " caused error: " & Err.Description & " in sheet " & DataRange.Parent.name)
            GoTo cleanup
        End If

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
        checkrst.Close()
        checkrst = Nothing

        headerRow = 1
        rowNum = headerRow + 1
        dbcnn.CursorLocation = CursorLocationEnum.adUseServer
        On Error GoTo saveRangeToDB_Err

        Dim finishLoop As Boolean
        ' walk through rows
        Do
            Dim primKeyCompound As String
            primKeyCompound = " WHERE "

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
            theHostApp.StatusBar = "Inserting/Updating data, " & primKeyCompound & " in table " & tableName

            On Error Resume Next
            rst.Open("SELECT * FROM " & tableName & primKeyCompound, dbcnn, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

            If Err.Number <> 0 Then
                LogWarn("Problem getting recordset, Error: " & Err.Description & " in sheet " & DataRange.Parent.name & ", & row " & rowNum)
                GoTo cleanup
            End If

            If rst.EOF Then
                If insertIfMissing Then
                    On Error GoTo saveRangeToDB_Err
                    rst.AddNew()
                    For i = 0 To UBound(primKeys)
                        rst.Fields(primKeys(i)).Value = IIf(DataRange.Cells(rowNum, i + 1).ToString().Length = 0, vbNull, DataRange.Cells(rowNum, i + 1).Value)
                    Next
                Else
                    DataRange.Parent.Activate
                    DataRange.Cells(rowNum, i + 1).Select
                    LogWarn("Problem getting recordset " & primKeyCompound & " from table '" & tableName & "', insertIfMissing = " & insertIfMissing & " in sheet " & DataRange.Parent.name & ", & row " & rowNum)
                    GoTo cleanup
                End If
            End If

            On Error GoTo saveRangeToDB_Err
            ' walk through columns and fill fields
            colNum = UBound(primKeys) + 1
            Do
                Dim fieldname As String
                fieldname = DataRange.Cells(headerRow, colNum).Value
                On Error Resume Next
                rst.Fields(fieldname).Value = IIf(DataRange.Cells(rowNum, colNum).ToString().Length = 0, vbNull, DataRange.Cells(rowNum, colNum).Value)

                If Err.Number <> 0 Then
                    DataRange.Parent.Activate
                    DataRange.Cells(rowNum, colNum).Select
                    LogWarn("Table: " & tableName & ", Field: " & fieldname & ", Error: " & Err.Description & " in sheet " & DataRange.Parent.name & ", & row " & rowNum & ", col: " & colNum)
                    GoTo cleanup
                End If
                On Error GoTo saveRangeToDB_Err
nextColumn:
                colNum += 1
            Loop Until colNum > DataRange.Columns.Count

            ' now do the update/insert in the DB
            On Error Resume Next
            rst.Update()
            If Err.Number <> 0 Then
                DataRange.Parent.Activate
                DataRange.Rows(rowNum).Select
                LogWarn("Table: " & rst.Source & ", Error: " & Err.Description & " in sheet " & DataRange.Parent.name & ", & row " & rowNum)
                GoTo cleanup
            End If
            rst.Close()
nextRow:
            On Error Resume Next
            finishLoop = IIf(DataRange.Cells(rowNum + 1, 1).ToString().Length = 0, True, False)

            If Err.Number <> 0 Then
                LogError("Error in primary column: Cells(" & rowNum + 1 & ",1)" & Err.Description)
                'finishLoop = True ' commented to allow erroneous data...
            End If
            Err.Clear()
            rowNum += 1
        Loop Until rowNum > DataRange.Rows.Count Or finishLoop

        If executeAdditionalProc.Length > 0 Then
            On Error Resume Next
            dbcnn.Execute(executeAdditionalProc)
            If Err.Number <> 0 Then
                LogError("Error in executing additional stored procedure:" & Err.Description)
                GoTo cleanup
            End If
        End If
        GoTo cleanup

saveRangeToDB_Err:
        LogError("Internal Error: " & Err.Description & ", line " & Erl() & " in Mapper.saveRangeToDB in sheet " & DataRange.Parent.name)
        On Error Resume Next
        rst.CancelBatch()
        rst.Close()
cleanup:
        rst = Nothing
        checkrst = Nothing
        dbcnn = Nothing
        theHostApp.StatusBar = False
    End Sub


    ''' <summary>check whether key with name "tblName" is contained in collection tblColl</summary>
    ''' <param name="tblName"></param>
    ''' <param name="tblColl"></param>
    ''' <returns>if name was found in collection</returns>
    Private Function existsInCollection(tblName As String, tblColl As Collection) As Boolean
        Dim dummy As Integer

        On Error GoTo err1
        existsInCollection = True
        dummy = tblColl(tblName)
        Exit Function
err1:
        Err.Clear()
        existsInCollection = False
    End Function

    ''' <summary>formats theVal to fit the type of column theHead having data type dataType</summary>
    ''' <param name="theVal"></param>
    ''' <param name="dataType"></param>
    ''' <returns>the formatted value</returns>
    Private Function dbFormatType(ByVal theVal As Object, dataType As CheckTypeFld) As String
        ' build where clause criteria..
        If dataType = CheckTypeFld.checkIsNumericFld Then
            dbFormatType = Replace(CStr(theVal), ",", ".")
        ElseIf dataType = CheckTypeFld.checkIsDateFld Then
            dbFormatType = "'" & Format$(theVal, "YYYYMMDD") & "'"
        ElseIf dataType = CheckTypeFld.checkIsTimeFld Then
            dbFormatType = "'" & Format$(theVal, "YYYYMMDD HH:MM:SS") & "'"
        ElseIf TypeName(theVal) = "Boolean" Then
            dbFormatType = IIf(theVal, 1, 0)
        ElseIf dataType = CheckTypeFld.checkIsStringFld Then
            ' quote strings
            theVal = Replace(theVal, "'", "''")
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
        Dim theConnString As String = fetchSetting("ConstConnString" & env, String.Empty)
        If theConnString = String.Empty Then
            LogError("No Connectionstring given for environment: " & env & ", please correct and rerun.")
            Exit Function
        End If

        On Error GoTo openConnection_Err
        openConnection = False
        dbcnn = Nothing
        dbcnn = New Connection
        Dim dbidentifier As String = fetchSetting("DBidentifierCCS" & env, String.Empty)
        If dbidentifier = String.Empty Then
            LogError("No DB identifier given for environment: " & env & ", please correct and rerun.")
            Exit Function
        End If
        dbcnn.ConnectionString = Change(theConnString, dbidentifier, database, ";")
        dbcnn.ConnectionTimeout = DBAddin.CnnTimeout
        dbcnn.CommandTimeout = DBAddin.CmdTimeout
        theHostApp.StatusBar = "Trying " & CnnTimeout & " sec. with connstring: " & theConnString
        On Error Resume Next
        dbcnn.Open()
        theHostApp.StatusBar = String.Empty
        If Err.Number <> 0 Then
            On Error GoTo openConnection_Err
            Dim exitMe As Boolean : exitMe = True
            LogWarn("openConnection: Error connecting to DB: " & Err.Description & ", connection string: " & theConnString, exitMe)
            dbcnn.Close()
            dbcnn = Nothing
        End If
        'dontTryConnection is only true if connection couldn't be succesfully opened until now..
        openConnection = Not dontTryConnection
        Exit Function

openConnection_Err:
        LogError("Error: " & Err.Description & ", line " & Erl() & " in Mapper.openConnection")
    End Function
End Module

Public Enum CheckTypeFld
    checkIsNumericFld = 0
    checkIsDateFld = 1
    checkIsTimeFld = 2
    checkIsStringFld = 3
End Enum
