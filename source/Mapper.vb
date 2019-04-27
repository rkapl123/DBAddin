Imports Microsoft.Office.Interop.Excel
Imports ADODB

''' <summary>Contains the public callable Mapper function for storing/updating tabular excel data</summary>
Public Enum checkTypeFld
    checkIsNumericFld = 0
    checkIsDateFld = 1
    checkIsTimeFld = 2
    checkIsStringFld = 3
End Enum

Public Class Mapper
    ''' <summary>main db connection For mapper</summary>
    Public dbcnn As ADODB.Connection
    ''' <summary>should the Mapper instance be used interactively, thus giving error messages to the user
    ''' or used in automation mode (invoked by VBA, needs to be explicitly set), collecting error messages for later retrieval</summary>
    Public isInteractive As Boolean
    ''' <summary>if Mapper instance is not interactive, Log procedures store all messages here...</summary>
    Public returnedErrorMessages As String

    ''' <summary>dump data given in dataRange to a database table specified by tableName and connID
    ''' this database table can have multiple primary columns specified by primKeysStr
    ''' assumption: layout of dataRange is
    ''' primKey1Val,primKey2Val,..,primKeyNVal,DataCol1Val,DataCol2Val,..,DataColNVal</summary>
    ''' <param name="DataRange">Excel Range, where Data is taken from</param>
    ''' <param name="tableName">Database Table, where Data is to be stored</param>
    ''' <param name="primKeysStr">String containing primary Key names for updating table data, comma separated</param>
    ''' <param name="database">Database to replace D</param>
    ''' <param name="env">Environment where connection id should be taken from</param>
    ''' <param name="insertIfMissing">then insert row into table if primary key is missing there. Default = False (only update)</param>
    ''' <param name="executeAdditionalProc">additional stored procedure to be executed after saving</param>
    ''' <returns>True if successful, false in case of errors.</returns>
    Public Function saveRangeToDB(DataRange As Range,
                                    tableName As String,
                                    primKeysStr As String,
                                    database As String,
                                    Optional env As Integer = 1,
                                    Optional insertIfMissing As Boolean = False,
                                    Optional executeAdditionalProc As String = "") As Boolean
        Dim rst As ADODB.Recordset
        Dim checkrst As ADODB.Recordset
        Dim checkTypes() As checkTypeFld = Nothing
        Dim primKeys() As String
        Dim i As Integer, headerRow As Integer
        Dim rowNum As Long, colNum As Long

        ' set up parameters
        On Error GoTo saveRangeToDB_Err
        If Not isInteractive Then
            automatedMapper = Me
            Me.returnedErrorMessages = String.Empty
        End If
        saveRangeToDB = False
        primKeys = Split(primKeysStr, ",")

        ' first, create/get a connection (dbcnn)
        LogInfo("saveRangeToDB Mapper: open connection...")

        'now create/get a connection (dbcnn) for env(ironment)
        If Not openConnection(env, database) Then Exit Function

        'checkrst is opened to get information about table schema (field types)
        checkrst = New ADODB.Recordset
        rst = New ADODB.Recordset
        On Error Resume Next
        checkrst.Open(tableName, dbcnn, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly, CommandTypeEnum.adCmdTableDirect)

        If Err.Number <> 0 Then
            LogWarn("Table: " & tableName & " caused error: " & Err.Description & " in sheet " & DataRange.Parent.name)
            GoTo cleanup
        End If

        For i = 0 To UBound(primKeys)
            ReDim Preserve checkTypes(i)

            If checkIsNumeric(checkrst.Fields(primKeys(i)).Type) Then
                checkTypes(i) = checkTypeFld.checkIsNumericFld
            ElseIf checkIsDate(checkrst.Fields(primKeys(i)).Type) Then
                checkTypes(i) = checkTypeFld.checkIsDateFld
            ElseIf checkIsTime(checkrst.Fields(primKeys(i)).Type) Then
                checkTypes(i) = checkTypeFld.checkIsTimeFld
            Else
                checkTypes(i) = checkTypeFld.checkIsStringFld
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
                        rst.Fields(primKeys(i)).Value = IIf(DataRange.Cells(rowNum, i + 1).ToString.Length = 0, vbNull, DataRange.Cells(rowNum, i + 1).Value)
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
                rst.Fields(fieldname).Value = IIf(DataRange.Cells(rowNum, colNum).ToString.Length = 0, vbNull, DataRange.Cells(rowNum, colNum).Value)

                If Err.Number <> 0 Then
                    DataRange.Parent.Activate
                    DataRange.Cells(rowNum, colNum).Select
                    LogWarn("Table: " & tableName & ", Field: " & fieldname & ", Error: " & Err.Description & " in sheet " & DataRange.Parent.name & ", & row " & rowNum & ", col: " & colNum)
                    GoTo cleanup
                End If
                On Error GoTo saveRangeToDB_Err
nextColumn:
                colNum = colNum + 1
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
            finishLoop = IIf(DataRange.Cells(rowNum + 1, 1).ToString.Length = 0, True, False)

            If Err.Number <> 0 Then
                LogError("Error in primary column: Cells(" & rowNum + 1 & ",1)" & Err.Description)
                'finishLoop = True ' commented to allow erroneous data...
            End If
            Err.Clear()
            rowNum = rowNum + 1
        Loop Until rowNum > DataRange.Rows.Count Or finishLoop

        If executeAdditionalProc.Length > 0 Then
            On Error Resume Next
            dbcnn.Execute(executeAdditionalProc)
            If Err.Number <> 0 Then
                LogError("Error in executing additional stored procedure:" & Err.Description)
                GoTo cleanup
            End If
        End If
        saveRangeToDB = True
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
    End Function


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
    Private Function dbFormatType(ByVal theVal As Object, dataType As checkTypeFld) As String
        ' build where clause criteria..
        If dataType = checkTypeFld.checkIsNumericFld Then
            dbFormatType = Replace(CStr(theVal), ",", ".")
        ElseIf dataType = checkTypeFld.checkIsDateFld Then
            dbFormatType = "'" & Format$(theVal, "YYYYMMDD") & "'"
        ElseIf dataType = checkTypeFld.checkIsTimeFld Then
            dbFormatType = "'" & Format$(theVal, "YYYYMMDD HH:MM:SS") & "'"
        ElseIf TypeName(theVal) = "Boolean" Then
            dbFormatType = IIf(theVal, 1, 0)
        ElseIf dataType = checkTypeFld.checkIsStringFld Then
            ' quote strings
            theVal = Replace(theVal, "'", "''")
            dbFormatType = "'" & theVal & "'"
        Else
            LogError("Error: unknown data type '" & dataType & "' given in Mapper.dbFormatType !!")
            dbFormatType = String.Empty
        End If
    End Function

    ''' <summary>new Mapper instances are always interactive by default. If you want automation mode, set isInteractive to false before calling saveRangeToDB</summary>
    Public Sub New()
        isInteractive = True
    End Sub

    ''' <summary>opens a database connection</summary>
    ''' <param name="env">number of the environment as given in the settings</param>
    ''' <param name="database">database to replace database selection parameter in connection string of environment</param>
    ''' <returns>True on success</returns>
    Public Function openConnection(env As Integer, database As String) As Boolean
        Dim theConnString As String = fetchSetting("ConstConnString" & env, String.Empty)

        On Error GoTo openConnection_Err
        openConnection = False
        If Not dbcnn Is Nothing Then
            If dbcnn.State = ObjectStateEnum.adStateClosed Then dbcnn = Nothing
        End If
        If dbcnn Is Nothing Then
            dbcnn = New Connection
            dbcnn.ConnectionString = Change(theConnString, fetchSetting("dbspecifier", "database="), database, ";")
            dbcnn.ConnectionTimeout = CnnTimeout
            dbcnn.CommandTimeout = CmdTimeout
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
        End If
        'dontTryConnection is only true if connection couldn't be succesfully opened until now..
        openConnection = Not dontTryConnection
        Exit Function

openConnection_Err:
        LogError("Error: " & Err.Description & ", line " & Erl() & " in Mapper.openConnection")
    End Function

End Class
