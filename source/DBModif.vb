Imports ADODB
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Collections.Generic

Friend Enum CheckTypeFld
    checkIsNumericFld = 0
    checkIsDateFld = 1
    checkIsTimeFld = 2
    checkIsStringFld = 3
End Enum

Public MustInherit Class DBModif

    '''<summary>unique key of DBModif</summary>
    Protected dbmapdefkey As String
    '''<summary>sheet where DBModif is defined (only DBMapper and DBAction)</summary>
    Protected DBmapSheet As String
    '''<summary>address of targetRange (only DBMapper and DBAction)</summary>
    Protected targetRangeAddress As String
    ''' <summary>Range where DBMapper data is located (only DBMapper and DBAction; paramText is stored in custom doc properties having the same Name)</summary>
    Protected TargetRange As Excel.Range
    '''<summary>parameter text for DBModif (def(...)</summary>
    Protected paramText As String
    '''<summary>should DBMap be saved / DBAction be done on Excel Saving? (default no)</summary>
    Protected execOnSave As Boolean = False

    Public Function getTargetRangeAddress() As String
        getTargetRangeAddress = targetRangeAddress
    End Function

    Public Function getTargetRange() As Excel.Range
        getTargetRange = TargetRange
    End Function

    Public  Function getParamText() As String
        getParamText = paramText
    End Function

    Public Overridable Sub doDBModif(Optional WbIsSaving As Boolean = False, Optional calledByDBSeq As String = "")
        Throw New NotImplementedException()
    End Sub

    ''' <summary>is saving needed for this DBModifier</summary>
    ''' <returns>true for saving necessary</returns>
    Public Overridable Function DBModifSaveNeeded() As Boolean
        Throw New NotImplementedException()
    End Function

    Public Overridable Sub setDBModifCreateFields(ByRef theDBModifCreateDlg As DBModifCreate)
        Throw New NotImplementedException()
    End Sub

    ''' <summary>formats theVal to fit the type of column theHead having data type dataType</summary>
    ''' <param name="theVal"></param>
    ''' <param name="dataType"></param>
    ''' <returns>the formatted value</returns>
    Friend Function dbFormatType(ByVal theVal As Object, dataType As CheckTypeFld) As String

        If dataType = CheckTypeFld.checkIsNumericFld Then ' only decimal points allowed in numeric data
            dbFormatType = Replace(CStr(theVal), ",", ".")
        ElseIf dataType = CheckTypeFld.checkIsDateFld Then
            dbFormatType = "'" & Format(theVal, "yyyyMMdd") & "'" ' standard SQL Date formatting
        ElseIf dataType = CheckTypeFld.checkIsTimeFld Then
            dbFormatType = "'" & Format(theVal, "yyyyMMdd hh:mm:ss") & "'" ' standard SQL Date/time formatting
        ElseIf TypeName(theVal) = "Boolean" Then
            dbFormatType = IIf(theVal, 1, 0)
        ElseIf dataType = CheckTypeFld.checkIsStringFld Then ' quote Strings
            theVal = Replace(theVal, "'", "''") ' quote quotes inside Strings
            dbFormatType = "'" & theVal & "'"
        Else
            ErrorMsg("Error: unknown data type '" & dataType)
            dbFormatType = String.Empty
        End If
    End Function

    ''' <summary>checks whether ADO type theType is a date or time type</summary>
    ''' <param name="theType"></param>
    ''' <returns>True if DateTime</returns>
    Friend Function checkIsDateTime(theType As ADODB.DataTypeEnum) As Boolean
        checkIsDateTime = False
        If theType = ADODB.DataTypeEnum.adDate Or theType = ADODB.DataTypeEnum.adDBDate Or theType = ADODB.DataTypeEnum.adDBTime Or theType = ADODB.DataTypeEnum.adDBTimeStamp Then
            checkIsDateTime = True
        End If
    End Function

    ''' <summary>checks whether ADO type theType is a date type</summary>
    ''' <param name="theType"></param>
    ''' <returns>True if Date</returns>
    Friend Function checkIsDate(theType As ADODB.DataTypeEnum) As Boolean
        checkIsDate = False
        If theType = ADODB.DataTypeEnum.adDate Or theType = ADODB.DataTypeEnum.adDBDate Then
            checkIsDate = True
        End If
    End Function

    ''' <summary>checks whether ADO type theType is a time type</summary>
    ''' <param name="theType"></param>
    ''' <returns>True if Time</returns>
    Friend Function checkIsTime(theType As ADODB.DataTypeEnum) As Boolean
        checkIsTime = False
        If theType = ADODB.DataTypeEnum.adDBTime Or theType = ADODB.DataTypeEnum.adDBTimeStamp Then
            checkIsTime = True
        End If
    End Function

    ''' <summary>checks whether ADO type theType is a numeric type</summary>
    ''' <param name="theType"></param>
    ''' <returns>True if numeric</returns>
    Friend Function checkIsNumeric(theType As ADODB.DataTypeEnum) As Boolean
        checkIsNumeric = False
        If theType = ADODB.DataTypeEnum.adNumeric Or theType = ADODB.DataTypeEnum.adInteger Or theType = ADODB.DataTypeEnum.adTinyInt Or theType = ADODB.DataTypeEnum.adSmallInt Or theType = ADODB.DataTypeEnum.adBigInt Or theType = ADODB.DataTypeEnum.adUnsignedInt Or theType = ADODB.DataTypeEnum.adUnsignedTinyInt Or theType = ADODB.DataTypeEnum.adUnsignedSmallInt Or theType = ADODB.DataTypeEnum.adDouble Or theType = ADODB.DataTypeEnum.adSingle Or theType = ADODB.DataTypeEnum.adCurrency Or theType = ADODB.DataTypeEnum.adUnsignedBigInt Then
            checkIsNumeric = True
        End If
    End Function

End Class

Public Class DBMapper : Inherits DBModif

    ''' <summary>DBModif name of target range</summary>
    Private paramTargetName As String
    ''' <summary>Environment (integer) where connection id should be taken from (if not existing, take from selectedEnvironment)</summary>
    Private env As Integer = 0
    ''' <summary>Database to store to</summary>
    Private database As String
    ''' <summary>Database Table, where Data is to be stored</summary>
    Private tableName As String = ""
    ''' <summary>String containing primary Key names for updating table data, comma separated</summary>
    Private primKeysStr As String = ""
    ''' <summary>if set, then insert row into table if primary key is missing there. Default = False (only update)</summary>
    Private insertIfMissing As Boolean = False
    ''' <summary>additional stored procedure to be executed after saving</summary>
    Private executeAdditionalProc As String = ""
    ''' <summary>columns to be ignored (helper columns)</summary>
    Private ignoreColumns As String = ""
    ''' <summary>respect C/U/D Flags (DBSheet functionality)</summary>
    Private CUDFlags As Boolean = False

    Public Sub New(sheetName As String, defkey As String, paramTarget As Excel.Range)
        dbmapdefkey = defkey
        DBmapSheet = sheetName
        ' if no target range is set, then no parameters can be found...
        If IsNothing(paramTarget) Then Exit Sub
        paramTargetName = getDBModifNameFromRange(paramTarget)
        If Left(paramTargetName, 8) <> "DBMapper" Then
            LogError("target " & paramTargetName & " not matching passed DBModif type DBMapper for " & DBmapSheet & "/" & dbmapdefkey & "!")
            Exit Sub
        End If
        Try
            paramText = paramTarget.Parent.Parent.CustomDocumentProperties(paramTargetName).Value
        Catch ex As Exception
            ErrorMsg("No paramText found in CustomDocumentProperties for " + paramTargetName)
            Exit Sub
        End Try
        TargetRange = paramTarget
        targetRangeAddress = paramTarget.Parent.Name + "!" + paramTarget.Address

        Dim DBModifParams() As String = functionSplit(paramText, ",", """", "def", "(", ")")
        If IsNothing(DBModifParams) Then Exit Sub
        ' check for completeness
        If DBModifParams.Length < 4 Then
            ErrorMsg("At least environment (can be empty), database, Tablename and primary keys have to be provided as DBMapper parameters !")
            Exit Sub
        End If

        ' fill parameters:
        If DBModifParams(0) <> "" Then env = Convert.ToInt16(DBModifParams(0))
        database = DBModifParams(1).Replace("""", "").Trim
        If database = "" Then
            ErrorMsg("No database given in DBMapper paramText!")
            Exit Sub
        End If
        tableName = DBModifParams(2).Replace("""", "").Trim ' remove all quotes and trim right and left
        If tableName = "" Then
            ErrorMsg("No Tablename given in DBMapper paramText!")
            Exit Sub
        End If
        primKeysStr = DBModifParams(3).Replace("""", "").Trim
        If primKeysStr = "" Then
            ErrorMsg("No primary keys given in DBMapper paramText!")
            Exit Sub
        End If
        If DBModifParams.Length > 4 AndAlso DBModifParams(4) <> "" Then insertIfMissing = Convert.ToBoolean(DBModifParams(4))
        If DBModifParams.Length > 5 AndAlso DBModifParams(5) <> "" Then executeAdditionalProc = DBModifParams(5).Replace("""", "").Trim
        If DBModifParams.Length > 6 AndAlso DBModifParams(6) <> "" Then ignoreColumns = DBModifParams(6).Replace("""", "").Trim
        If DBModifParams.Length > 7 AndAlso DBModifParams(7) <> "" Then execOnSave = Convert.ToBoolean(DBModifParams(7))
        If DBModifParams.Length > 8 AndAlso DBModifParams(8) <> "" Then CUDFlags = Convert.ToBoolean(DBModifParams(8))
    End Sub

    ''' <summary>is saving needed for this DBModifier</summary>
    ''' <returns>true for saving necessary</returns>
    Public Overrides Function DBModifSaveNeeded() As Boolean
        DBModifSaveNeeded = execOnSave
    End Function

    Public Overrides Sub doDBModif(Optional WbIsSaving As Boolean = False, Optional calledByDBSeq As String = "")
        If env = 0 Then env = Globals.selectedEnvironment + 1 ' if Environment is not existing, take from selectedEnvironment
        ' extend DataRange if it is only one cell ...
        If TargetRange.Rows.Count = 1 And TargetRange.Columns.Count = 1 Then
            Dim rowEnd = TargetRange.End(Excel.XlDirection.xlDown).Row
            Dim colEnd = TargetRange.End(Excel.XlDirection.xlToRight).Column
            TargetRange = TargetRange.Parent.Range(TargetRange, TargetRange.Parent.Cells(rowEnd, colEnd))
        End If

        ' set up parameters
        If WbIsSaving And Not execOnSave Then Exit Sub
        Dim primKeys() As String = Split(primKeysStr, ",")
        ignoreColumns = LCase(ignoreColumns) + "," ' lowercase and add comma for better retrieval

        'now create/get a connection (dbcnn) for env(ironment)
        If Not openConnection(env, database) Then Exit Sub

        'checkrst is opened to get information about table schema (field types)
        Dim checkrst As ADODB.Recordset = New ADODB.Recordset
        Dim rst As ADODB.Recordset = New ADODB.Recordset
        Try
            checkrst.Open(tableName, dbcnn, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly, CommandTypeEnum.adCmdTableDirect)
        Catch ex As Exception
            LogError("Table: " & tableName & " caused error: " & ex.Message & " in sheet " & TargetRange.Parent.Name)
            checkrst.Close()
            GoTo cleanup
        End Try

        ' to find the record to be updated, get types for primKeyCompound to build WHERE Clause with it
        Dim checkTypes() As CheckTypeFld = Nothing
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
        Dim colNum As Long = 1
        Do
            Dim fieldname As String = Trim(TargetRange.Cells(1, colNum).Value)
            ' only if not ignored...
            If InStr(1, LCase(fieldname) + ",", ignoreColumns) = 0 Then
                Try
                    Dim testExist As String = checkrst.Fields(fieldname).Name
                Catch ex As Exception
                    TargetRange.Parent.Activate
                    TargetRange.Cells(1, colNum).Select
                    LogError("Field '" & fieldname & "' does not exist in Table '" & tableName & "' and is not in ignoreColumns, Error in sheet " & TargetRange.Parent.Name)
                    GoTo cleanup
                End Try
            End If
            colNum += 1
        Loop Until colNum > TargetRange.Columns.Count
        checkrst.Close()

        Dim rowNum As Long = 2
        dbcnn.CursorLocation = CursorLocationEnum.adUseServer

        Dim finishLoop As Boolean
        ' walk through rows
        Do

            ' try to find record for update, construct where clause with primary key columns
            Dim primKeyCompound As String = " WHERE "
            For i As Integer = 0 To UBound(primKeys)
                Dim primKeyValue
                primKeyValue = TargetRange.Cells(rowNum, i + 1).Value
                primKeyCompound = primKeyCompound & primKeys(i) & " = " & dbFormatType(primKeyValue, checkTypes(i)) & IIf(i = UBound(primKeys), "", " AND ")
                If IsError(primKeyValue) Then
                    ErrorMsg("Error in primary key value, cell (" & rowNum & "," & i + 1 & ") in sheet " & TargetRange.Parent.Name & ", & row " & rowNum)
                    GoTo nextRow
                End If
                If primKeyValue.ToString().Length = 0 Then
                    ErrorMsg("Empty primary key value, cell (" & rowNum & "," & i + 1 & ") in sheet " & TargetRange.Parent.Name & ", & row " & rowNum)
                    GoTo nextRow
                End If
            Next
            Try
                rst.Open("SELECT * FROM " & tableName & primKeyCompound, dbcnn, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)
            Catch ex As Exception
                LogError("Problem getting recordset, Error: " & ex.Message & " in sheet " & TargetRange.Parent.Name & ", & row " & rowNum)
                rst.Close()
                GoTo cleanup
            End Try

            ' didn't find record, so add a new record if insertIfMissing flag is set
            If rst.EOF Then
                Dim i As Integer
                If insertIfMissing Then
                    ExcelDnaUtil.Application.StatusBar = "Inserting " & primKeyCompound & " in table " & tableName
                    rst.AddNew()
                    For i = 0 To UBound(primKeys)
                        rst.Fields(primKeys(i)).Value = IIf(TargetRange.Cells(rowNum, i + 1).ToString().Length = 0, vbNull, TargetRange.Cells(rowNum, i + 1).Value)
                    Next
                Else
                    TargetRange.Parent.Activate
                    TargetRange.Cells(rowNum, i + 1).Select
                    LogError("Problem getting recordset " & primKeyCompound & " from table '" & tableName & "', insertIfMissing = " & insertIfMissing & " in sheet " & TargetRange.Parent.Name & ", & row " & rowNum)
                    rst.Close()
                    GoTo cleanup
                End If
            Else
                ExcelDnaUtil.Application.StatusBar = "Updating " & primKeyCompound & " in table " & tableName
            End If

            ' walk through non primary columns and fill fields
            colNum = UBound(primKeys) + 1
            Do
                Dim fieldname As String = TargetRange.Cells(1, colNum).Value
                If InStr(1, ignoreColumns, LCase(fieldname) + ",") = 0 Then
                    Try
                        rst.Fields(fieldname).Value = IIf(TargetRange.Cells(rowNum, colNum).ToString().Length = 0, vbNull, TargetRange.Cells(rowNum, colNum).Value)
                    Catch ex As Exception
                        TargetRange.Parent.Activate
                        TargetRange.Cells(rowNum, colNum).Select

                        ErrorMsg("Field Value Update Error: " & ex.Message & " with Table: " & tableName & ", Field: " & fieldname & ", in sheet " & TargetRange.Parent.Name & ", & row " & rowNum & ", col: " & colNum)
                        rst.CancelUpdate()
                        rst.Close()
                        GoTo cleanup
                    End Try
                End If
                colNum += 1
            Loop Until colNum > TargetRange.Columns.Count

            ' now do the update/insert in the DB
            Try
                rst.Update()
            Catch ex As Exception
                TargetRange.Parent.Activate
                TargetRange.Rows(rowNum).Select
                ErrorMsg("Row Update Error, Table: " & rst.Source & ", Error: " & ex.Message & " in sheet " & TargetRange.Parent.Name & ", & row " & rowNum)
                rst.CancelUpdate()
                rst.Close()
                GoTo cleanup
            End Try
            rst.Close()
nextRow:
            Try
                finishLoop = IIf(TargetRange.Cells(rowNum + 1, 1).ToString().Length = 0, True, False)
            Catch ex As Exception
                ErrorMsg("Error in first primary column: Cells(" & rowNum + 1 & ",1): " & ex.Message)
                'finishLoop = True ' commented to allow erroneous data...
            End Try
            rowNum += 1
        Loop Until rowNum > TargetRange.Rows.Count Or finishLoop

        ' any additional stored procedures to execute?
        If executeAdditionalProc.Length > 0 Then
            Try
                dbcnn.Execute(executeAdditionalProc)
            Catch ex As Exception
                ErrorMsg("Error in executing additional stored procedure: " & ex.Message)
                GoTo cleanup
            End Try
        End If
cleanup:
        ' close connection to return it to the pool...
        dbcnn.Close()
        ExcelDnaUtil.Application.StatusBar = False
    End Sub

    Public Overrides Sub setDBModifCreateFields(ByRef theDBModifCreateDlg As DBModifCreate)
        With theDBModifCreateDlg
            Try
                If env > 0 Then .envSel.SelectedIndex = env - 1
            Catch ex As Exception
                ErrorMsg("Error setting environment " & env & " (correct environment manually in docproperty " & paramTargetName & "): " & ex.Message)
                Exit Sub
            End Try
            .TargetRangeAddress.Text = targetRangeAddress
            .Database.Text = database
            .execOnSave.Checked = execOnSave
            .Tablename.Text = tableName
            .PrimaryKeys.Text = primKeysStr
            .insertIfMissing.Checked = insertIfMissing
            .addStoredProc.Text = executeAdditionalProc
            .IgnoreColumns.Text = ignoreColumns
            .CUDflags.Checked = CUDFlags
        End With
    End Sub
End Class

Public Class DBAction
    Inherits DBModif

    ''' <summary>Environment (integer) where connection id should be taken from (if not existing, take from selectedEnvironment)</summary>
    Private env As Integer = 0
    ''' <summary>Database to store to</summary>
    Private database As String
    ''' <summary>DBModif name of target range</summary>
    Private paramTargetName As String

    Public Sub New(sheetName As String, defkey As String, paramTarget As Excel.Range)
        DBmapSheet = sheetName
        dbmapdefkey = defkey
        ' if no target range is set, then no parameters can be found...
        If IsNothing(paramTarget) Then Exit Sub
        paramTargetName = getDBModifNameFromRange(paramTarget)
        If Left(paramTargetName, 8) <> "DBAction" Then
            LogError("target " & paramTargetName & " not matching passed DBModif type DBAction for " & DBmapSheet & "/" & dbmapdefkey & " !")
            Exit Sub
        End If
        ' set up parameters
        If paramTarget.Cells(1, 1).Text = "" Then
            ErrorMsg("No Action defined in " + paramTargetName)
            Exit Sub
        End If
        Try
            paramText = paramTarget.Parent.Parent.CustomDocumentProperties(paramTargetName).Value
        Catch ex As Exception
            ErrorMsg("No paramText found in CustomDocumentProperties for " + paramTargetName)
            Exit Sub
        End Try
        TargetRange = paramTarget
        Dim DBModifParams() As String = functionSplit(paramText, ",", """", "def", "(", ")")
        If IsNothing(DBModifParams) Then Exit Sub
        ' check for completeness
        If DBModifParams.Length < 2 Then
            ErrorMsg("At least environment (can be empty) and database have to be provided as DBAction parameters !")
            Exit Sub
        End If
        ' fill parameters:
        If DBModifParams(0) <> "" Then env = Convert.ToInt16(DBModifParams(0))
        database = DBModifParams(1).Replace("""", "").Trim
        If database = "" Then
            ErrorMsg("No database given in DBAction paramText!")
            Exit Sub
        End If
        If DBModifParams.Length > 2 AndAlso DBModifParams(2) <> "" Then execOnSave = Convert.ToBoolean(DBModifParams(2))
    End Sub

    ''' <summary>is saving needed for this DBModifier</summary>
    ''' <returns>true for saving necessary</returns>
    Public Overrides Function DBModifSaveNeeded() As Boolean
        DBModifSaveNeeded = execOnSave
    End Function

    Public Overrides Sub doDBModif(Optional WbIsSaving As Boolean = False, Optional calledByDBSeq As String = "")
        If env = 0 Then env = Globals.selectedEnvironment + 1 ' if Environment is not existing, take from selectedEnvironment
        If WbIsSaving And Not execOnSave Then Exit Sub
        'now create/get a connection (dbcnn) for env(ironment)
        If Not openConnection(env, database) Then Exit Sub
        Dim result As Long = 0
        Try
            dbcnn.Execute(TargetRange.Cells(1, 1).Text, result, Options:=CommandTypeEnum.adCmdText)
        Catch ex As Exception
            ErrorMsg("Error: " & paramTargetName & ": " & ex.Message)
            Exit Sub
        End Try
        If Not WbIsSaving Then
            MsgBox("DBAction " & paramTargetName & " executed, affected records: " & result)
        End If
        ' close connection to return it to the pool...
        dbcnn.Close()
    End Sub

    Public Overrides Sub setDBModifCreateFields(ByRef theDBModifCreateDlg As DBModifCreate)
        With theDBModifCreateDlg
            Try
                If env > 0 Then .envSel.SelectedIndex = env - 1
            Catch ex As Exception
                ErrorMsg("Error setting environment " & env & " (correct environment manually in docproperty " & paramTargetName & "): " & ex.Message)
                Exit Sub
            End Try
            .TargetRangeAddress.Text = targetRangeAddress
            .Database.Text = database
            .execOnSave.Checked = execOnSave
        End With
    End Sub
End Class

Public Class DBSeqnce
    Inherits DBModif

    ''' <summary>sequence of DBModifiers being executed in this sequence</summary>
    Private sequenceParams() As String

    Public Sub New(defkey As String, DBSequenceText As String)
        dbmapdefkey = defkey
        paramText = DBSequenceText
        If paramText = "" Then
            ErrorMsg("No Sequence defined in " + dbmapdefkey)
            Exit Sub
        End If
        ' parse parameters: 1st item is execOnSave, rest defines sequence (tripletts of Type:WsheetID:Name)
        sequenceParams = Split(paramText, ",")
        execOnSave = Convert.ToBoolean(sequenceParams(0)) ' should DBSequence be done on Excel Saving?
    End Sub

    ''' <summary>is saving needed for this DBModifier</summary>
    ''' <returns>true for saving necessary</returns>
    Public Overrides Function DBModifSaveNeeded() As Boolean
        DBModifSaveNeeded = execOnSave
    End Function

    Public Overrides Sub doDBModif(Optional WbIsSaving As Boolean = False, Optional calledByDBSeq As String = "")
        If WbIsSaving And Not execOnSave Then Exit Sub
        Dim i As Integer
        For i = 1 To UBound(sequenceParams)
            Dim definition() As String = Split(sequenceParams(i), ":")
            If definition(0) <> "DBRefrsh" Then
                DBModifDefColl(definition(1)).Item(definition(2)).doDBModif(WbIsSaving, calledByDBSeq:=dbmapdefkey)
            Else
                ' reset query cache, so we really get new data !
                Functions.queryCache.Clear()
                Functions.StatusCollection.Clear()
                ' refresh DBFunction in sequence
                Dim underlyingName As String = definition(2)
                If Not ExcelDnaUtil.Application.Range(underlyingName).Parent Is ExcelDnaUtil.Application.ActiveSheet Then
                    ExcelDnaUtil.Application.ScreenUpdating = False
                    origWS = ExcelDnaUtil.Application.ActiveSheet
                    Try : ExcelDnaUtil.Application.Range(underlyingName).Parent.Select : Catch ex As Exception : End Try
                End If
                ExcelDnaUtil.Application.Range(underlyingName).Dirty()
                If ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual Then ExcelDnaUtil.Application.Calculate()
            End If
        Next
    End Sub

    Public Overrides Sub setDBModifCreateFields(ByRef theDBModifCreateDlg As DBModifCreate)
        theDBModifCreateDlg.execOnSave.Checked = execOnSave
    End Sub
End Class

''' <summary>Contains DBModif functions for storing/updating tabular excel data (DBMapper), doing DBActions, doing DBSequences (combinations of DBMapper/DBAction) and some helper functions</summary>
Public Module DBModifs

    ''' <summary>main db connection For mapper</summary>
    Public dbcnn As ADODB.Connection

    ''' <summary>opens a database connection</summary>
    ''' <param name="env">number of the environment as given in the settings</param>
    ''' <param name="database">database to replace database selection parameter in connection string of environment</param>
    ''' <returns>True on success</returns>
    Public Function openConnection(env As Integer, database As String) As Boolean
        openConnection = False

        Dim theConnString As String = fetchSetting("ConstConnString" & env, String.Empty)
        If theConnString = String.Empty Then
            ErrorMsg("No Connectionstring given for environment: " & env & ", please correct and rerun.")
            Exit Function
        End If
        Dim dbidentifier As String = fetchSetting("DBidentifierCCS" & env, String.Empty)
        If dbidentifier = String.Empty Then
            ErrorMsg("No DB identifier given for environment: " & env & ", please correct and rerun.")
            Exit Function
        End If

        ' connections are pooled by ADO depending on the connection string:
        dbcnn = New Connection
        theConnString = Change(theConnString, dbidentifier, database, ";")
        LogInfo("open connection with " & theConnString)
        ExcelDnaUtil.Application.StatusBar = "Trying " & Globals.CnnTimeout & " sec. with connstring: " & theConnString
        Try
            dbcnn.ConnectionString = theConnString
            dbcnn.ConnectionTimeout = Globals.CnnTimeout
            dbcnn.CommandTimeout = Globals.CmdTimeout
            dbcnn.Open()
        Catch ex As Exception
            LogError("Error connecting to DB: " & ex.Message & ", connection string: " & theConnString)
            If dbcnn.State = ADODB.ObjectStateEnum.adStateOpen Then dbcnn.Close()
            dbcnn = Nothing
        End Try
        ExcelDnaUtil.Application.StatusBar = String.Empty
        openConnection = True
    End Function

    ''' <summary>creates a DBModif at the current active cell or edits an existing one being there or after being called from ribbon + Ctrl</summary>
    Sub createDBModif(type As String, Optional targetRange As Excel.Range = Nothing, Optional targetDefName As String = "", Optional DBSequenceText As String = "")
        ' clipboard helper for legacy definitions:
        ' if saveRangeToDB macro calls were copied, rename saveRangeToDB<Single> To def, 1st parameter (datarange) removed (empty), connid moved to 2nd place as database name (remove MSSQL)
        'mapper.saveRangeToDBSingle(Range("DB_DefName"), "tableName", "primKey1,primKey2,primKey3", "MSSQLDB_NAME", True)
        '--> def(, "DB_NAME", "tableName", "primKey1,primKey2,primKey3", True)    DBMapperName = DB_DefName
        'mapper.saveRangeToDBSingle(Range("DB_DefName"), "tableName", "primKey1,primKey2,primKey3", "MSSQLDB_NAME")
        '--> def(, "DB_NAME", "tableName", "primKey1,primKey2,primKey3")          DBMapperName = DB_DefName
        'def(, "DB_NAME", True), "tableName", "primKey1,primKey2,primKey3", "MSSQLDB_NAME", True)", "MSSQLDB_NAME", True)
        Dim createdDBMapperFromClipboard As Boolean = False
        If Clipboard.ContainsText() And type = "DBMapper" Then
            Dim cpbdtext As String = Clipboard.GetText()
            If InStr(cpbdtext.ToLower(), "saverangetodb") > 0 Then
                Dim firstBracket As Integer = InStr(cpbdtext, "(")
                Dim firstComma As Integer = InStr(cpbdtext, ",")
                Dim connDefStart As Integer = InStrRev(cpbdtext, """MSSQL")
                Dim commaBeforeConnDef As Integer = InStrRev(cpbdtext, ",", connDefStart)
                ' after conndef, all parameters are optional, so in case there is no comma afterwards, set this to end of whole definition string
                Dim commaAfterConnDef As Integer = IIf(InStr(connDefStart, cpbdtext, ",") > 0, InStr(connDefStart, cpbdtext, ","), Len(cpbdtext))
                Dim DB_DefName, newDefString As String
                Try : DB_DefName = "DBMapper" + Replace(Replace(Mid(cpbdtext, firstBracket + 1, firstComma - firstBracket - 1), "Range(""", ""), """)", "")
                Catch ex As Exception : ErrorMsg("Error when retrieving DB_DefName from clipboard: " & ex.Message) : Exit Sub : End Try
                Try : newDefString = "def(" + Replace(Mid(cpbdtext, commaBeforeConnDef, commaAfterConnDef - commaBeforeConnDef), "MSSQL", "") + Mid(cpbdtext, firstComma, commaBeforeConnDef - firstComma - 1) + Mid(cpbdtext, commaAfterConnDef - 1)
                Catch ex As Exception : ErrorMsg("Error when building new definition from clipboard: " & ex.Message) : Exit Sub : End Try
                If IsNothing(targetRange) Then targetRange = ExcelDnaUtil.Application.ActiveCell
                Try : targetRange.Name = DB_DefName
                Catch ex As Exception
                    ErrorMsg("Error when assigning name '" & DB_DefName & "' to active cell: " & ex.Message)
                    targetRange.Name.Delete
                    Exit Sub
                End Try
                Try : ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties(DB_DefName).Delete : Catch ex As Exception : End Try
                Try
                    ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties.Add(Name:=DB_DefName, LinkToContent:=False, Type:=Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, Value:=newDefString)
                Catch ex As Exception
                    ErrorMsg("Error when adding CustomDocumentProperty with DBModif parameters (Name:" & DB_DefName & ",content: " & newDefString & "): " & ex.Message)
                    ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties(DB_DefName).Delete
                    Exit Sub
                End Try
                createdDBMapperFromClipboard = True
                Clipboard.Clear()
            End If
        End If

        ' start normal creation
        Dim activeCellName As String = ""
        If Not IsNothing(targetRange) Then activeCellName = getDBModifNameFromRange(targetRange) ' try regular defined name
        If type = "DBSeqnce" Then activeCellName = "DBSeqnce" & targetDefName
        If type = "DBMapper" Then
            ' try potential name to ListObject (parts), only possible if manually defined !
            If activeCellName = "" And Not IsNothing(ExcelDnaUtil.Application.Selection.ListObject) Then
                For Each listObjectCol In ExcelDnaUtil.Application.Selection.ListObject.ListColumns
                    For Each aName As Excel.Name In ExcelDnaUtil.Application.ActiveWorkbook.Names
                        If aName.RefersTo = "=" & ExcelDnaUtil.Application.Selection.ListObject.Name & "[[#Headers],[" & ExcelDnaUtil.Application.Selection.Value & "]]" Then
                            activeCellName = aName.Name
                            Exit For
                        End If
                    Next
                    If activeCellName <> "" Then Exit For
                Next
            End If
        End If

        ' fetch parameters if existing target range and matching definition...
        Dim existingDBModif As DBModif = Nothing
        Try
            If type = "DBMapper" Then
                existingDBModif = New DBMapper("", "", targetRange)
            ElseIf type = "DBAction" Then
                existingDBModif = New DBAction("", "", targetRange)
            ElseIf type = "DBSeqnce" Then
                existingDBModif = New DBSeqnce(targetDefName, DBSequenceText)
            Else
                LogError("Error, not supported DBModiftype: " & type)
            End If
        Catch ex As Exception : End Try

        ' prepare DBModifier Create Dialog
        Dim theDBModifCreateDlg As DBModifCreate = New DBModifCreate()
        With theDBModifCreateDlg
            .envSel.DataSource = Globals.environdefs
            .envSel.SelectedIndex = -1
            If Not IsNothing(existingDBModif) Then
                If InStr(1, activeCellName, type) > 0 Then .DBModifName.Text = Replace(activeCellName, type, "")
                ' delegate filling of dialog fields to created DBModif object
                existingDBModif.setDBModifCreateFields(theDBModifCreateDlg)
            End If
            .NameLabel.Text = IIf(type = "DBSeqnce", "DBSequence", type) & " Name:"
            .Text = "Edit " & IIf(type = "DBSeqnce", "DBSequence", type) & " definition"
            If type <> "DBMapper" Then
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
            End If
            If type = "DBSeqnce" Then
                .DBModifName.Text = targetDefName
                ' hide controls irrelevant for DBSeqnce
                .TargetRangeLabel.Hide()
                .TargetRangeAddress.Hide()
                .envSel.Hide()
                .EnvironmentLabel.Hide()
                .Database.Hide()
                .DatabaseLabel.Hide()
                .DBSeqenceDataGrid.Top = 55
                .DBSeqenceDataGrid.Height = 320
                .execOnSave.Top = .TargetRangeLabel.Top
                ' fill Datagridview for DBSequence
                Dim cb As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn()
                cb.HeaderText = "Sequence Step"
                cb.ReadOnly = False
                cb.ValueType() = GetType(String)
                Dim ds As List(Of String) = New List(Of String)
                ' first add the DBMapper and DBAction definitions available in the Workbook
                For Each sheetId As String In DBModifDefColl.Keys
                    For Each nodeName As String In DBModifDefColl(sheetId).Keys
                        ' avoid DB Sequences (might be - indirectly - self referencing, leading to endless recursion)
                        If sheetId <> "ID0" Then
                            ' for DBMapper and DBAction, full name is only available from the target range name
                            Dim targetRangeName As String = getDBModifNameFromRange(DBModifDefColl(sheetId).Item(nodeName).getTargetRange())
                            ds.Add(Left(targetRangeName, 8) & ":" & sheetId & ":" & Right(targetRangeName, Len(targetRangeName) - 8))
                        End If
                    Next
                Next
                Dim searchCell As Excel.Range
                ' then add DBRefresh items for allowing refreshing DBFunctions (DBListFetch and DBSetQuery) during a Sequence
                For Each ws As Excel.Worksheet In ExcelDnaUtil.Application.ActiveWorkbook.Worksheets
                    For Each theFunc As String In {"DBListFetch(", "DBSetQuery("}
                        searchCell = ws.Cells.Find(What:=theFunc, After:=ws.Range("A1"), LookIn:=Excel.XlFindLookIn.xlFormulas, LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlNext, MatchCase:=False)
                        If Not (searchCell Is Nothing) Then
                            If searchCell.Rows.Count > 1 Or searchCell.Rows.Count > 1 Then
                                LogError(theFunc & " (in " & searchCell.Parent.Name & "!" & searchCell.Address & ") has multiple " & IIf(searchCell.Rows.Count > 1, "rows !", "columns !") & ", which leads to problems in DBSequences...")
                                Continue For
                            End If
                            Dim underlyingName As String = getDBunderlyingNameFromRange(searchCell)
                            ds.Add("DBRefrsh:" & theFunc & searchCell.Parent.Name & "!" & searchCell.Address & "):" & underlyingName)
                            ' reset the cell find dialog....
                            searchCell = Nothing
                        End If
                    Next
                    ' reset the cell find dialog....
                    searchCell = Nothing
                    searchCell = ws.Cells.Find(What:="", After:=ws.Range("A1"), LookIn:=Excel.XlFindLookIn.xlFormulas, LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlNext, MatchCase:=False)
                Next
                cb.DataSource() = ds
                .DBSeqenceDataGrid.Columns.Add(cb)
                ' fill possible existing sequence definitions into form/Datagridview
                If Len(DBSequenceText) > 0 Then
                    Dim params() As String = Split(DBSequenceText, ",")
                    .execOnSave.Checked = Convert.ToBoolean(params(0))
                    For i As Integer = 1 To UBound(params)
                        .DBSeqenceDataGrid.Rows.Add(params(i))
                    Next
                End If
                .DBSeqenceDataGrid.Columns(0).Width = 200
            Else
                ' hide controls irrelevant for DBMapper and DBAction
                .up.Hide()
                .down.Hide()
                .DBSeqenceDataGrid.Hide()
            End If
            ' store DBModification type in tag for validation purposes...
            theDBModifCreateDlg.Tag = type
            ' display dialog for parameters
            If theDBModifCreateDlg.ShowDialog() = DialogResult.Cancel Then
                ' remove name and customdocproperty created in clipboard helper
                If createdDBMapperFromClipboard Then
                    Try
                        targetRange.Name.Delete()
                        ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties(activeCellName).Delete
                    Catch ex As Exception : End Try
                End If
                Exit Sub
            End If

            ' only for DBMapper or DBAction: change target range name
            If type <> "DBSeqnce" Then
                If IsNothing(targetRange) Then targetRange = ExcelDnaUtil.Application.ActiveCell
                ' set content range name: first clear name
                If InStr(1, activeCellName, type) > 0 Then   ' fetch parameters if existing comment and DBMapper definition...
                    Try
                        For Each DBModifName As Excel.Name In ExcelDnaUtil.Application.ActiveWorkbook.Names
                            If DBModifName.Name = activeCellName Then DBModifName.Delete()
                        Next
                    Catch ex As Exception : ErrorMsg("Error when removing name '" + activeCellName + "' from active cell: " & ex.Message) : End Try
                End If
                ' then (re)set name
                Try : targetRange.Name = type + .DBModifName.Text
                Catch ex As Exception : ErrorMsg("Error when assigning name '" & type & .DBModifName.Text & "' to active cell: " & ex.Message)
                End Try
            End If

            ' create parameter definition string ...
            Dim newParamText As String = ""
            If type = "DBAction" Then
                newParamText = "def(" + IIf(.envSel.SelectedIndex = -1, "", (.envSel.SelectedIndex + 1).ToString()) + "," + """" + .Database.Text + """," + .execOnSave.Checked.ToString() + ")"
            ElseIf type = "DBMapper" Then
                newParamText = "def(" +
                    IIf(.envSel.SelectedIndex = -1, "", (.envSel.SelectedIndex + 1).ToString()) + "," +
                    """" + .Database.Text + """," + """" + .Tablename.Text + """," + """" + .PrimaryKeys.Text + """," + .insertIfMissing.Checked.ToString() + "," +
                    """" + IIf(Len(.addStoredProc.Text) = 0, "", .addStoredProc.Text) + """," +
                    """" + IIf(Len(.IgnoreColumns.Text) = 0, "", .IgnoreColumns.Text) + """," +
                    .execOnSave.Checked.ToString() + "," + .CUDflags.Checked.ToString() + ")"
            ElseIf type = "DBSeqnce" Then
                newParamText = .execOnSave.Checked.ToString()
                ' need that because empty row at the end is passed along with Rows() !!
                For i As Integer = 0 To .DBSeqenceDataGrid.Rows().Count - 2
                    newParamText += "," + .DBSeqenceDataGrid.Rows(i).Cells(0).Value
                Next
            End If
            ' ... and store in docproperty (rename docproperty first to current name, might have been changed)
            Try : ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties(activeCellName).Delete : Catch ex As Exception : End Try
            Try
                ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties.Add(Name:=type + .DBModifName.Text, LinkToContent:=False, Type:=Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, Value:=newParamText)
            Catch ex As Exception : ErrorMsg("Error when adding property with DBModif parameters: " & ex.Message) : End Try
        End With
        ' refresh mapper definitions to reflect changes immediately...
        getDBModifDefinitions()
    End Sub

    ''' <summary>gets defined names for DBModifier (DBMapper/DBAction/DBSeqnce) invocation in the current workbook and updates Ribbon with it</summary>
    Public Sub getDBModifDefinitions()
        ' load DBModifier definitions
        Try
            Globals.DBModifDefColl = New Dictionary(Of String, Dictionary(Of String, DBModif))
            ' add DB sequences on Workbook level...
            For Each docproperty In ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties
                If TypeName(docproperty.Value) = "String" And Left(docproperty.Name, 8) = "DBSeqnce" Then
                    Dim nodeName As String = Replace(docproperty.Name, "DBSeqnce", "")

                    Dim defColl As Dictionary(Of String, DBModif)
                    If Not DBModifDefColl.ContainsKey("ID0") Then
                        ' add to new sheet "menu"
                        defColl = New Dictionary(Of String, DBModif)
                        defColl.Add(nodeName, New DBSeqnce(nodeName, docproperty.Value))
                        DBModifDefColl.Add("ID0", defColl)
                    Else
                        ' add definition to existing sheet "menu"
                        defColl = DBModifDefColl("ID0")
                        defColl.Add(nodeName, New DBSeqnce(nodeName, docproperty.Value))
                    End If
                End If
            Next
            For Each namedrange As Excel.Name In ExcelDnaUtil.Application.ActiveWorkbook.Names
                Dim cleanname As String = Replace(namedrange.Name, namedrange.Parent.Name & "!", "")
                Dim DBModiftype As String = Left(cleanname, 8)
                If DBModiftype = "DBMapper" Or DBModiftype = "DBAction" Then
                    If InStr(namedrange.RefersTo, "#REF!") > 0 Then
                        ErrorMsg(DBModiftype + " definitions range " + namedrange.Parent.Name + "!" + namedrange.Name + " contains #REF!")
                        Continue For
                    End If
                    Dim nodeName As String = Replace(cleanname, DBModiftype, "")

                    Dim sheetID As String = namedrange.RefersToRange.Parent.Name
                    Dim defColl As Dictionary(Of String, DBModif)
                    Dim newDBModif As DBModif
                    If DBModiftype = "DBMapper" Then
                        newDBModif = New DBMapper(sheetID, nodeName, namedrange.RefersToRange)
                    ElseIf DBModiftype = "DBAction" Then
                        newDBModif = New DBAction(sheetID, nodeName, namedrange.RefersToRange)
                    Else
                        LogError("Error, not supported DBModiftype: " & DBModiftype)
                        newDBModif = Nothing
                    End If
                    If Not DBModifDefColl.ContainsKey(sheetID) Then
                        ' add to new sheet "menu"
                        defColl = New Dictionary(Of String, DBModif)
                        defColl.Add(nodeName, newDBModif)
                        If DBModifDefColl.Count = 15 Then
                            ErrorMsg("Not more than 15 sheets with DBMapper/DBAction/DBSequence definitions possible, ignoring definitions in sheet " + namedrange.Parent.Name)
                            Exit For
                        End If
                        DBModifDefColl.Add(sheetID, defColl)
                    Else
                        ' add definition to existing sheet "menu"
                        defColl = DBModifDefColl(sheetID)
                        defColl.Add(nodeName, newDBModif)
                    End If
                End If
            Next
            Globals.theRibbon.Invalidate()
        Catch ex As Exception
            LogError("Error: " & ex.Message)
        End Try
    End Sub

End Module

