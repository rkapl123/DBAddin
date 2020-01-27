Imports ADODB
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports System.Collections.Generic
Imports System.Windows.Forms


''' <summary>Contains DBModif functions for storing/updating tabular excel data (DBMapper), doing DBActions, doing DBSequences (combinations of DBMapper/DBAction) and some helper functions</summary>
Public Module DBModif

    ''' <summary>main db connection For mapper</summary>
    Public dbcnn As ADODB.Connection

    ''' <summary>do abstraction procedure for a given DBModifier in current workbook defined by DBmapSheet/dbmapdefkey</summary>
    ''' <param name="DBmapSheet">sheet of DBModifier</param>
    ''' <param name="dbmapdefkey">key (name) of of DBModifier</param>
    Public Sub doDBModif(DBmapSheet As String, dbmapdefkey As String, Optional WbIsSaving As Boolean = True)
        If Left(dbmapdefkey, 8) = "DBSeqnce" Then
            ' DB sequence actions (the sequence to be done) are stored directly in DBMapperDefColl, so different invocation here
            doDBSeqnce(dbmapdefkey, DBModifDefColl(DBmapSheet).Item(dbmapdefkey), WbIsSaving:=WbIsSaving)
        Else
            Dim rngName As String = getDBModifNameFromRange(DBModifDefColl(DBmapSheet).Item(dbmapdefkey))
            If Left(rngName, 8) = "DBMapper" Then
                doDBMapper(DBModifDefColl(DBmapSheet).Item(dbmapdefkey), WbIsSaving:=WbIsSaving)
            ElseIf Left(rngName, 8) = "DBAction" Then
                doDBAction(DBModifDefColl(DBmapSheet).Item(dbmapdefkey), WbIsSaving:=WbIsSaving)
            End If
        End If
    End Sub

    ''' <summary>is saving needed for the DBModifier defined by DBmapSheet/dbmapdefkey</summary>
    ''' <param name="DBmapSheet">sheet of DBModifier</param>
    ''' <param name="dbmapdefkey">key (name) of of DBModifier</param>
    ''' <returns>true for saving necessary</returns>
    Public Function DBModifSaveNeeded(DBmapSheet As String, dbmapdefkey As String) As Boolean
        DBModifSaveNeeded = False
        If Left(dbmapdefkey, 8) = "DBSeqnce" Then
            DBModifSaveNeeded = Convert.ToBoolean(Split(DBModifDefColl(DBmapSheet).Item(dbmapdefkey), ",")(0))
        Else
            Dim paramTargetName As String = getDBModifNameFromRange(DBModifDefColl(DBmapSheet).Item(dbmapdefkey))
            Dim tableName As String = ""                         ' Database Table, where Data is to be stored
            Dim primKeysStr As String = ""                       ' String containing primary Key names for updating table data, comma separated
            Dim database As String = ""                          ' Database to store to
            Dim env As Integer = Globals.selectedEnvironment + 1 ' Environment where connection id should be taken from (if not existing, take from selectedEnvironment)
            Dim insertIfMissing As Boolean = False               ' if set, then insert row into table if primary key is missing there. Default = False (only update)
            Dim executeAdditionalProc As String = ""             ' additional stored procedure to be executed after saving
            Dim ignoreColumns As String = ""                     ' columns to be ignored (helper columns)
            Dim execOnSave As Boolean = False                    ' should DBMap be saved on Excel Saving? (default no)
            Dim paramText As String = ""
            Try
                paramText = DBModifDefColl(DBmapSheet).Item(dbmapdefkey).Parent.Parent.CustomDocumentProperties(paramTargetName).Value
            Catch ex As Exception : End Try
            ' set up parameters
            If Not getParametersFromParamText(paramText, paramType:=Left(dbmapdefkey, 8), env:=env, database:=database, tableName:=tableName, primKeysStr:=primKeysStr, insertIfMissing:=insertIfMissing, executeAdditionalProc:=executeAdditionalProc, ignoreColumns:=ignoreColumns, execOnSave:=execOnSave) Then Exit Function
            DBModifSaveNeeded = execOnSave
        End If
    End Function

    ''' <summary>execute sequence of DBAction and DBMapper invocations defined in DBSequenceText</summary>
    ''' <param name="DBSequenceName">Name of DB Sequence</param>
    ''' <param name="DBSequenceText">Definition of DB Sequence: (storeDBMapOnSave flag),(Type1:WsheetID1:Name1),(Type2:WsheetID2:Name2)...</param>
    ''' <param name="WbIsSaving">special flag to indicate calling of the procedure during saving of the Workbook</param>
    Public Sub doDBSeqnce(DBSequenceName As String, DBSequenceText As String, Optional WbIsSaving As Boolean = False)
        If DBSequenceText = "" Then
            ErrorMsg("No Sequence defined in " + DBSequenceName)
            Exit Sub
        End If
        ' parse parameters: 1st item is storeDBMapOnSave, rest defines sequence (Type:WsheetID:Name)
        Dim params() As String = Split(DBSequenceText, ",")
        Dim storeDBMapOnSave As Boolean = Convert.ToBoolean(params(0)) ' should DBSequence be done on Excel Saving?
        If WbIsSaving And Not storeDBMapOnSave Then Exit Sub
        Dim i As Integer
        For i = 1 To UBound(params)
            Dim definition() As String = Split(params(i), ":")
            If definition(0) = "DBAction" Then
                doDBAction(DataRange:=DBModifDefColl(definition(1)).Item(definition(2)), calledByDBSeq:=DBSequenceName)
            ElseIf definition(0) = "DBMapper" Then
                doDBMapper(DataRange:=DBModifDefColl(definition(1)).Item(definition(2)), calledByDBSeq:=DBSequenceName)
            ElseIf definition(0) = "DBRefrsh" Then
                ' reset query cache, so we really get new data !
                queryCache = New Collection
                StatusCollection = New Collection
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

    ''' <summary>execute Database Action defined in DataRange (single cell), this can be any DML code</summary>
    ''' <param name="DataRange">Excel Range, where Database action is taken from</param>
    ''' <param name="WbIsSaving">special flag to indicate calling of the procedure during saving of the Workbook</param>
    ''' <param name="calledByDBSeq">DBSequenceName of calling DB Sequence, indicates possible duplicate invocation during saving of Workbook</param>
    Public Sub doDBAction(DataRange As Object, Optional WbIsSaving As Boolean = False, Optional calledByDBSeq As String = "")
        Dim database As String = ""                          ' Database to store to
        Dim env As Integer = Globals.selectedEnvironment + 1 ' Environment where connection id should be taken from (if not existing, take from selectedEnvironment)
        Dim execOnSave As Boolean = False                    ' should DBaction be done on Excel Saving? (default no)

        Dim DBActionName As String = getDBModifNameFromRange(DataRange)
        ' set up parameters
        Dim paramText As String = getParamTextFromTargetRange(paramTarget:=DataRange, paramType:="DBAction")
        If Not getParametersFromParamText(paramText:=paramText, paramType:="DBAction", env:=env, database:=database, execOnSave:=execOnSave) Then Exit Sub
        If DataRange.Cells(1, 1).Text = "" Then
            ErrorMsg("No Action defined in " + DBActionName)
            Exit Sub
        End If
        If WbIsSaving And Not execOnSave Then Exit Sub
        'now create/get a connection (dbcnn) for env(ironment)
        If Not openConnection(env, database) Then Exit Sub
        Dim result As Long = 0
        Try
            dbcnn.Execute(DataRange.Cells(1, 1).Text, result, Options:=CommandTypeEnum.adCmdText)
        Catch ex As Exception
            ErrorMsg("Error: " & DBActionName & ":" & ex.Message)
            Exit Sub
        End Try
        If Not WbIsSaving Then
            MsgBox("DBAction " & DBActionName & " executed, affected records: " & result)
        End If
    End Sub

    ''' <summary>dump data given in dataRange to a database table specified by tableName and connID
    ''' this database table can have multiple primary columns specified by primKeysStr
    ''' assumption: layout of dataRange is
    ''' primKey1Val,primKey2Val,..,primKeyNVal,DataCol1Val,DataCol2Val,..,DataColNVal</summary>
    ''' <param name="DataRange">Excel Range, where Data is taken from</param>
    ''' <param name="WbIsSaving">special flag to indicate calling of the procedure during saving of the Workbook</param>
    ''' <param name="calledByDBSeq">DBSequenceName of calling DB Sequence, indicates possible duplicate invocation during saving of Workbook</param>
    Public Sub doDBMapper(DataRange As Object, Optional WbIsSaving As Boolean = False, Optional calledByDBSeq As String = "")
        Dim tableName As String = ""                         ' Database Table, where Data is to be stored
        Dim primKeysStr As String = ""                       ' String containing primary Key names for updating table data, comma separated
        Dim database As String = ""                          ' Database to store to
        Dim env As Integer = Globals.selectedEnvironment + 1 ' Environment where connection id should be taken from (if not existing, take from selectedEnvironment)
        Dim insertIfMissing As Boolean = False               ' if set, then insert row into table if primary key is missing there. Default = False (only update)
        Dim executeAdditionalProc As String = ""             ' additional stored procedure to be executed after saving
        Dim ignoreColumns As String = ""                     ' columns to be ignored (helper columns)
        Dim execOnSave As Boolean = False                    ' should DBMap be saved on Excel Saving? (default no)

        Dim rst As ADODB.Recordset
        Dim checkrst As ADODB.Recordset
        Dim primKeys() As String
        Dim rowNum As Long, colNum As Long

        ' extend DataRange if it is only one cell ...
        If DataRange.Rows.Count = 1 And DataRange.Columns.Count = 1 Then
            Dim rowEnd = DataRange.End(Excel.XlDirection.xlDown).Row
            Dim colEnd = DataRange.End(Excel.XlDirection.xlToRight).Column
            DataRange = DataRange.Parent.Range(DataRange, DataRange.Parent.Cells(rowEnd, colEnd))
        End If

        ' set up parameters
        Dim paramText As String = getParamTextFromTargetRange(paramTarget:=DataRange, paramType:="DBAction")
        If Not getParametersFromParamText(paramText:=paramText, paramType:="DBMapper", env:=env, database:=database, tableName:=tableName, primKeysStr:=primKeysStr, insertIfMissing:=insertIfMissing, executeAdditionalProc:=executeAdditionalProc, ignoreColumns:=ignoreColumns, execOnSave:=execOnSave) Then Exit Sub
        If WbIsSaving And Not execOnSave Then Exit Sub
        primKeys = Split(primKeysStr, ",")
        ignoreColumns = LCase(ignoreColumns) + "," ' lowercase and add comma for better retrieval

        'now create/get a connection (dbcnn) for env(ironment)
        If Not openConnection(env, database) Then Exit Sub

        'checkrst is opened to get information about table schema (field types)
        checkrst = New ADODB.Recordset
        rst = New ADODB.Recordset
        Try
            checkrst.Open(tableName, dbcnn, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly, CommandTypeEnum.adCmdTableDirect)
        Catch ex As Exception
            LogError("Table: " & tableName & " caused error: " & ex.Message & " in sheet " & DataRange.Parent.Name)
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
        colNum = 1
        Do
            Dim fieldname As String = Trim(DataRange.Cells(1, colNum).Value)
            ' only if not ignored...
            If InStr(1, LCase(fieldname) + ",", ignoreColumns) = 0 Then
                Try
                    Dim testExist As String = checkrst.Fields(fieldname).Name
                Catch ex As Exception
                    DataRange.Parent.Activate
                    DataRange.Cells(1, colNum).Select
                    LogError("Field '" & fieldname & "' does not exist in Table '" & tableName & "' and is not in ignoreColumns, Error in sheet " & DataRange.Parent.Name)
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

            ' try to find record for update, construct where clause with primary key columns
            Dim primKeyCompound As String = " WHERE "
            For i As Integer = 0 To UBound(primKeys)
                Dim primKeyValue
                primKeyValue = DataRange.Cells(rowNum, i + 1).Value
                primKeyCompound = primKeyCompound & primKeys(i) & " = " & dbFormatType(primKeyValue, checkTypes(i)) & IIf(i = UBound(primKeys), "", " AND ")
                If IsError(primKeyValue) Then
                    ErrorMsg("Error in primary key value, cell (" & rowNum & "," & i + 1 & ") in sheet " & DataRange.Parent.Name & ", & row " & rowNum)
                    GoTo nextRow
                End If
                If primKeyValue.ToString().Length = 0 Then
                    ErrorMsg("Empty primary key value, cell (" & rowNum & "," & i + 1 & ") in sheet " & DataRange.Parent.Name & ", & row " & rowNum)
                    GoTo nextRow
                End If
            Next
            Try
                rst.Open("SELECT * FROM " & tableName & primKeyCompound, dbcnn, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)
            Catch ex As Exception
                LogError("Problem getting recordset, Error: " & ex.Message & " in sheet " & DataRange.Parent.Name & ", & row " & rowNum)
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
                        rst.Fields(primKeys(i)).Value = IIf(DataRange.Cells(rowNum, i + 1).ToString().Length = 0, vbNull, DataRange.Cells(rowNum, i + 1).Value)
                    Next
                Else
                    DataRange.Parent.Activate
                    DataRange.Cells(rowNum, i + 1).Select
                    LogError("Problem getting recordset " & primKeyCompound & " from table '" & tableName & "', insertIfMissing = " & insertIfMissing & " in sheet " & DataRange.Parent.Name & ", & row " & rowNum)
                    rst.Close()
                    GoTo cleanup
                End If
            Else
                ExcelDnaUtil.Application.StatusBar = "Updating " & primKeyCompound & " in table " & tableName
            End If

            ' walk through non primary columns and fill fields
            colNum = UBound(primKeys) + 1
            Do
                Dim fieldname As String = DataRange.Cells(1, colNum).Value
                If InStr(1, ignoreColumns, LCase(fieldname) + ",") = 0 Then
                    Try
                        rst.Fields(fieldname).Value = IIf(DataRange.Cells(rowNum, colNum).ToString().Length = 0, vbNull, DataRange.Cells(rowNum, colNum).Value)
                    Catch ex As Exception
                        DataRange.Parent.Activate
                        DataRange.Cells(rowNum, colNum).Select

                        ErrorMsg("Field Value Update Error: " & ex.Message & " with Table: " & tableName & ", Field: " & fieldname & ", in sheet " & DataRange.Parent.Name & ", & row " & rowNum & ", col: " & colNum)
                        rst.CancelUpdate()
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
                ErrorMsg("Row Update Error, Table: " & rst.Source & ", Error: " & ex.Message & " in sheet " & DataRange.Parent.Name & ", & row " & rowNum)
                rst.CancelUpdate()
                rst.Close()
                GoTo cleanup
            End Try
            rst.Close()
nextRow:
            Try
                finishLoop = IIf(DataRange.Cells(rowNum + 1, 1).ToString().Length = 0, True, False)
            Catch ex As Exception
                ErrorMsg("Error in first primary column: Cells(" & rowNum + 1 & ",1): " & ex.Message)
                'finishLoop = True ' commented to allow erroneous data...
            End Try
            rowNum += 1
        Loop Until rowNum > DataRange.Rows.Count Or finishLoop

        ' any additional stored procedures to execute?
        If executeAdditionalProc.Length > 0 Then
            Try
                dbcnn.Execute(executeAdditionalProc)
            Catch ex As Exception
                ErrorMsg("Error in executing additional stored procedure:" & ex.Message)
                GoTo cleanup
            End Try
        End If
cleanup:
        dbcnn = Nothing
        ExcelDnaUtil.Application.StatusBar = False
    End Sub

    ''' <summary>formats theVal to fit the type of column theHead having data type dataType</summary>
    ''' <param name="theVal"></param>
    ''' <param name="dataType"></param>
    ''' <returns>the formatted value</returns>
    Private Function dbFormatType(ByVal theVal As Object, dataType As CheckTypeFld) As String

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

    ''' <summary>gets defined names for DBModifier (DBMapper/DBAction/DBSeqnce) invocation in the current workbook and updates Ribbon with it</summary>
    Public Sub getDBModifDefinitions()
        ' load DBModifier definitions
        Try
            Globals.DBModifDefColl = New Dictionary(Of String, Dictionary(Of String, Object))
            ' add DB sequences on Workbook level...
            For Each docproperty In ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties
                If TypeName(docproperty.Value) = "String" And Left(docproperty.Name, 8) = "DBSeqnce" Then
                    Dim nodeName As String = Replace(docproperty.Name, "DBSeqnce", "")

                    Dim defColl As Dictionary(Of String, Object)
                    If Not DBModifDefColl.ContainsKey("ID0") Then
                        ' add to new sheet "menu"
                        defColl = New Dictionary(Of String, Object)
                        defColl.Add(nodeName, docproperty.Value)
                        DBModifDefColl.Add("ID0", defColl)
                    Else
                        ' add definition to existing sheet "menu"
                        defColl = DBModifDefColl("ID0")
                        defColl.Add(nodeName, docproperty.Value)
                    End If
                End If
            Next
            For Each namedrange As Excel.Name In ExcelDnaUtil.Application.ActiveWorkbook.Names
                Dim cleanname As String = Replace(namedrange.Name, namedrange.Parent.Name & "!", "")
                If Left(cleanname, 8) = "DBMapper" Or Left(cleanname, 8) = "DBAction" Then
                    Dim DBModiftype As String = Left(cleanname, 8)
                    If InStr(namedrange.RefersTo, "#REF!") > 0 Then
                        ErrorMsg(DBModiftype + " definitions range " + namedrange.Parent.Name + "!" + namedrange.Name + " contains #REF!")
                        Continue For
                    End If
                    Dim nodeName As String = Replace(cleanname, DBModiftype, "")

                    Dim sheetID As String = namedrange.RefersToRange.Parent.Name
                    Dim defColl As Dictionary(Of String, Object)
                    If Not DBModifDefColl.ContainsKey(sheetID) Then
                        ' add to new sheet "menu"
                        defColl = New Dictionary(Of String, Object)
                        defColl.Add(nodeName, namedrange.RefersToRange)
                        If DBModifDefColl.Count = 15 Then
                            ErrorMsg("Not more than 15 sheets with DBMapper/DBAction/DBSequence definitions possible, ignoring definitions in sheet " + namedrange.Parent.Name)
                            Exit For
                        End If
                        DBModifDefColl.Add(sheetID, defColl)
                    Else
                        ' add definition to existing sheet "menu"
                        defColl = DBModifDefColl(sheetID)
                        defColl.Add(nodeName, namedrange.RefersToRange)
                    End If
                End If
            Next
            Globals.theRibbon.Invalidate()
        Catch ex As Exception
            LogError("Error: " & ex.Message)
        End Try
    End Sub

    Function getParamTextFromTargetRange(paramTarget As Excel.Range, paramType As String) As String
        getParamTextFromTargetRange = ""
        ' if no target range is set, then no parameters can be found...
        If IsNothing(paramTarget) Then Return False
        Dim paramTargetName As String = getDBModifNameFromRange(paramTarget)
        If Left(paramTargetName, 8) <> paramType Then
            LogError("target not matching passed DBModif type " & paramType & " !")
            Return False
        End If
        Try
            getParamTextFromTargetRange = paramTarget.Parent.Parent.CustomDocumentProperties(paramTargetName).Value
        Catch ex As Exception : End Try
    End Function

    ''' <summary>get parameters for passed target Range paramTarget (stored in custom doc properties having the same DBModif Name)</summary>
    ''' <param name="paramType">DBMapper or DBAction</param>
    ''' <param name="env">Environment (integer) where connection id should be taken from (if not existing, take from selectedEnvironment)</param>
    ''' <param name="database">Database to store to</param>
    ''' <param name="tableName">Database Table, where Data is to be stored</param>
    ''' <param name="primKeysStr">String containing primary Key names for updating table data, comma separated</param>
    ''' <param name="insertIfMissing">if set, then insert row into table if primary key is missing there. Default = False (only update)</param>
    ''' <param name="executeAdditionalProc">additional stored procedure to be executed after saving</param>
    ''' <param name="ignoreColumns">columns to be ignored (helper columns)</param>
    ''' <param name="execOnSave">should DBMap be saved / DBAction be done on Excel Saving? (default no)</param>
    Function getParametersFromParamText(paramText As String, paramType As String, ByRef env As Integer, ByRef database As String,
                                   Optional ByRef tableName As String = "", Optional ByRef primKeysStr As String = "",
                                   Optional ByRef insertIfMissing As Boolean = False, Optional ByRef executeAdditionalProc As String = "", Optional ByRef ignoreColumns As String = "",
                                   Optional ByRef execOnSave As Boolean = False, Optional ByRef CUDFlags As Boolean = False) As Boolean

        Dim DBModifParams() As String = functionSplit(paramText, ",", """", "def", "(", ")")
        If IsNothing(DBModifParams) Then Return False
        ' check for completeness
        If DBModifParams.Length < 4 And paramType = "DBMapper" Then
            ErrorMsg("At least environment (can be empty), database, Tablename and primary keys have to be provided as DBMapper parameters !")
            Return False
        End If
        If DBModifParams.Length < 2 And paramType = "DBAction" Then
            ErrorMsg("At least environment (can be empty) and database have to be provided as DBAction parameters !")
            Return False
        End If
        ' fill parameters:
        If DBModifParams(0) <> "" Then env = Convert.ToInt16(DBModifParams(0))
        database = DBModifParams(1).Replace("""", "").Trim
        If database = "" Then
            ErrorMsg("No database given in " & paramType & " comment!")
            Return False
        End If
        If paramType = "DBAction" Then
            execOnSave = False
            If DBModifParams.Length > 2 AndAlso DBModifParams(2) <> "" Then execOnSave = Convert.ToBoolean(DBModifParams(2))
            Return True
        End If
        tableName = DBModifParams(2).Replace("""", "").Trim ' remove all quotes and trim right and left
        If tableName = "" Then
            ErrorMsg("No Tablename given in " & paramType & " comment!")
            Return False
        End If
        primKeysStr = DBModifParams(3).Replace("""", "").Trim
        If primKeysStr = "" Then
            ErrorMsg("No primary keys given in " & paramType & " comment!")
            Return False
        End If
        If DBModifParams.Length > 4 AndAlso DBModifParams(4) <> "" Then insertIfMissing = Convert.ToBoolean(DBModifParams(4))
        If DBModifParams.Length > 5 AndAlso DBModifParams(5) <> "" Then executeAdditionalProc = DBModifParams(5).Replace("""", "").Trim
        If DBModifParams.Length > 6 AndAlso DBModifParams(6) <> "" Then ignoreColumns = DBModifParams(6).Replace("""", "").Trim
        If DBModifParams.Length > 7 AndAlso DBModifParams(7) <> "" Then execOnSave = Convert.ToBoolean(DBModifParams(7))
        If DBModifParams.Length > 8 AndAlso DBModifParams(8) <> "" Then CUDFlags = Convert.ToBoolean(DBModifParams(8))
        Return True
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
        Dim tableName As String = ""                         ' Database Table, where Data is to be stored
        Dim primKeysStr As String = ""                       ' String containing primary Key names for updating table data, comma separated
        Dim database As String = ""                          ' Database to store to
        Dim env As Integer = -1                              ' Environment where connection id should be taken from
        Dim insertIfMissing As Boolean = False               ' if set, then insert row into table if primary key is missing there. Default = False (only update)
        Dim executeAdditionalProc As String = ""             ' additional stored procedure to be executed after saving
        Dim ignoreColumns As String = ""                     ' columns to be ignored (helper columns)
        Dim execOnSave As Boolean = False                    ' should DBMap be saved / DBAction be done on Excel Saving? (default no)
        Dim CUDFlags As Boolean = False                      ' respect C/U/D Flags (DBSheet functionality)
        Dim existingDefinition As Boolean = False
        Dim paramText = getParamTextFromTargetRange(paramTarget:=targetRange, paramType:=type)
        If type = "DBAction" Then
            existingDefinition = getParametersFromParamText(paramText:=paramText, paramType:=type, env:=env, database:=database, execOnSave:=execOnSave)
        ElseIf type = "DBMapper" Then
            existingDefinition = getParametersFromParamText(paramText:=paramText, paramType:=type, env:=env, database:=database, tableName:=tableName, primKeysStr:=primKeysStr, insertIfMissing:=insertIfMissing, executeAdditionalProc:=executeAdditionalProc, ignoreColumns:=ignoreColumns, execOnSave:=execOnSave, CUDFlags:=CUDFlags)
        End If
        ' prepare DBModifier Create Dialog
        Dim theDBModifCreateDlg As DBModifCreate = New DBModifCreate()
        With theDBModifCreateDlg
            .envSel.DataSource = Globals.environdefs
            .envSel.SelectedIndex = -1
            If existingDefinition Then
                If InStr(1, activeCellName, type) > 0 Then .DBModifName.Text = Replace(activeCellName, type, "")
                Try
                    If env > 0 Then .envSel.SelectedIndex = env - 1
                Catch ex As Exception
                    ErrorMsg("Error setting environment " & env & " (correct environment manually in docproperty " & activeCellName & "): " & ex.Message)
                    Exit Sub
                End Try
                .Database.Text = database
                .execOnSave.Checked = execOnSave
                If type = "DBMapper" Then
                    .Tablename.Text = tableName
                    .PrimaryKeys.Text = primKeysStr
                    .insertIfMissing.Checked = insertIfMissing
                    .addStoredProc.Text = executeAdditionalProc
                    .IgnoreColumns.Text = ignoreColumns
                    .CUDflags.Checked = CUDFlags
                End If
                .TargetRangeAddress.Text = targetRange.Parent.Name & "!" & targetRange.Address
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
                            Dim targetRangeName As String = getDBModifNameFromRange(DBModifDefColl(sheetId).Item(nodeName))
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
                                LogError(theFunc & searchCell.Parent.Name & "!" & searchCell.Address & ") has multiple " & IIf(searchCell.Rows.Count > 1, "rows !", "columns !") & ", which leads to problems in DBSequences...")
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

End Module

Public Enum CheckTypeFld
    checkIsNumericFld = 0
    checkIsDateFld = 1
    checkIsTimeFld = 2
    checkIsStringFld = 3
End Enum
