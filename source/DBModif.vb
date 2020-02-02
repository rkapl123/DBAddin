Imports ADODB
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Collections.Generic

''' <summary>Abstraction of a DB Modification Object (concrete classes DB Mapper, DB Action or DB Sequence)</summary>
Public MustInherit Class DBModif

    ''' <summary>needed for field formatting in DB Mapper</summary>
    Protected Enum CheckTypeFld
        checkIsNumericFld = 0
        checkIsDateFld = 1
        checkIsTimeFld = 2
        checkIsStringFld = 3
    End Enum

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
    ''' <summary>the original stored parameters from the definition string</summary>
    Protected DBModifParams() As String
    ''' <summary>ask for confirmation before executtion of DBModif</summary>
    Protected askBeforeExecute As Boolean = True

    Public Function getTargetRangeAddress() As String
        getTargetRangeAddress = targetRangeAddress
    End Function

    Public Function getTargetRange() As Excel.Range
        getTargetRange = TargetRange
    End Function

    Public Function getParamText() As String
        getParamText = paramText
    End Function

    ''' <summary>Environment (integer) where connection id should be taken from (if not existing, take from selectedEnvironment)</summary>
    Protected Function getEnv(Optional defaultEnv As Integer = 0) As Integer
        getEnv = defaultEnv
        If TypeName(Me.GetType()) = "DBSeqnce" Then Throw New NotImplementedException()
        If DBModifParams(0) <> "" Then getEnv = Convert.ToInt16(DBModifParams(0))
    End Function

    ''' <summary>does the actual DB Modification</summary>
    ''' <param name="WbIsSaving"></param>
    ''' <param name="calledByDBSeq"></param>
    Public Overridable Sub doDBModif(Optional WbIsSaving As Boolean = False, Optional calledByDBSeq As String = "")
        Throw New NotImplementedException()
    End Sub

    ''' <summary>is saving needed for this DBModifier</summary>
    ''' <returns>true for saving necessary</returns>
    Public Function DBModifSaveNeeded() As Boolean
        Return execOnSave
    End Function

    ''' <summary>sets the content of the DBModif Create/Edit Dialog</summary>
    ''' <param name="theDBModifCreateDlg"></param>
    Public Overridable Sub setDBModifCreateFields(ByRef theDBModifCreateDlg As DBModifCreate)
        Throw New NotImplementedException()
    End Sub

    ''' <summary>formats theVal to fit the type of record column having data type dataType</summary>
    ''' <param name="theVal"></param>
    ''' <param name="dataType"></param>
    ''' <returns>the formatted value</returns>
    Protected Function dbFormatType(ByVal theVal As Object, dataType As CheckTypeFld) As String
        If dataType = CheckTypeFld.checkIsNumericFld Then ' only decimal points allowed in numeric data
            dbFormatType = Replace(CStr(theVal), ",", ".")
        ElseIf dataType = CheckTypeFld.checkIsDateFld Then
            dbFormatType = "'" & Format(theVal, "yyyy-MM-dd") & "'" ' ISO 8601 standard SQL Date formatting
        ElseIf dataType = CheckTypeFld.checkIsTimeFld Then
            dbFormatType = "'" & Format(theVal, "yyyy-MM-dd HH:mm:ss.fff") & "'" ' ISO 8601 standard SQL Date/time formatting, 24h format...
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
    Protected Function checkIsDateTime(theType As ADODB.DataTypeEnum) As Boolean
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

    Public Sub New(defkey As String, paramDefs As String, paramTarget As Excel.Range)
        dbmapdefkey = defkey
        ' if no target range is set, then no parameters can be found...
        If IsNothing(paramTarget) Then
            Throw New Exception("paramTarget is Nothing")
        End If
        paramTargetName = getDBModifNameFromRange(paramTarget)
        DBmapSheet = paramTarget.Parent.Name
        targetRangeAddress = DBmapSheet + "!" + paramTarget.Address
        If Left(paramTargetName, 8) <> "DBMapper" Then
            Throw New Exception("target " & paramTargetName & " not matching passed DBModif type DBMapper for " & targetRangeAddress & "/" & dbmapdefkey & "!")
        End If
        paramText = paramDefs
        TargetRange = paramTarget

        DBModifParams = functionSplit(paramText, ",", """", "def", "(", ")")
        If IsNothing(DBModifParams) Then Exit Sub
        ' check for completeness
        If DBModifParams.Length < 4 Then
            Throw New Exception("At least environment (can be empty), database, Tablename and primary keys have to be provided as DBMapper parameters !")
        End If

        ' fill parameters:
        database = DBModifParams(1).Replace("""", "").Trim
        If database = "" Then
            Throw New Exception("No database given in DBMapper paramText!")
        End If
        tableName = DBModifParams(2).Replace("""", "").Trim ' remove all quotes and trim right and left
        If tableName = "" Then
            Throw New Exception("No Tablename given in DBMapper paramText!")
        End If
        primKeysStr = DBModifParams(3).Replace("""", "").Trim
        If primKeysStr = "" Then
            Throw New Exception("No primary keys given in DBMapper paramText!")
        End If
        If DBModifParams.Length > 4 AndAlso DBModifParams(4) <> "" Then insertIfMissing = Convert.ToBoolean(DBModifParams(4))
        If DBModifParams.Length > 5 AndAlso DBModifParams(5) <> "" Then executeAdditionalProc = DBModifParams(5).Replace("""", "").Trim
        If DBModifParams.Length > 6 AndAlso DBModifParams(6) <> "" Then ignoreColumns = DBModifParams(6).Replace("""", "").Trim
        If DBModifParams.Length > 7 AndAlso DBModifParams(7) <> "" Then execOnSave = Convert.ToBoolean(DBModifParams(7))
        If DBModifParams.Length > 8 AndAlso DBModifParams(8) <> "" Then CUDFlags = Convert.ToBoolean(DBModifParams(8))
        If DBModifParams.Length > 9 AndAlso DBModifParams(9) <> "" Then askBeforeExecute = Convert.ToBoolean(DBModifParams(9))
    End Sub

    Public Overrides Sub doDBModif(Optional WbIsSaving As Boolean = False, Optional calledByDBSeq As String = "")
        If WbIsSaving And Not execOnSave Then Exit Sub
        If Not WbIsSaving And askBeforeExecute And calledByDBSeq = "" Then
            Dim retval As MsgBoxResult = MsgBox("Really execute DB Mapper " & dbmapdefkey & "?", MsgBoxStyle.Question + vbOKCancel, "Execute DB Mapper")
            If retval = vbCancel Then Exit Sub
        End If
        ' if Environment is not existing, take from selectedEnvironment
        Dim env As Integer = getEnv(Globals.selectedEnvironment + 1)
        ' extend DataRange to whole area ...
        Dim rowEnd = TargetRange.Cells(1, 1).End(Excel.XlDirection.xlDown).Row
        Dim colEnd = TargetRange.Cells(1, 1).End(Excel.XlDirection.xlToRight).Column
        For Each DBModifName As Excel.Name In TargetRange.Parent.Parent.Names
            If DBModifName.Name = paramTargetName Then
                ' then set name to offset function covering the whole area...
                Try
                    DBModifName.RefersTo = TargetRange.Parent.Range(TargetRange.Cells(1, 1), TargetRange.Parent.Cells(rowEnd, colEnd))
                    Exit For
                Catch ex As Exception
                    ErrorMsg("Error when assigning name '" & paramTargetName & "': " & ex.Message)
                    Exit Sub
                End Try
            End If
        Next
        TargetRange = TargetRange.Parent.Range(paramTargetName)
        targetRangeAddress = TargetRange.Parent.Name + "!" + TargetRange.Address

        Dim primKeys() As String = Split(primKeysStr, ",")
        'now create/get a connection (dbcnn) for env(ironment)
        If Not openConnection(env, database) Then Exit Sub

        'checkrst is opened to get information about table schema (field types)
        Dim checkrst As ADODB.Recordset = New ADODB.Recordset
        Dim rst As ADODB.Recordset = New ADODB.Recordset
        Try
            checkrst.Open(tableName, dbcnn, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly, CommandTypeEnum.adCmdTableDirect)
        Catch ex As Exception
            MsgBox("Opening table '" & tableName & "' caused following error: " & ex.Message & " for DBMapper " & paramTargetName, MsgBoxStyle.Critical, "DBMapper Error")
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
            If InStr(1, LCase(ignoreColumns) + ",", LCase(fieldname) + ",") = 0 Then
                Try
                    Dim testExist As String = checkrst.Fields(fieldname).Name
                Catch ex As Exception
                    TargetRange.Parent.Activate
                    TargetRange.Cells(1, colNum).Select
                    MsgBox("Field '" & fieldname & "' does not exist in Table '" & tableName & "' and is not in ignoreColumns, Error in sheet " & TargetRange.Parent.Name, MsgBoxStyle.Critical, "DBMapper Error")
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
            ' if CUDFlags are set, only insert/update/delete if CUDFlags column (right to DBMapper range) is filled...
            If Not CUDFlags Or (CUDFlags And TargetRange.Cells(rowNum, TargetRange.Columns.Count + 1).Value <> "") Then
                Dim rowCUDFlag As String = TargetRange.Cells(rowNum, TargetRange.Columns.Count + 1).Value
                ' try to find record for update, construct where clause with primary key columns
                Dim primKeyCompound As String = " WHERE "
                For i As Integer = 0 To UBound(primKeys)
                    Dim primKeyValue
                    primKeyValue = TargetRange.Cells(rowNum, i + 1).Value
                    primKeyCompound = primKeyCompound & primKeys(i) & " = " & dbFormatType(primKeyValue, checkTypes(i)) & IIf(i = UBound(primKeys), "", " AND ")
                    If IsError(primKeyValue) Then
                        MsgBox("Error in primary key value, cell (" & rowNum & "," & i + 1 & ") in sheet " & TargetRange.Parent.Name & ", & row " & rowNum, MsgBoxStyle.Critical, "DBMapper Error")
                        GoTo nextRow
                    End If
                    If IsNothing(primKeyValue) Then primKeyValue = ""
                    If primKeyValue.ToString().Length = 0 Then
                        MsgBox("Empty primary key value, cell (" & rowNum & "," & i + 1 & ") in sheet " & TargetRange.Parent.Name & ", & row " & rowNum, MsgBoxStyle.Critical, "DBMapper Error")
                        GoTo nextRow
                    End If
                Next
                Dim getStmt As String = "SELECT * FROM " & tableName & primKeyCompound
                Try
                    rst.Open(getStmt, dbcnn, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)
                Catch ex As Exception
                    MsgBox("Problem getting recordset, Error: " & ex.Message & " in sheet " & TargetRange.Parent.Name & ", & row " & rowNum)
                    rst.Close()
                    GoTo cleanup
                End Try

                ' didn't find record, so add a new record if insertIfMissing flag is set
                If rst.EOF Then
                    Dim i As Integer
                    If insertIfMissing Or rowCUDFlag = "i" Then
                        ExcelDnaUtil.Application.StatusBar = "Inserting " & primKeyCompound & " in table " & tableName
                        rst.AddNew()
                        For i = 0 To UBound(primKeys)
                            rst.Fields(primKeys(i)).Value = IIf(TargetRange.Cells(rowNum, i + 1).ToString().Length = 0, vbNull, TargetRange.Cells(rowNum, i + 1).Value)
                        Next
                    Else
                        TargetRange.Parent.Activate
                        TargetRange.Cells(rowNum, i + 1).Select
                        MsgBox("Did not find recordset with statement '" & getStmt & "', insertIfMissing = " & insertIfMissing & " in sheet " & TargetRange.Parent.Name & ", & row " & rowNum, MsgBoxStyle.Critical, "DBMapper Error")
                        rst.Close()
                        GoTo cleanup
                    End If
                Else
                    ExcelDnaUtil.Application.StatusBar = "Updating " & primKeyCompound & " in table " & tableName
                End If

                If Not CUDFlags Or (rowCUDFlag = "u") Then
                    ' walk through non primary columns and fill fields
                    colNum = UBound(primKeys) + 1
                    Do
                        Dim fieldname As String = TargetRange.Cells(1, colNum).Value
                        If InStr(1, LCase(ignoreColumns) + ",", LCase(fieldname) + ",") = 0 Then
                            Try
                                rst.Fields(fieldname).Value = IIf(TargetRange.Cells(rowNum, colNum).ToString().Length = 0, vbNull, TargetRange.Cells(rowNum, colNum).Value)
                            Catch ex As Exception
                                TargetRange.Parent.Activate
                                TargetRange.Cells(rowNum, colNum).Select

                                MsgBox("Field Value Update Error: " & ex.Message & " with Table: " & tableName & ", Field: " & fieldname & ", in sheet " & TargetRange.Parent.Name & ", & row " & rowNum & ", col: " & colNum, MsgBoxStyle.Critical, "DBMapper Error")
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
                        MsgBox("Row Update Error, Table: " & rst.Source & ", Error: " & ex.Message & " in sheet " & TargetRange.Parent.Name & ", & row " & rowNum, MsgBoxStyle.Critical, "DBMapper Error")
                        rst.CancelUpdate()
                        rst.Close()
                        GoTo cleanup
                    End Try
                End If
                If (CUDFlags And rowCUDFlag = "d") Then
                    ExcelDnaUtil.Application.StatusBar = "Deleting " & primKeyCompound & " in table " & tableName
                    rst.Delete(AffectEnum.adAffectCurrent)
                End If
                rst.Close()
nextRow:
                Try
                    finishLoop = IIf(TargetRange.Cells(rowNum + 1, 1).ToString().Length = 0, True, False)
                Catch ex As Exception
                    MsgBox("Error in first primary column: Cells(" & rowNum + 1 & ",1): " & ex.Message, MsgBoxStyle.Critical, "DBMapper Error")
                    'finishLoop = True ' commented to allow erroneous data...
                End Try
            End If
            rowNum += 1
        Loop Until rowNum > TargetRange.Rows.Count Or finishLoop

        ' any additional stored procedures to execute?
        If executeAdditionalProc.Length > 0 Then
            Try
                ExcelDnaUtil.Application.StatusBar = "executing stored procedure " & executeAdditionalProc
                dbcnn.Execute(executeAdditionalProc)
            Catch ex As Exception
                MsgBox("Error in executing additional stored procedure: " & ex.Message, MsgBoxStyle.Critical, "DBMapper Error")
                GoTo cleanup
            End Try
        End If
cleanup:
        ExcelDnaUtil.Application.StatusBar = False
        ' close connection to return it to the pool...
        dbcnn.Close()
    End Sub

    Public Overrides Sub setDBModifCreateFields(ByRef theDBModifCreateDlg As DBModifCreate)
        With theDBModifCreateDlg
            .envSel.SelectedIndex = getEnv() - 1
            .TargetRangeAddress.Text = targetRangeAddress
            .Database.Text = database
            .execOnSave.Checked = execOnSave
            .Tablename.Text = tableName
            .PrimaryKeys.Text = primKeysStr
            .insertIfMissing.Checked = insertIfMissing
            .addStoredProc.Text = executeAdditionalProc
            .IgnoreColumns.Text = ignoreColumns
            .CUDflags.Checked = CUDFlags
            .AskForExecute.Checked = askBeforeExecute
        End With
    End Sub

End Class

Public Class DBAction : Inherits DBModif

    ''' <summary>Database to store to</summary>
    Private database As String
    ''' <summary>DBModif name of target range</summary>
    Private paramTargetName As String

    Public Sub New(defkey As String, paramDefs As String, paramTarget As Excel.Range)
        dbmapdefkey = defkey
        ' if no target range is set, then no parameters can be found...
        If IsNothing(paramTarget) Then Exit Sub
        paramTargetName = getDBModifNameFromRange(paramTarget)
        DBmapSheet = paramTarget.Parent.Name
        targetRangeAddress = DBmapSheet + "!" + paramTarget.Address
        If Left(paramTargetName, 8) <> "DBAction" Then
            MsgBox("target " & paramTargetName & " not matching passed DBModif type DBAction for " & targetRangeAddress & "/" & dbmapdefkey & " !", vbCritical, "DBAction Error")
            Exit Sub
        End If
        ' set up parameters
        If paramTarget.Cells(1, 1).Text = "" Then
            MsgBox("No Action defined in " + paramTargetName + "(" + targetRangeAddress + ")", vbCritical, "DBAction Error")
            Exit Sub
        End If
        paramText = paramDefs
        TargetRange = paramTarget
        DBModifParams = functionSplit(paramText, ",", """", "def", "(", ")")
        If IsNothing(DBModifParams) Then Exit Sub
        ' check for completeness
        If DBModifParams.Length < 2 Then
            MsgBox("At least environment (can be empty) and database have to be provided as DBAction parameters !", vbCritical, "DBAction Error")
            Exit Sub
        End If
        ' fill parameters:
        database = DBModifParams(1).Replace("""", "").Trim
        If database = "" Then
            MsgBox("No database given in DBAction paramText!", vbCritical, "DBAction Error")
            Exit Sub
        End If
        If DBModifParams.Length > 2 AndAlso DBModifParams(2) <> "" Then execOnSave = Convert.ToBoolean(DBModifParams(2))
        If DBModifParams.Length > 3 AndAlso DBModifParams(3) <> "" Then askBeforeExecute = Convert.ToBoolean(DBModifParams(3))
    End Sub

    Public Overrides Sub doDBModif(Optional WbIsSaving As Boolean = False, Optional calledByDBSeq As String = "")
        If WbIsSaving And Not execOnSave Then Exit Sub
        If Not WbIsSaving And askBeforeExecute And calledByDBSeq = "" Then
            Dim retval As MsgBoxResult = MsgBox("Really execute DB Action " & dbmapdefkey & "?", MsgBoxStyle.Question + vbOKCancel, "Execute DB Action")
            If retval = vbCancel Then Exit Sub
        End If
        ' if Environment is not existing, take from selectedEnvironment
        Dim env As Integer = getEnv(Globals.selectedEnvironment + 1)
        'now create/get a connection (dbcnn) for env(ironment)
        If Not openConnection(env, database) Then Exit Sub
        Dim result As Long = 0
        Try
            ExcelDnaUtil.Application.StatusBar = "executing DBAction " & paramTargetName
            dbcnn.Execute(TargetRange.Cells(1, 1).Text, result, Options:=CommandTypeEnum.adCmdText)
        Catch ex As Exception
            MsgBox("Error: " & paramTargetName & ": " & ex.Message, vbCritical, "DBAction Error")
            Exit Sub
        End Try
        If Not WbIsSaving And calledByDBSeq = "" Then
            MsgBox("DBAction " & paramTargetName & " executed, affected records: " & result)
        End If
        ' close connection to return it to the pool...
        ExcelDnaUtil.Application.StatusBar = False
        dbcnn.Close()
    End Sub

    Public Overrides Sub setDBModifCreateFields(ByRef theDBModifCreateDlg As DBModifCreate)
        With theDBModifCreateDlg
            .envSel.SelectedIndex = getEnv() - 1
            .TargetRangeAddress.Text = targetRangeAddress
            .Database.Text = database
            .execOnSave.Checked = execOnSave
            .AskForExecute.Checked = askBeforeExecute
        End With
    End Sub
End Class

Public Class DBSeqnce : Inherits DBModif

    ''' <summary>sequence of DB Mappers, DB Actions and DB Refreshes being executed in this sequence</summary>
    Private sequenceParams() As String = {}

    Public Sub New(defkey As String, DBSequenceText As String)
        dbmapdefkey = defkey
        paramText = DBSequenceText
        If paramText = "" Then
            MsgBox("No Sequence defined in " + dbmapdefkey, vbCritical, "DB Sequence Error")
            Exit Sub
        End If
        ' parse parameters: 1st item is execOnSave, 2nd askBeforeExecute, rest defines sequence (tripletts of DBModifType:DBModifName)
        DBModifParams = Split(paramText, ",")
        execOnSave = Convert.ToBoolean(DBModifParams(0)) ' should DBSequence be done on Excel Saving?
        If Boolean.TryParse(value:=DBModifParams(1), result:=askBeforeExecute) Then
            ReDim sequenceParams(DBModifParams.Length() - 3)
            Array.Copy(DBModifParams, 2, sequenceParams, 0, DBModifParams.Length() - 2)
        Else
            ReDim sequenceParams(DBModifParams.Length() - 2)
            Array.Copy(DBModifParams, 1, sequenceParams, 0, DBModifParams.Length() - 1)
        End If
    End Sub

    Public Overrides Sub doDBModif(Optional WbIsSaving As Boolean = False, Optional calledByDBSeq As String = "")
        If WbIsSaving And Not execOnSave Then Exit Sub
        If Not WbIsSaving And askBeforeExecute Then
            Dim retval As MsgBoxResult = MsgBox("Really execute DB Sequence " & dbmapdefkey & "?", MsgBoxStyle.Question + vbOKCancel, "Execute DB Sequence")
            If retval = vbCancel Then Exit Sub
        End If

        For i As Integer = 0 To UBound(sequenceParams)
            Dim definition() As String = Split(sequenceParams(i), ":")
            If definition(0) <> "DBRefrsh" Then
                DBModifDefColl(definition(0)).Item(definition(1)).doDBModif(WbIsSaving, calledByDBSeq:=dbmapdefkey)
            Else
                ' reset query cache, so we really get new data !
                Functions.queryCache.Clear()
                Functions.StatusCollection.Clear()
                ' refresh DBFunction in sequence
                Dim underlyingName As String = definition(1)
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
        With theDBModifCreateDlg
            .RepairDBSeqnce.Text = paramText
            .execOnSave.Checked = execOnSave
            .AskForExecute.Checked = askBeforeExecute
            For i As Integer = 0 To UBound(sequenceParams)
                .DBSeqenceDataGrid.Rows.Add(sequenceParams(i))
            Next
        End With
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
            MsgBox("No Connectionstring given for environment: " & env & ", please correct and rerun.", vbCritical, "Open Connection Error")
            Exit Function
        End If
        Dim dbidentifier As String = fetchSetting("DBidentifierCCS" & env, String.Empty)
        If dbidentifier = String.Empty Then
            MsgBox("No DB identifier given for environment: " & env & ", please correct and rerun.", vbCritical, "Open Connection Error")
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
            MsgBox("Error connecting to DB: " & ex.Message & ", connection string: " & theConnString, vbCritical, "Open Connection Error")
            If dbcnn.State = ADODB.ObjectStateEnum.adStateOpen Then dbcnn.Close()
            dbcnn = Nothing
        End Try
        ExcelDnaUtil.Application.StatusBar = String.Empty
        openConnection = True
    End Function

    ''' <summary>creates a DBModif at the current active cell or edits an existing one defined in targetDefName (after being called in defined range or from ribbon + Ctrl + Shift)</summary>
    Sub createDBModif(type As String, Optional targetDefName As String = "")
        ' clipboard helper for legacy definitions:
        ' if saveRangeToDB macro calls were copied, rename saveRangeToDB<Single> To def, 1st parameter (datarange) removed (empty), connid moved to 2nd place as database name (remove MSSQL)
        'mapper.saveRangeToDBSingle(Range("DB_DefName"), "tableName", "primKey1,primKey2,primKey3", "MSSQLDB_NAME", True)
        '--> def(, "DB_NAME", "tableName", "primKey1,primKey2,primKey3", True)    DBMapperName = DB_DefName
        'mapper.saveRangeToDBSingle(Range("DB_DefName"), "tableName", "primKey1,primKey2,primKey3", "MSSQLDB_NAME")
        '--> def(, "DB_NAME", "tableName", "primKey1,primKey2,primKey3")          DBMapperName = DB_DefName
        'def(, "DB_NAME", True), "tableName", "primKey1,primKey2,primKey3", "MSSQLDB_NAME", True)", "MSSQLDB_NAME", True)
        Dim existingDBModif As DBModif = Nothing
        Dim activeCellName As String = targetDefName
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
                ' assign new name to active cell
                ' Add doesn't work directly with ExcelDnaUtil.Application.ActiveWorkbook.Names (late binding), so create an object here...
                Dim NamesList As Excel.Names = ExcelDnaUtil.Application.ActiveWorkbook.Names
                Try : NamesList.Add(Name:=DB_DefName, RefersTo:=ExcelDnaUtil.Application.ActiveCell)
                Catch ex As Exception
                    MsgBox("Error when assigning name '" & DB_DefName & "' to active cell: " & ex.Message, vbCritical, "DBMapper Legacy Creation Error")
                    Exit Sub
                End Try
                ' store parameters in same named docproperty
                Try : ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties(DB_DefName).Delete : Catch ex As Exception : End Try
                Try
                    ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties.Add(Name:=DB_DefName, LinkToContent:=False, Type:=Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, Value:=newDefString)
                Catch ex As Exception
                    MsgBox("Error when adding CustomDocumentProperty with DBModif parameters (Name:" & DB_DefName & ",content: " & newDefString & "): " & ex.Message, vbCritical, "DBMapper Legacy Creation Error")
                    ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties(DB_DefName).Delete
                    Exit Sub
                End Try
                existingDBModif = New DBMapper(defkey:=DB_DefName, paramDefs:=newDefString, paramTarget:=ExcelDnaUtil.Application.ActiveCell)
                activeCellName = DB_DefName
                createdDBMapperFromClipboard = True
                Clipboard.Clear()
            End If
        End If

        ' for DBMappers defined in ListObjects, try potential name to ListObject (parts), only possible if manually defined !
        If type = "DBMapper" Then
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

        ' fetch parameters if there is an existing definition...
        If DBModifDefColl.ContainsKey(type) AndAlso DBModifDefColl(type).ContainsKey(targetDefName) Then existingDBModif = DBModifDefColl(type).Item(targetDefName)

        ' prepare DBModifier Create Dialog
        Dim theDBModifCreateDlg As DBModifCreate = New DBModifCreate()
        With theDBModifCreateDlg
            ' store DBModification type in tag for validation purposes...
            .Tag = type
            .envSel.DataSource = Globals.environdefs
            .envSel.SelectedIndex = -1
            .DBModifName.Text = Replace(activeCellName, type, "")
            .RepairDBSeqnce.Hide()
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
                .AskForExecute.Top = .TargetRangeLabel.Top
                ' fill Datagridview for DBSequence
                Dim cb As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn()
                cb.HeaderText = "Sequence Step"
                cb.ReadOnly = False
                cb.ValueType() = GetType(String)
                Dim ds As List(Of String) = New List(Of String)

                ' first add the DBMapper and DBAction definitions available in the Workbook
                For Each DBModiftype As String In DBModifDefColl.Keys
                    ' avoid DB Sequences (might be - indirectly - self referencing, leading to endless recursion)
                    If DBModiftype <> "DBSeqnce" Then
                        For Each nodeName As String In DBModifDefColl(DBModiftype).Keys
                            ds.Add(DBModiftype & ":" & nodeName)
                        Next
                    End If
                Next

                ' then add DBRefresh items for allowing refreshing DBFunctions (DBListFetch and DBSetQuery) during a Sequence
                Dim searchCell As Excel.Range
                For Each ws As Excel.Worksheet In ExcelDnaUtil.Application.ActiveWorkbook.Worksheets
                    For Each theFunc As String In {"DBListFetch(", "DBSetQuery("}
                        searchCell = ws.Cells.Find(What:=theFunc, After:=ws.Range("A1"), LookIn:=Excel.XlFindLookIn.xlFormulas, LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlNext, MatchCase:=False)
                        If Not (searchCell Is Nothing) Then
                            If searchCell.Rows.Count > 1 Or searchCell.Rows.Count > 1 Then
                                LogError(theFunc & " (in " & searchCell.Parent.Name & "!" & searchCell.Address & ") has multiple " & IIf(searchCell.Rows.Count > 1, "rows !", "columns !") & ", which leads to problems in DBSequences...")
                                Continue For
                            End If
                            Dim underlyingName As String = getDBunderlyingNameFromRange(searchCell)
                            ds.Add("DBRefrsh:" & underlyingName & ":" & theFunc & "in " & searchCell.Parent.Name & "!" & searchCell.Address & ")")
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
                .DBSeqenceDataGrid.Columns(0).Width = 200
            Else
                ' hide controls irrelevant for DBMapper and DBAction
                .up.Hide()
                .down.Hide()
                .DBSeqenceDataGrid.Hide()
            End If

            ' delegate filling of dialog fields to created DBModif object
            If Not IsNothing(existingDBModif) Then existingDBModif.setDBModifCreateFields(theDBModifCreateDlg)

            ' display dialog for parameters
            If theDBModifCreateDlg.ShowDialog() = DialogResult.Cancel Then
                ' remove targetRange Name and customdocproperty created in clipboard helper
                If createdDBMapperFromClipboard Then
                    Try
                        ExcelDnaUtil.Application.ActiveWorkbook.Names(activeCellName).Delete
                        ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties(activeCellName).Delete
                    Catch ex As Exception : End Try
                End If
                Exit Sub
            End If

            ' only for DBMapper or DBAction: change target range name
            If type <> "DBSeqnce" Then
                Dim targetRange As Excel.Range = existingDBModif.getTargetRange()
                If IsNothing(targetRange) Then targetRange = ExcelDnaUtil.Application.ActiveCell
                ' set content range name: first clear name
                Try : ExcelDnaUtil.Application.ActiveWorkbook.Names(activeCellName).Delete : Catch ex As Exception : End Try
                ' then (re)set name
                Dim NamesList As Excel.Names = ExcelDnaUtil.Application.ActiveWorkbook.Names
                Try
                    NamesList.Add(Name:=type + .DBModifName.Text, RefersTo:=targetRange)
                Catch ex As Exception : ErrorMsg("Error when assigning name '" & type & .DBModifName.Text & "' to active cell: " & ex.Message)
                End Try
            End If

            ' create parameter definition string ...
            Dim newParamText As String = ""
            If type = "DBAction" Then
                newParamText = "def(" + IIf(.envSel.SelectedIndex = -1, "", (.envSel.SelectedIndex + 1).ToString()) + "," + """" + .Database.Text + """," +
                    .execOnSave.Checked.ToString() + "," + .AskForExecute.Checked.ToString() + ")"
            ElseIf type = "DBMapper" Then
                newParamText = "def(" +
                    IIf(.envSel.SelectedIndex = -1, "", (.envSel.SelectedIndex + 1).ToString()) + "," +
                    """" + .Database.Text + """," + """" + .Tablename.Text + """," + """" + .PrimaryKeys.Text + """," + .insertIfMissing.Checked.ToString() + "," +
                    """" + IIf(Len(.addStoredProc.Text) = 0, "", .addStoredProc.Text) + """," +
                    """" + IIf(Len(.IgnoreColumns.Text) = 0, "", .IgnoreColumns.Text) + """," +
                    .execOnSave.Checked.ToString() + "," + .CUDflags.Checked.ToString() + "," + .AskForExecute.Checked.ToString() + ")"
            ElseIf type = "DBSeqnce" Then
                If .Tag = "repaired" Then
                    newParamText = .RepairDBSeqnce.Text
                Else
                    newParamText = .execOnSave.Checked.ToString() + "," + .AskForExecute.Checked.ToString()
                    ' need that because empty row at the end is passed along with Rows() !!
                    For i As Integer = 0 To .DBSeqenceDataGrid.Rows().Count - 2
                        newParamText += "," + .DBSeqenceDataGrid.Rows(i).Cells(0).Value
                    Next

                End If
            End If
            ' ... and store in docproperty (rename docproperty first to current name, might have been changed)
            Try : ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties(activeCellName).Delete : Catch ex As Exception : End Try
            Try
                ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties.Add(Name:=type + .DBModifName.Text, LinkToContent:=False, Type:=Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, Value:=newParamText)
            Catch ex As Exception : MsgBox("Error when adding property with DBModif parameters: " & ex.Message, vbCritical, "DBModif Creation Error") : End Try
        End With
        ' refresh mapper definitions to reflect changes immediately...
        getDBModifDefinitions()
    End Sub

    ''' <summary>gets defined names for DBModifier (DBMapper/DBAction/DBSeqnce) invocation in the current workbook and updates Ribbon with it</summary>
    Public Sub getDBModifDefinitions()
        ' load DBModifier definitions (objects) into Global collection DBModifDefColl
        Try
            Globals.DBModifDefColl = New Dictionary(Of String, Dictionary(Of String, DBModif))
            ' get DBModifier definitions from docproperties
            For Each docproperty In ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties
                Dim DBModiftype As String = Left(docproperty.Name, 8)
                If TypeName(docproperty.Value) = "String" And (DBModiftype = "DBSeqnce" Or DBModiftype = "DBMapper" Or DBModiftype = "DBAction") Then
                    Dim nodeName As String = docproperty.Name
                    Dim targetRange As Excel.Range = Nothing
                    ' for DBMappers and DBActions the data of the DBModification is stored in Ranges, so check for those and get the Range
                    If DBModiftype = "DBMapper" Or DBModiftype = "DBAction" Then
                        For Each rangename As Excel.Name In ExcelDnaUtil.Application.ActiveWorkbook.Names
                            Dim rangenameName As String = Replace(rangename.Name, rangename.Parent.Name & "!", "")
                            If rangenameName = nodeName And InStr(rangename.RefersTo, "#REF!") > 0 Then
                                MsgBox(DBModiftype + " definitions range [" + rangename.Parent.Name + "]" + rangename.Name + " contains #REF!", vbCritical, "DBModifier Definitions Error")
                                Exit For
                            ElseIf rangenameName = nodeName Then
                                targetRange = rangename.RefersToRange
                                Exit For
                            End If
                        Next
                        If IsNothing(targetRange) Then
                            MsgBox("Error, required target range named '" & nodeName & "' not existing for " & DBModiftype & "." & vbCrLf & "either create target range or delete docproperty named  '" & nodeName & "' !", vbCritical, "DBModifier Definitions Error")
                            Continue For
                        End If
                    End If
                    ' finally create the DBModif Object ...
                    Dim newDBModif As DBModif
                    If DBModiftype = "DBMapper" Then
                        newDBModif = New DBMapper(docproperty.Name, docproperty.Value, targetRange)
                    ElseIf DBModiftype = "DBAction" Then
                        newDBModif = New DBAction(docproperty.Name, docproperty.Value, targetRange)
                    ElseIf DBModiftype = "DBSeqnce" Then
                        newDBModif = New DBSeqnce(docproperty.Name, docproperty.Value)
                    Else
                        MsgBox("Error, not supported DBModiftype: " & DBModiftype, vbCritical, "DBModifier Definitions Error")
                        newDBModif = Nothing
                    End If
                    ' ... and add it to the collection DBModifDefColl
                    Dim defColl As Dictionary(Of String, DBModif) ' definition lookup collection for DBModifiername -> object
                    If Not DBModifDefColl.ContainsKey(DBModiftype) Then
                        ' add to new DBModiftype "menu"
                        defColl = New Dictionary(Of String, DBModif)
                        defColl.Add(docproperty.Name, newDBModif)
                        DBModifDefColl.Add(DBModiftype, defColl)
                    Else
                        ' add definition to existing DBModiftype "menu"
                        defColl = DBModifDefColl(DBModiftype)
                        defColl.Add(docproperty.Name, newDBModif)
                    End If
                End If
            Next
            Globals.theRibbon.Invalidate()
        Catch ex As Exception
            LogError("Error: " & ex.Message)
        End Try
    End Sub

    ''' <summary>gets DB Modification Name (DBMapper or DBAction) from theRange</summary>
    ''' <param name="theRange"></param>
    ''' <returns>the retrieved name as a string (not name object !)</returns>
    Public Function getDBModifNameFromRange(theRange As Excel.Range) As String
        Dim nm As Excel.Name
        Dim rng, testRng As Excel.Range

        getDBModifNameFromRange = ""
        Try
            ' try all names in workbook
            For Each nm In theRange.Parent.Parent.Names
                rng = Nothing
                ' test whether range referring to that name (if it is a real range)...
                Try : rng = nm.RefersToRange : Catch ex As Exception : End Try
                If Not rng Is Nothing Then
                    testRng = Nothing
                    ' ...intersects with the passed range
                    Try : testRng = ExcelDnaUtil.Application.Intersect(theRange, rng) : Catch ex As Exception : End Try
                    If Not IsNothing(testRng) And (InStr(1, nm.Name, "DBMapper") >= 1 Or InStr(1, nm.Name, "DBAction") >= 1) Then
                        ' and pass back the name if it does and is a DBMapper or a DBAction
                        getDBModifNameFromRange = nm.Name
                        Exit Function
                    End If
                End If
            Next
        Catch ex As Exception
            LogError("Error: " & ex.Message)
        End Try
    End Function

End Module

