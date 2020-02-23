Imports ADODB
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports Microsoft.Office.Core

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
    'TODO: remove this after legacy migration...
    '''<summary>parameter text for DBModif (def(...)</summary>
    Protected paramText As String
    '''<summary>should DBMap be saved / DBAction be done on Excel Saving? (default no)</summary>
    Protected execOnSave As Boolean = False
    'TODO: remove this after legacy migration...
    ''' <summary>the original stored parameters from the definition string</summary>
    Protected DBModifParams() As String
    ''' <summary>ask for confirmation before executtion of DBModif</summary>
    Protected askBeforeExecute As Boolean = True
    ''' <summary>environment specific for the DBModif object, if left empty then set to default environment (either 0 or currently selected environment)</summary>
    Protected env As String = ""

    ''' <summary>public accessor function</summary>
    ''' <returns>the targetRangeAddress</returns>
    Public Function getTargetRangeAddress() As String
        getTargetRangeAddress = targetRangeAddress
    End Function

    ''' <summary>public accessor function</summary>
    ''' <returns>the targetRange itself</returns>
    Public Function getTargetRange() As Excel.Range
        getTargetRange = TargetRange
    End Function

    ''' <summary>public accessor function: get Environment (integer) where connection id should be taken from (if not existing, take from selectedEnvironment being passed in defaultEnv)</summary>
    ''' <param name="defaultEnv">optionally passed selected Environment</param>
    ''' <returns>the Environment of the DBModif</returns>
    Protected Function getEnv(Optional defaultEnv As Integer = 0) As Integer
        getEnv = defaultEnv
        If TypeName(Me.GetType()) = "DBSeqnce" Then Throw New NotImplementedException()
        If env <> "" Then getEnv = Convert.ToInt16(env)
    End Function

    ''' <summary>get the DBModifparameters for legacy conversion in DBModifs.getDBModifDefinitions</summary>
    ''' <returns>the DBModif parameters</returns>
    Public Function GetDBModifParams() As String()
        Return DBModifParams
    End Function

    ''' <summary>is saving needed for this DBModifier</summary>
    ''' <returns>true for saving necessary</returns>
    Public Function DBModifSaveNeeded() As Boolean
        Return execOnSave
    End Function

    ''' <summary>does the actual DB Modification</summary>
    ''' <param name="WbIsSaving"></param>
    ''' <param name="calledByDBSeq"></param>
    Public Overridable Sub doDBModif(Optional WbIsSaving As Boolean = False, Optional calledByDBSeq As String = "")
        Throw New NotImplementedException()
    End Sub

    ''' <summary>sets the content of the DBModif Create/Edit Dialog</summary>
    ''' <param name="theDBModifCreateDlg"></param>
    Public Overridable Sub setDBModifCreateFields(ByRef theDBModifCreateDlg As DBModifCreate)
        Throw New NotImplementedException()
    End Sub

    ''' <summary>when resizing target ranges from functions as DBListFetch and DBSetQuery, need to notify also DBModif objects (DBMapper)</summary>
    ''' <param name="newTargetRange"></param>
    Public Sub setTargetRange(newTargetRange As Excel.Range)
        TargetRange = newTargetRange
    End Sub

    ''' <summary>formats theVal to fit the type of record column having data type dataType</summary>
    ''' <param name="theVal"></param>
    ''' <param name="dataType"></param>
    ''' <returns>the formatted value</returns>
    Protected Function dbFormatType(ByVal theVal As Object, dataType As CheckTypeFld) As String
        If IsNothing(theVal) Then
            dbFormatType = "NULL"
        ElseIf dataType = CheckTypeFld.checkIsNumericFld Then ' only decimal points allowed in numeric data
            dbFormatType = Replace(CStr(theVal), ",", ".")
        ElseIf dataType = CheckTypeFld.checkIsDateFld Then
            dbFormatType = "'" & Format(Date.FromOADate(theVal), "yyyy-MM-dd") & "'" ' ISO 8601 standard SQL Date formatting
        ElseIf dataType = CheckTypeFld.checkIsTimeFld Then
            dbFormatType = "'" & Format(Date.FromOADate(theVal), "yyyy-MM-dd HH:mm:ss.fff") & "'" ' ISO 8601 standard SQL Date/time formatting, 24h format...
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
    Protected Function checkIsDate(theType As ADODB.DataTypeEnum) As Boolean
        checkIsDate = False
        If theType = ADODB.DataTypeEnum.adDate Or theType = ADODB.DataTypeEnum.adDBDate Then
            checkIsDate = True
        End If
    End Function

    ''' <summary>checks whether ADO type theType is a time type</summary>
    ''' <param name="theType"></param>
    ''' <returns>True if Time</returns>
    Protected Function checkIsTime(theType As ADODB.DataTypeEnum) As Boolean
        checkIsTime = False
        If theType = ADODB.DataTypeEnum.adDBTime Or theType = ADODB.DataTypeEnum.adDBTimeStamp Then
            checkIsTime = True
        End If
    End Function

    ''' <summary>checks whether ADO type theType is a numeric type</summary>
    ''' <param name="theType"></param>
    ''' <returns>True if numeric</returns>
    Protected Function checkIsNumeric(theType As ADODB.DataTypeEnum) As Boolean
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

    ''' <summary>legacy constructor for mapping existing DBMapper macro calls (copy in clipboard)</summary>
    ''' <param name="defkey"></param>
    ''' <param name="paramDefs"></param>
    ''' <param name="paramTarget"></param>
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

        'TODO: change this after legacy migration...
        'Dim DBModifParams() As String = functionSplit(paramText, ",", """", "def", "(", ")")
        DBModifParams = functionSplit(paramText, ",", """", "def", "(", ")")
        If IsNothing(DBModifParams) Then Exit Sub
        ' check for completeness
        If DBModifParams.Length < 4 Then
            Throw New Exception("At least environment (can be empty), database, Tablename and primary keys have to be provided as DBMapper parameters !")
        End If

        ' fill parameters:
        env = DBModifParams(0)
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

    Public Sub New(definitionXML As CustomXMLNode, paramTarget As Excel.Range)
        dbmapdefkey = definitionXML.BaseName
        ' if no target range is set, then no parameters can be found...
        If IsNothing(paramTarget) Then Throw New Exception("paramTarget is Nothing")
        paramTargetName = getDBModifNameFromRange(paramTarget)
        DBmapSheet = paramTarget.Parent.Name
        targetRangeAddress = DBmapSheet + "!" + paramTarget.Address
        If Left(paramTargetName, 8) <> "DBMapper" Then Throw New Exception("target " & paramTargetName & " not matching passed DBModif type DBMapper for " & targetRangeAddress & "/" & dbmapdefkey & "!")
        TargetRange = paramTarget

        ' fill parameters from definition
        env = definitionXML.SelectSingleNode("ns0:env").Text
        database = definitionXML.SelectSingleNode("ns0:database").Text
        If database = "" Then Throw New Exception("No database given in DBMapper definition!")
        tableName = definitionXML.SelectSingleNode("ns0:tableName").Text
        If tableName = "" Then Throw New Exception("No Tablename given in DBMapper definition!")
        primKeysStr = definitionXML.SelectSingleNode("ns0:primKeysStr").Text
        If primKeysStr = "" Then Throw New Exception("No primary keys given in DBMapper definition!")
        execOnSave = Convert.ToBoolean(definitionXML.SelectSingleNode("ns0:execOnSave").Text)
        askBeforeExecute = Convert.ToBoolean(definitionXML.SelectSingleNode("ns0:askBeforeExecute").Text)
        insertIfMissing = Convert.ToBoolean(definitionXML.SelectSingleNode("ns0:insertIfMissing").Text)
        executeAdditionalProc = definitionXML.SelectSingleNode("ns0:executeAdditionalProc").Text
        ignoreColumns = definitionXML.SelectSingleNode("ns0:ignoreColumns").Text
        CUDFlags = Convert.ToBoolean(definitionXML.SelectSingleNode("ns0:CUDFlags").Text)
    End Sub

    ''' <summary>inserts CUD (Create/Update/Delete) Marks at the right end of the DBMapper range</summary>
    ''' <param name="changedRange">passed TargetRange by Change Event or delete button</param>
    ''' <param name="deleteFlag">if delete button was pressed, this is true</param>
    Public Sub doCUDMarks(changedRange As Excel.Range, Optional deleteFlag As Boolean = False)
        If Not CUDFlags Then Exit Sub
        ExcelDnaUtil.Application.AutoCorrect.AutoExpandListRange = False ' to prevent automatic creation of new column
        ' DBMapper ranges always have a header row, so changedRange.Row - 1...
        If deleteFlag Then
            For Each changedRow As Excel.Range In changedRange.Rows
                TargetRange.Cells(changedRow.Row - 1, TargetRange.Columns.Count + 1).Value = "d"
                TargetRange.Rows(changedRow.Row - 1).Font.Strikethrough = True
            Next
        Else
            For Each changedRow As Excel.Range In changedRange.Rows
                ' change only if not already set
                If TargetRange.Cells(changedRow.Row - 1, TargetRange.Columns.Count + 1).Value = "" Then
                    Dim RowContainsData As Boolean = False
                    For Each containedCell As Excel.Range In TargetRange.Rows(changedRow.Row - 1).Cells
                        ' check if whole row is empty (except for the changedRange)
                        If Not IsNothing(containedCell.Value) AndAlso containedCell.Address <> changedRange.Address Then
                            RowContainsData = True
                            Exit For
                        End If
                    Next
                    ' if Row Contains Data (not every cell empty except currently modified (changedRange), this is for adding rows below data range) then "u"pdate
                    If RowContainsData Then
                        TargetRange.Cells(changedRow.Row - 1, TargetRange.Columns.Count + 1).Value = "u"
                        TargetRange.Rows(changedRow.Row - 1).Font.Italic = True
                    Else ' else "i"nsert
                        TargetRange.Cells(changedRow.Row - 1, TargetRange.Columns.Count + 1).Value = "i"
                    End If
                End If
            Next
        End If
        ExcelDnaUtil.Application.AutoCorrect.AutoExpandListRange = True
    End Sub

    ''' <summary>extend DataRange to "whole" DBMApper area (first row (field names) to the right and first column (primary key) down)</summary>
    ''' <param name="ignoreCUDFlag"></param>
    Public Sub extendDataRange(Optional ignoreCUDFlag As Boolean = False)
        ' only extend if no CUD Flags present (may have non existing first (primary) columns -> auto identity columns !)
        If Not CUDFlags And Not ignoreCUDFlag Then
            preventChangeWhileFetching = True
            Dim rowEnd = TargetRange.Cells(1, 1).End(Excel.XlDirection.xlDown).Row
            Dim colEnd = TargetRange.Cells(1, 1).End(Excel.XlDirection.xlToRight).Column
            Dim NamesList As Excel.Names = ExcelDnaUtil.Application.ActiveWorkbook.Names
            Try : NamesList.Add(Name:=paramTargetName, RefersTo:=TargetRange.Parent.Range(TargetRange.Cells(1, 1), TargetRange.Parent.Cells(rowEnd, colEnd)))
            Catch ex As Exception
                Throw New Exception("Error when reassigning name '" & paramTargetName & "' to DBMapper while extending DataRange: " & ex.Message)
            Finally
                preventChangeWhileFetching = False
            End Try
        End If
        ' even if CUD Flags are present, the Data range might have been extended (by inserting rows), so reassign it to the TargetRange
        TargetRange = TargetRange.Parent.Range(paramTargetName)
        targetRangeAddress = TargetRange.Parent.Name + "!" + TargetRange.Address
    End Sub

    ''' <summary>reset CUD FLags, either after completion of doDBModif or on request (refresh)</summary>
    Public Sub resetCUDFlags()
        If CUDFlags Then
            ExcelDnaUtil.Application.AutoCorrect.AutoExpandListRange = False ' to prevent automatic creation of new column
            TargetRange.Columns(TargetRange.Columns.Count + 1).ClearContents
            ExcelDnaUtil.Application.AutoCorrect.AutoExpandListRange = True ' to prevent automatic creation of new column
        End If
    End Sub

    Public Overrides Sub doDBModif(Optional WbIsSaving As Boolean = False, Optional calledByDBSeq As String = "")
        If WbIsSaving And Not execOnSave Then Exit Sub
        If Not WbIsSaving And askBeforeExecute And calledByDBSeq = "" Then
            Dim retval As MsgBoxResult = MsgBox("Really execute DB Mapper " & dbmapdefkey & "?", MsgBoxStyle.Question + vbOKCancel, "Execute DB Mapper")
            If retval = vbCancel Then Exit Sub
        End If
        ' if Environment is not existing, take from selectedEnvironment
        Dim env As Integer = getEnv(Globals.selectedEnvironment + 1)
        extendDataRange()

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
            Dim rowCUDFlag As String = TargetRange.Cells(rowNum, TargetRange.Columns.Count + 1).Value
            If Not CUDFlags Or (CUDFlags And rowCUDFlag <> "") Then

                ' try to find record for update, construct where clause with primary key columns
                Dim primKeyCompound As String = " WHERE "
                Dim primKeyDisplay As String
                For i As Integer = 0 To UBound(primKeys)
                    Dim primKeyValue
                    If primKeys(i).ToUpper <> TargetRange.Cells(1, i + 1).Value.ToString.ToUpper Then
                        MsgBox("Defined primary key " & primKeys(i) & " does not match primary key " & TargetRange.Cells(1, i + 1).Value & " in DBMapper Data Range, cell (1," & i + 1 & ") !" & vbCrLf & "Primary keys have to be defined in the same order as in the Data Range", MsgBoxStyle.Critical, "DBMapper Error")
                        GoTo cleanup
                    End If
                    primKeyValue = TargetRange.Cells(rowNum, i + 1).Value
                    primKeyCompound = primKeyCompound & primKeys(i) & " = " & dbFormatType(primKeyValue, checkTypes(i)) & IIf(i = UBound(primKeys), "", " AND ")
                    If IsError(primKeyValue) Then
                        MsgBox("Error in primary key value, cell (" & rowNum & "," & i + 1 & ") in sheet " & TargetRange.Parent.Name & ", & row " & rowNum, MsgBoxStyle.Critical, "DBMapper Error")
                        GoTo nextRow
                    End If
                    If IsNothing(primKeyValue) Then primKeyValue = ""
                    ' with CUDFlags there can be empty primary keys (auto identity columns), leave error checking to database in this case ...
                    If (Not CUDFlags Or (CUDFlags And rowCUDFlag = "u")) And primKeyValue.ToString().Length = 0 Then
                        MsgBox("Empty primary key value, cell (" & rowNum & "," & i + 1 & ") in sheet " & TargetRange.Parent.Name & ", & row " & rowNum, MsgBoxStyle.Critical, "DBMapper Error")
                        GoTo nextRow
                    End If
                Next
                Dim getStmt As String = "SELECT * FROM " & tableName & primKeyCompound
                Try
                    rst.Open(getStmt, dbcnn, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)
                    Dim check As Boolean = rst.EOF
                Catch ex As Exception
                    MsgBox("Problem getting recordset, Error: " & ex.Message & " in sheet " & TargetRange.Parent.Name & " and row " & rowNum & ", doing " & getStmt)
                    GoTo cleanup
                End Try
                primKeyDisplay = Replace(Mid(primKeyCompound, 7), " AND ", ";")

                ' If we didn't find record, add a new record if insertIfMissing flag is set or CUD Flag insert is given
                If rst.EOF Then
                    Dim i As Integer
                    If insertIfMissing Or rowCUDFlag = "i" Then
                        ExcelDnaUtil.Application.StatusBar = Left("Inserting " & primKeyDisplay & " into " & tableName, 255)
                        rst.AddNew()
                        For i = 0 To UBound(primKeys)
                            Try
                                ' ignore empty primary field values for identity fields (error message from DB later)..
                                If Not (IsNothing(TargetRange.Cells(rowNum, i + 1).Value) OrElse TargetRange.Cells(rowNum, i + 1).Value.ToString().Length = 0) Then
                                    rst.Fields(primKeys(i)).Value = TargetRange.Cells(rowNum, i + 1).Value
                                End If
                            Catch ex As Exception
                                MsgBox("Error inserting primary key value into table " & tableName & ": " & dbcnn.Errors(0).Description, MsgBoxStyle.Critical, "DBMapper Error")
                            End Try
                        Next
                    Else
                        TargetRange.Parent.Activate
                        TargetRange.Cells(rowNum, i + 1).Select
                        MsgBox("Did not find recordset with statement '" & getStmt & "', insertIfMissing = " & insertIfMissing & " in sheet " & TargetRange.Parent.Name & ", & row " & rowNum, MsgBoxStyle.Critical, "DBMapper Error")
                        GoTo cleanup
                    End If
                Else
                    ExcelDnaUtil.Application.StatusBar = Left("Updating " & primKeyDisplay & " in " & tableName, 255)
                End If

                If Not CUDFlags Or (CUDFlags And (rowCUDFlag = "i" Or rowCUDFlag = "u")) Then
                    ' walk through non primary columns and fill fields
                    colNum = primKeys.Length() + 1
                    Do
                        Dim fieldname As String = TargetRange.Cells(1, colNum).Value
                        If InStr(1, LCase(ignoreColumns) + ",", LCase(fieldname) + ",") = 0 Then
                            Try
                                Dim fieldval As Object = TargetRange.Cells(rowNum, colNum).Value
                                If Not IsNothing(fieldval) Then
                                    rst.Fields(fieldname).Value = IIf(fieldval.ToString().Length = 0, vbNull, fieldval)
                                End If
                            Catch ex As Exception
                                TargetRange.Parent.Activate
                                TargetRange.Cells(rowNum, colNum).Select
                                MsgBox("Field Value Update Error: " & ex.Message & " with Table: " & tableName & ", Field: " & fieldname & ", in sheet " & TargetRange.Parent.Name & ", & row " & rowNum & ", col: " & colNum, MsgBoxStyle.Critical, "DBMapper Error")
                                rst.CancelUpdate()
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
                        GoTo cleanup
                    End Try
                End If
                If (CUDFlags And rowCUDFlag = "d") Then
                    ExcelDnaUtil.Application.StatusBar = Left("Deleting " & primKeyDisplay & " in " & tableName, 255)
                    rst.Delete(AffectEnum.adAffectCurrent)
                End If
                rst.Close()
nextRow:
                Try
                    If IsNothing(TargetRange.Cells(rowNum + 1, 1).Value) OrElse TargetRange.Cells(rowNum + 1, 1).Value.ToString().Length = 0 Then finishLoop = True
                Catch ex As Exception
                    MsgBox("Error in first primary column: Cells(" & rowNum + 1 & ",1): " & ex.Message, MsgBoxStyle.Critical, "DBMapper Error")
                    'finishLoop = True '-> do not finish to allow erroneous data  !!
                End Try
            End If
            rowNum += 1
        Loop Until rowNum > TargetRange.Rows.Count Or (finishLoop And Not CUDFlags)
        ' clear CUD marks after completion
        resetCUDFlags()

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
        ' close connection to return it to the pool (automatically closes recordset object)...
        dbcnn.Close()
        ' DBSheet surrogate (CUDFlags), ask for refresh after DB Modification was done
        If CUDFlags And askBeforeExecute And calledByDBSeq = "" Then
            Dim retval As MsgBoxResult = MsgBox("Refresh DBListfetch/DBSetQuery for Data Range of DB Mapper?", MsgBoxStyle.Question + vbOKCancel, "Refresh DB Mapper")
            If retval = vbOK Then
                TargetRange.Cells(1, 1).Select()
                Globals.refreshData()
            End If
        End If
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
        If Left(paramTargetName, 8) <> "DBAction" Then Throw New Exception("target " & paramTargetName & " not matching passed DBModif type DBAction for " & targetRangeAddress & "/" & dbmapdefkey & " !")
        If paramTarget.Cells(1, 1).Text = "" Then Throw New Exception("No Action defined in " + paramTargetName + "(" + targetRangeAddress + ")")
        paramText = paramDefs
        TargetRange = paramTarget

        ' fill parameters:
        DBModifParams = functionSplit(paramText, ",", """", "def", "(", ")")
        If IsNothing(DBModifParams) Then Exit Sub
        ' check for completeness
        If DBModifParams.Length < 2 Then Throw New Exception("At least environment (can be empty) and database have to be provided as DBAction parameters !")
        env = DBModifParams(0)
        database = DBModifParams(1).Replace("""", "").Trim
        If database = "" Then
            Throw New Exception("No database given in DBAction paramText!")
        End If
        If DBModifParams.Length > 2 AndAlso DBModifParams(2) <> "" Then execOnSave = Convert.ToBoolean(DBModifParams(2))
        If DBModifParams.Length > 3 AndAlso DBModifParams(3) <> "" Then askBeforeExecute = Convert.ToBoolean(DBModifParams(3))
    End Sub

    Public Sub New(definitionXML As CustomXMLNode, paramTarget As Excel.Range)
        dbmapdefkey = definitionXML.BaseName
        ' if no target range is set, then no parameters can be found...
        If IsNothing(paramTarget) Then Exit Sub
        paramTargetName = getDBModifNameFromRange(paramTarget)
        DBmapSheet = paramTarget.Parent.Name
        targetRangeAddress = DBmapSheet + "!" + paramTarget.Address
        If Left(paramTargetName, 8) <> "DBAction" Then Throw New Exception("target " & paramTargetName & " not matching passed DBModif type DBAction for " & targetRangeAddress & "/" & dbmapdefkey & " !")
        If paramTarget.Cells(1, 1).Text = "" Then Throw New Exception("No Action defined in " + paramTargetName + "(" + targetRangeAddress + ")")
        TargetRange = paramTarget

        ' fill parameters from definition
        env = definitionXML.SelectSingleNode("ns0:env").Text
        database = definitionXML.SelectSingleNode("ns0:database").Text
        If database = "" Then Throw New Exception("No database given in DBAction definition!")
        execOnSave = Convert.ToBoolean(definitionXML.SelectSingleNode("ns0:execOnSave").Text)
        askBeforeExecute = Convert.ToBoolean(definitionXML.SelectSingleNode("ns0:askBeforeExecute").Text)
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
        If paramText = "" Then Throw New Exception("No Sequence defined in " + dbmapdefkey)
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

    Public Sub New(definitionXML As CustomXMLNode)
        dbmapdefkey = definitionXML.BaseName
        Try
            execOnSave = Convert.ToBoolean(definitionXML.SelectSingleNode("ns0:execOnSave").Text) ' should DBSequence be done on Excel Saving?
            askBeforeExecute = Convert.ToBoolean(definitionXML.SelectSingleNode("ns0:askBeforeExecute").Text) ' should DBSequence be done on Excel Saving?
        Catch ex As Exception
            Throw New Exception("problem with setting execOnSave or askBeforeExecute: " + ex.Message)
        End Try
        Dim seqSteps As Integer = definitionXML.SelectNodes("ns0:seqStep").Count
        If seqSteps = 0 Then
            Throw New Exception("no steps defined in DBSequence definition!")
        Else
            ReDim sequenceParams(seqSteps - 1)
            For i = 1 To seqSteps
                sequenceParams(i - 1) = definitionXML.SelectNodes("ns0:seqStep")(i).Text
            Next
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
    ''' <summary>avoid entering Application.SheetChange Event handler during listfetch/setquery</summary>
    Public preventChangeWhileFetching As Boolean = False

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
        openConnection = True
    End Function

    ''' <summary>in case there is a defined DBMapper underlying the DBListFetch/DBSetQuery target area then change the extent of that to the new area given in theRange</summary>
    ''' <param name="theRange"></param>
    Public Sub resizeDBMapperRange(theRange As Excel.Range)
        Dim dbMapperRangeName As String = getDBModifNameFromRange(theRange)
        If Left(dbMapperRangeName, 8) = "DBMapper" Then
            Dim NamesList As Excel.Names = ExcelDnaUtil.Application.ActiveWorkbook.Names
            Try : NamesList.Add(Name:=dbMapperRangeName, RefersTo:=theRange)
            Catch ex As Exception
                Throw New Exception("Error when assigning name '" & dbMapperRangeName & "' to ListObject Range: " & ex.Message)
            End Try
            Dim theDBMapper As DBMapper = Globals.DBModifDefColl("DBMapper").Item(dbMapperRangeName)
            ' notify DBMapper object of new target range
            theDBMapper.setTargetRange(theRange)
            ' in case of CUDFlags, reset them now...
            theDBMapper.resetCUDFlags()
        End If
    End Sub

    ''' <summary>creates a DBModif at the current active cell or edits an existing one defined in targetDefName (after being called in defined range or from ribbon + Ctrl + Shift)</summary>
    Sub createDBModif(createdDBModifType As String, Optional targetDefName As String = "")
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
        If Clipboard.ContainsText() And createdDBModifType = "DBMapper" Then
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
                Catch ex As Exception : MsgBox("Error when retrieving DB_DefName from clipboard: " & ex.Message, vbCritical, "DBMapper Legacy Creation Error") : Exit Sub : End Try
                Try : newDefString = "def(" + Replace(Mid(cpbdtext, commaBeforeConnDef, commaAfterConnDef - commaBeforeConnDef), "MSSQL", "") + Mid(cpbdtext, firstComma, commaBeforeConnDef - firstComma - 1) + Mid(cpbdtext, commaAfterConnDef - 1)
                Catch ex As Exception : MsgBox("Error when building new definition from clipboard: " & ex.Message, vbCritical, "DBMapper Legacy Creation Error") : Exit Sub : End Try
                ' assign new name to active cell
                ' Add doesn't work directly with ExcelDnaUtil.Application.ActiveWorkbook.Names (late binding), so create an object here...
                Dim NamesList As Excel.Names = ExcelDnaUtil.Application.ActiveWorkbook.Names
                Try : NamesList.Add(Name:=DB_DefName, RefersTo:=ExcelDnaUtil.Application.ActiveCell)
                Catch ex As Exception
                    MsgBox("Error when assigning name '" & DB_DefName & "' to active cell: " & ex.Message, vbCritical, "DBMapper Legacy Creation Error")
                    Exit Sub
                End Try

                ' build a new DBMapper with legacy constructor
                existingDBModif = New DBMapper(defkey:=DB_DefName, paramDefs:=newDefString, paramTarget:=ExcelDnaUtil.Application.ActiveCell)
                activeCellName = DB_DefName
                createdDBMapperFromClipboard = True
                Clipboard.Clear()
            End If
        End If

        ' for DBMappers defined in ListObjects, try potential name to ListObject (parts), only possible if manually defined !
        If createdDBModifType = "DBMapper" Then
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
        If DBModifDefColl.ContainsKey(createdDBModifType) AndAlso DBModifDefColl(createdDBModifType).ContainsKey(targetDefName) Then existingDBModif = DBModifDefColl(createdDBModifType).Item(targetDefName)

        ' prepare DBModifier Create Dialog
        Dim theDBModifCreateDlg As DBModifCreate = New DBModifCreate()
        With theDBModifCreateDlg
            ' store DBModification type in tag for validation purposes...
            .Tag = createdDBModifType
            .envSel.DataSource = Globals.environdefs
            .envSel.SelectedIndex = -1
            .DBModifName.Text = Replace(activeCellName, createdDBModifType, "")
            .RepairDBSeqnce.Hide()
            .NameLabel.Text = IIf(createdDBModifType = "DBSeqnce", "DBSequence", createdDBModifType) & " Name:"
            .Text = "Edit " & IIf(createdDBModifType = "DBSeqnce", "DBSequence", createdDBModifType) & " definition"
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
            End If
            If createdDBModifType = "DBSeqnce" Then
                ' hide controls irrelevant for DBSeqnce
                .TargetRangeAddress.Hide()
                .envSel.Hide()
                .EnvironmentLabel.Hide()
                .Database.Hide()
                .DatabaseLabel.Hide()
                .DBSeqenceDataGrid.Top = 55
                .DBSeqenceDataGrid.Height = 320
                .execOnSave.Top = .up.Top
                .AskForExecute.Top = .up.Top
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
                                MsgBox(theFunc & " (in " & searchCell.Parent.Name & "!" & searchCell.Address & ") has multiple " & IIf(searchCell.Rows.Count > 1, "rows !", "columns !") & ", which leads to problems in DBSequences...", vbCritical, "DB Sequence Creation Error")
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
                    Try : ExcelDnaUtil.Application.ActiveWorkbook.Names(activeCellName).Delete : Catch ex As Exception : End Try
                End If
                Exit Sub
            End If

            ' only for DBMapper or DBAction: change target range name
            If createdDBModifType <> "DBSeqnce" Then
                Dim targetRange As Excel.Range
                If IsNothing(existingDBModif) Then
                    targetRange = ExcelDnaUtil.Application.ActiveCell
                Else
                    targetRange = existingDBModif.getTargetRange()
                End If

                ' set content range name: first clear name
                Try : ExcelDnaUtil.Application.ActiveWorkbook.Names(activeCellName).Delete : Catch ex As Exception : End Try
                ' then (re)set name
                Dim NamesList As Excel.Names = ExcelDnaUtil.Application.ActiveWorkbook.Names
                Try
                    NamesList.Add(Name:=createdDBModifType + .DBModifName.Text, RefersTo:=targetRange)
                Catch ex As Exception : MsgBox("Error when assigning name '" & createdDBModifType & .DBModifName.Text & "' to active cell: " & ex.Message, vbCritical, "DB Sequence Creation Error")
                End Try
            End If

            Dim CustomXmlParts As Object = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
            If CustomXmlParts.Count = 0 Then ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.Add("<root xmlns=""DBModifDef""></root>")
            CustomXmlParts = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
            ' remove old node in case of renaming DBModifier...
            CustomXmlParts(1).SelectSingleNode("/ns0:root/ns0:" + activeCellName).Delete
            ' NamespaceURI:="DBModifDef" is required to avoid adding a xmlns attribute to each element.
            CustomXmlParts(1).SelectSingleNode("/ns0:root").AppendChildNode(createdDBModifType + .DBModifName.Text, NamespaceURI:="DBModifDef")
            Dim dbModifNode As CustomXMLNode = CustomXmlParts(1).SelectSingleNode("/ns0:root/ns0:" + createdDBModifType + .DBModifName.Text)
            If createdDBModifType = "DBMapper" Then
                dbModifNode.AppendChildNode("env", NamespaceURI:="DBModifDef", NodeValue:=IIf(.envSel.SelectedIndex = -1, "", (.envSel.SelectedIndex + 1).ToString()))
                dbModifNode.AppendChildNode("database", NamespaceURI:="DBModifDef", NodeValue:= .Database.Text)
                dbModifNode.AppendChildNode("tableName", NamespaceURI:="DBModifDef", NodeValue:= .Tablename.Text)
                dbModifNode.AppendChildNode("primKeysStr", NamespaceURI:="DBModifDef", NodeValue:= .PrimaryKeys.Text)
                dbModifNode.AppendChildNode("insertIfMissing", NamespaceURI:="DBModifDef", NodeValue:= .insertIfMissing.Checked.ToString())
                dbModifNode.AppendChildNode("executeAdditionalProc", NamespaceURI:="DBModifDef", NodeValue:= .addStoredProc.Text)
                dbModifNode.AppendChildNode("ignoreColumns", NamespaceURI:="DBModifDef", NodeValue:= .IgnoreColumns.Text)
                dbModifNode.AppendChildNode("CUDFlags", NamespaceURI:="DBModifDef", NodeValue:= .CUDflags.Checked.ToString())
                dbModifNode.AppendChildNode("execOnSave", NamespaceURI:="DBModifDef", NodeValue:= .execOnSave.Checked.ToString())
                dbModifNode.AppendChildNode("askBeforeExecute", NamespaceURI:="DBModifDef", NodeValue:= .AskForExecute.Checked.ToString())
            ElseIf createdDBModifType = "DBAction" Then
                dbModifNode.AppendChildNode("env", NamespaceURI:="DBModifDef", NodeValue:=IIf(.envSel.SelectedIndex = -1, "", (.envSel.SelectedIndex + 1).ToString()))
                dbModifNode.AppendChildNode("database", NamespaceURI:="DBModifDef", NodeValue:= .Database.Text)
                dbModifNode.AppendChildNode("execOnSave", NamespaceURI:="DBModifDef", NodeValue:= .execOnSave.Checked.ToString())
                dbModifNode.AppendChildNode("askBeforeExecute", NamespaceURI:="DBModifDef", NodeValue:= .AskForExecute.Checked.ToString())
            ElseIf createdDBModifType = "DBSeqnce" Then
                If .Tag = "repaired" Then
                    'TODO: modify to new xml defs.
                    'newParamText = .RepairDBSeqnce.Text
                Else
                    dbModifNode.AppendChildNode("execOnSave", NamespaceURI:="DBModifDef", NodeValue:= .execOnSave.Checked.ToString()) ' should DBSequence be done on Excel Saving?
                    dbModifNode.AppendChildNode("askBeforeExecute", NamespaceURI:="DBModifDef", NodeValue:= .AskForExecute.Checked.ToString()) ' should DBSequence be done on Excel Saving?
                    For i As Integer = 0 To .DBSeqenceDataGrid.Rows().Count - 2
                        dbModifNode.AppendChildNode("seqStep", NamespaceURI:="DBModifDef", NodeValue:= .DBSeqenceDataGrid.Rows(i).Cells(0).Value)
                    Next
                End If
            End If
            ' refresh mapper definitions to reflect changes immediately...
            getDBModifDefinitions()
            ' extend Datarange for new DBMappers immediately after definition...
            If createdDBModifType = "DBMapper" Then
                DirectCast(Globals.DBModifDefColl("DBMapper").Item(createdDBModifType + .DBModifName.Text), DBMapper).extendDataRange(ignoreCUDFlag:=True)
            End If
        End With
    End Sub

    ''' <summary>gets defined names for DBModifier (DBMapper/DBAction/DBSeqnce) invocation in the current workbook and updates Ribbon with it</summary>
    Public Sub getDBModifDefinitions()
        ' load DBModifier definitions (objects) into Global collection DBModifDefColl
        Try
            Globals.DBModifDefColl = New Dictionary(Of String, Dictionary(Of String, DBModif))
            Dim CustomXmlParts As Object = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
            If CustomXmlParts.Count = 0 Then
                ' get DBModifier definitions from docproperties
                For Each docproperty As DocumentProperty In ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties
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
                                    ' might fail...
                                    Try : targetRange = rangename.RefersToRange : Catch ex As Exception : End Try
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
                        Dim DBModifParams() As String
                        ' fill parameters into CustomXMLPart:
                        ' create CustomXMLPart to migrate docproperty definitions
                        If CustomXmlParts.Count = 0 Then ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.Add("<root xmlns=""DBModifDef""></root>")
                        CustomXmlParts = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
                        ' NamespaceURI:="DBModifDef" is required to avoid adding a xmlns attribute to each element.
                        CustomXmlParts(1).SelectSingleNode("/ns0:root").AppendChildNode(docproperty.Name, NamespaceURI:="DBModifDef")
                        Dim dbModifNode As CustomXMLNode = CustomXmlParts(1).SelectSingleNode("/ns0:root/ns0:" + docproperty.Name)
                        If DBModiftype = "DBMapper" Then
                            ' legacy constructor for DBModifParams
                            newDBModif = New DBMapper(docproperty.Name, docproperty.Value, targetRange)
                            DBModifParams = newDBModif.GetDBModifParams()
                            dbModifNode.AppendChildNode("env", NamespaceURI:="DBModifDef", NodeValue:=DBModifParams(0))
                            dbModifNode.AppendChildNode("database", NamespaceURI:="DBModifDef", NodeValue:=DBModifParams(1).Replace("""", "").Trim)
                            dbModifNode.AppendChildNode("tableName", NamespaceURI:="DBModifDef", NodeValue:=DBModifParams(2).Replace("""", "").Trim)
                            dbModifNode.AppendChildNode("primKeysStr", NamespaceURI:="DBModifDef", NodeValue:=DBModifParams(3).Replace("""", "").Trim)
                            dbModifNode.AppendChildNode("insertIfMissing", NamespaceURI:="DBModifDef", NodeValue:=If(DBModifParams.Length > 4, DBModifParams(4), "False"))
                            dbModifNode.AppendChildNode("executeAdditionalProc", NamespaceURI:="DBModifDef", NodeValue:=If(DBModifParams.Length > 5, DBModifParams(5).Replace("""", "").Trim, ""))
                            dbModifNode.AppendChildNode("ignoreColumns", NamespaceURI:="DBModifDef", NodeValue:=If(DBModifParams.Length > 6, DBModifParams(6).Replace("""", "").Trim, ""))
                            dbModifNode.AppendChildNode("execOnSave", NamespaceURI:="DBModifDef", NodeValue:=If(DBModifParams.Length > 7, DBModifParams(7), "False"))
                            dbModifNode.AppendChildNode("CUDFlags", NamespaceURI:="DBModifDef", NodeValue:=If(DBModifParams.Length > 8, DBModifParams(8), "False"))
                            dbModifNode.AppendChildNode("askBeforeExecute", NamespaceURI:="DBModifDef", NodeValue:=If(DBModifParams.Length > 9, DBModifParams(9), "True"))
                        ElseIf DBModiftype = "DBAction" Then
                            ' legacy constructor for DBModifParams
                            newDBModif = New DBAction(docproperty.Name, docproperty.Value, targetRange)
                            DBModifParams = newDBModif.GetDBModifParams()
                            dbModifNode.AppendChildNode("env", NamespaceURI:="DBModifDef", NodeValue:=DBModifParams(0))
                            dbModifNode.AppendChildNode("database", NamespaceURI:="DBModifDef", NodeValue:=DBModifParams(1).Replace("""", "").Trim)
                            dbModifNode.AppendChildNode("execOnSave", NamespaceURI:="DBModifDef", NodeValue:=If(DBModifParams.Length > 2, DBModifParams(2), "False"))
                            dbModifNode.AppendChildNode("askBeforeExecute", NamespaceURI:="DBModifDef", NodeValue:=If(DBModifParams.Length > 3, DBModifParams(3), "True"))
                        ElseIf DBModiftype = "DBSeqnce" Then
                            ' legacy constructor for DBModifParams
                            newDBModif = New DBSeqnce(docproperty.Name, docproperty.Value)
                            DBModifParams = newDBModif.GetDBModifParams()
                            dbModifNode.AppendChildNode("execOnSave", NamespaceURI:="DBModifDef", NodeValue:=DBModifParams(0)) ' should DBSequence be done on Excel Saving?
                            If DBModifParams(1) = Boolean.FalseString Or DBModifParams(1) = Boolean.TrueString Then
                                dbModifNode.AppendChildNode("askBeforeExecute", NamespaceURI:="DBModifDef", NodeValue:=DBModifParams(1)) ' should DBSequence be done on Excel Saving?
                                For i As Integer = 2 To UBound(DBModifParams)
                                    dbModifNode.AppendChildNode("seqStep", NamespaceURI:="DBModifDef", NodeValue:=DBModifParams(i))
                                Next
                            Else ' legacy: no askBeforeExecute in old versions ...
                                dbModifNode.AppendChildNode("askBeforeExecute", NamespaceURI:="DBModifDef", NodeValue:="True")
                                For i As Integer = 1 To UBound(DBModifParams)
                                    dbModifNode.AppendChildNode("seqStep", NamespaceURI:="DBModifDef", NodeValue:=DBModifParams(i))
                                Next
                            End If
                        Else
                            MsgBox("Error, not supported DBModiftype: " & DBModiftype, vbCritical, "DBModifier Definitions Error")
                            newDBModif = Nothing
                        End If
                        ' remove the old docproperty
                        Dim docpropertyName As String = docproperty.Name
                        Try : ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties(docpropertyName).Delete : Catch ex As Exception : End Try
                        ' ... and add it to the collection DBModifDefColl
                        Dim defColl As Dictionary(Of String, DBModif) ' definition lookup collection for DBModifiername -> object
                        If Not DBModifDefColl.ContainsKey(DBModiftype) Then
                            ' add to new DBModiftype "menu"
                            defColl = New Dictionary(Of String, DBModif)
                            defColl.Add(docpropertyName, newDBModif)
                            DBModifDefColl.Add(DBModiftype, defColl)
                        Else
                            ' add definition to existing DBModiftype "menu"
                            defColl = DBModifDefColl(DBModiftype)
                            defColl.Add(docpropertyName, newDBModif)
                        End If
                    End If
                Next
            Else
                ' read definitions from CustomXMLParts
                ' get DBModifier definitions from docproperties
                For Each customXMLNodeDef As CustomXMLNode In CustomXmlParts(1).SelectSingleNode("/ns0:root").ChildNodes
                    Dim DBModiftype As String = Left(customXMLNodeDef.BaseName, 8)
                    If DBModiftype = "DBSeqnce" Or DBModiftype = "DBMapper" Or DBModiftype = "DBAction" Then
                        Dim nodeName As String = customXMLNodeDef.BaseName
                        Dim targetRange As Excel.Range = Nothing
                        ' for DBMappers and DBActions the data of the DBModification is stored in Ranges, so check for those and get the Range
                        If DBModiftype = "DBMapper" Or DBModiftype = "DBAction" Then
                            For Each rangename As Excel.Name In ExcelDnaUtil.Application.ActiveWorkbook.Names
                                Dim rangenameName As String = Replace(rangename.Name, rangename.Parent.Name & "!", "")
                                If rangenameName = nodeName And InStr(rangename.RefersTo, "#REF!") > 0 Then
                                    MsgBox(DBModiftype + " definitions range [" + rangename.Parent.Name + "]" + rangename.Name + " contains #REF!", vbCritical, "DBModifier Definitions Error")
                                    Exit For
                                ElseIf rangenameName = nodeName Then
                                    ' might fail...
                                    Try : targetRange = rangename.RefersToRange : Catch ex As Exception : End Try
                                    Exit For
                                End If
                            Next
                            If IsNothing(targetRange) Then
                                MsgBox("Error, required target range named '" & nodeName & "' not existing for " & DBModiftype & "." & vbCrLf & "either create target range or delete CustomXML definition node named '" & nodeName & "' (Ctrl-Shift-Click on Store DBModif Data dialogBox launcher)!", vbCritical, "DBModifier Definitions Error")
                                Continue For
                            End If
                        End If
                        ' finally create the DBModif Object ...
                        Dim newDBModif As DBModif
                        ' fill parameters into CustomXMLPart:
                        If DBModiftype = "DBMapper" Then
                            newDBModif = New DBMapper(customXMLNodeDef, targetRange)
                        ElseIf DBModiftype = "DBAction" Then
                            newDBModif = New DBAction(customXMLNodeDef, targetRange)
                        ElseIf DBModiftype = "DBSeqnce" Then
                            newDBModif = New DBSeqnce(customXMLNodeDef)
                        Else
                            MsgBox("Error, not supported DBModiftype: " & DBModiftype, vbCritical, "DBModifier Definitions Error")
                            newDBModif = Nothing
                        End If
                        ' ... and add it to the collection DBModifDefColl
                        Dim defColl As Dictionary(Of String, DBModif) ' definition lookup collection for DBModifiername -> object
                        If Not DBModifDefColl.ContainsKey(DBModiftype) Then
                            ' add to new DBModiftype "menu"
                            defColl = New Dictionary(Of String, DBModif)
                            defColl.Add(customXMLNodeDef.BaseName, newDBModif)
                            DBModifDefColl.Add(DBModiftype, defColl)
                        Else
                            ' add definition to existing DBModiftype "menu"
                            defColl = DBModifDefColl(DBModiftype)
                            defColl.Add(customXMLNodeDef.BaseName, newDBModif)
                        End If
                    End If
                Next
            End If
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

