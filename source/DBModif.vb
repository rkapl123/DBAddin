Imports ADODB
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports Microsoft.Office.Core
Imports System.Data
Imports System.Data.SqlClient


''' <summary>Abstraction of a DB Modification Object (concrete classes DBMapper, DBAction or DBSeqnce)</summary>
Public MustInherit Class DBModif

    '''<summary>unique key of DBModif</summary>
    Protected dbmodifName As String
    ''' <summary>Range where DBMapper data is located (only DBMapper and DBAction; paramText is stored in custom doc properties having the same Name)</summary>
    Protected TargetRange As Excel.Range
    '''<summary>should DBMap be saved / DBAction be done on Excel Saving? (default no)</summary>
    Protected execOnSave As Boolean = False
    ''' <summary>ask for confirmation before executtion of DBModif</summary>
    Protected askBeforeExecute As Boolean = True
    ''' <summary>environment specific for the DBModif object, if left empty then set to default environment (either 0 or currently selected environment)</summary>
    Protected env As String = ""
    ''' <summary>Text displayed for confirmation before doing dbModif instead of standard text</summary>
    Protected confirmText As String = ""

    Public Sub New(definitionXML As CustomXMLNode)
        If definitionXML.Attributes.Count > 0 Then
            dbmodifName = definitionXML.BaseName + definitionXML.Attributes(1).Text
        Else
            dbmodifName = definitionXML.BaseName
        End If
        execOnSave = Convert.ToBoolean(getParamFromXML(definitionXML, "execOnSave"))
        askBeforeExecute = Convert.ToBoolean(getParamFromXML(definitionXML, "askBeforeExecute"))
        confirmText = getParamFromXML(definitionXML, "confirmText")
    End Sub

    ''' <summary>needed for legacy DBmapper constructor</summary>
    Public Sub New()

    End Sub

    ''' <summary>public accessor function</summary>
    ''' <returns>the targetRangeAddress nicely formatted</returns>
    Public Function getTargetRangeAddress() As String
        getTargetRangeAddress = ""
        If TypeName(Me) <> "DBSeqnce" Then
            getTargetRangeAddress = TargetRange.Parent.Name + "!" + TargetRange.Address
        End If
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
        If TypeName(Me) = "DBSeqnce" Then Throw New NotImplementedException()
        If env <> "" Then getEnv = Convert.ToInt16(env)
    End Function

    ''' <summary>is saving needed for this DBModifier</summary>
    ''' <returns>true for saving necessary</returns>
    Public Function DBModifSaveNeeded() As Boolean
        Return execOnSave
    End Function

    ''' <summary>gets the name for this DBModifier</summary>
    ''' <returns></returns>
    Public Function getName() As String
        Return dbmodifName
    End Function

    ''' <summary>wrapper to get the single definition element values from the DBModifier CustomXML node, also checks for multiple definition elements</summary>
    ''' <param name="definitionXML">the CustomXML node for the DBModifier</param>
    ''' <param name="nodeName">the definition element's name (eg "env")</param>
    ''' <returns>the definition element's value</returns>
    ''' <exception cref="Exception">if multiple elements exist for the definition element's name throw warning !</exception>
    Protected Function getParamFromXML(definitionXML As CustomXMLNode, nodeName As String) As String
        Dim nodeCount As Integer = definitionXML.SelectNodes("ns0:" + nodeName).Count
        If nodeCount = 0 Then
            getParamFromXML = "" ' optional nodes become empty strings
        Else
            getParamFromXML = definitionXML.SelectSingleNode("ns0:" + nodeName).Text
        End If
    End Function

    ''' <summary>to re-add hidden features only available in XML</summary>
    ''' <param name="definitionXML">the definition node of the DB Modifier where the hidden features should be added</param>
    Public Overridable Sub addHiddenFeatureDefs(definitionXML As CustomXMLNode)
        definitionXML.AppendChildNode("confirmText", NamespaceURI:="DBModifDef", NodeValue:=confirmText)
    End Sub

    ''' <summary>does the actual DB Modification</summary>
    ''' <param name="WbIsSaving"></param>
    ''' <param name="calledByDBSeq"></param>
    Public Overridable Sub doDBModif(Optional WbIsSaving As Boolean = False, Optional calledByDBSeq As String = "", Optional TransactionOpen As Boolean = False)
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
        If TypeName(Me) = "DBSeqnce" Then Throw New NotImplementedException() ' DB Sequences have no Target Range
        TargetRange = newTargetRange
    End Sub

    ''' <summary>simply open a database connection, required for DBBegin Transaction (from next step)</summary>
    ''' <returns></returns>
    Public Overridable Function openDatabase() As Boolean
        Throw New NotImplementedException()
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

    ''' <summary>refresh a DB Function (currently only DBListFetch and DBSetQuery) by invoking its respective DB*Action procedure
    ''' It is necessary to prepare the inputs to this DB*Action procedure as the UDF cannot be invoked from here</summary>
    ''' <param name="srcExtent">the unique hidden name of the DB Function cell (DBFsource(GUID))</param>
    ''' <param name="executedDBMappers">in a DB Sequence, this parameter notifies of DBMappers that were executed before to allow avoidance of refreshing changes</param>
    ''' <param name="modifiedDBMappers">in a DB Sequence, this parameter notifies of a DBMapper that had changes, necessary to avoid deadlocks</param>
    ''' <param name="TransactionIsOpen">in a DB Sequence, this parameter notifies of an open transaction, necessary to avoid deadlocks</param>
    ''' <returns></returns>
    Protected Function doDBRefresh(srcExtent As String, Optional executedDBMappers As Dictionary(Of String, Boolean) = Nothing, Optional modifiedDBMappers As Dictionary(Of String, Boolean) = Nothing, Optional TransactionIsOpen As Boolean = False) As Boolean
        If IsNothing(executedDBMappers) Then executedDBMappers = New Dictionary(Of String, Boolean)
        If IsNothing(modifiedDBMappers) Then modifiedDBMappers = New Dictionary(Of String, Boolean)
        doDBRefresh = False
        ' refresh DBFunction in sequence, invoke this "manually", simulating the call of the user defined function by excel
        Dim caller As Excel.Range
        Try : caller = ExcelDnaUtil.Application.Range(srcExtent)
        Catch ex As Exception
            ErrorMsg("Didn't find caller cell of DBRefresh using srcExtent " + srcExtent + "!", "Refresh of DB Functions")
            Exit Function
        End Try
        If caller.Parent.ProtectContents Then
            ErrorMsg("Worksheet " + caller.Parent.Name + " is content protected, can't refresh DB Function !")
            Exit Function
        End If
        Dim targetExtent = Replace(srcExtent, "DBFsource", "DBFtarget")
        Dim target As Excel.Range
        Try : target = ExcelDnaUtil.Application.Range(targetExtent)
        Catch ex As Exception
            ErrorMsg("Didn't find target of DBRefresh using targetExtent " + targetExtent + "!", "Refresh of DB Functions")
            Exit Function
        End Try
        If target.Parent.ProtectContents Then
            ErrorMsg("Worksheet " + target.Parent.Name + " is content protected, can't refresh DB Function !")
            Exit Function
        End If
        Dim DBMapperUnderlying As String = getDBModifNameFromRange(target)
        Dim targetExtentF = Replace(srcExtent, "DBFsource", "DBFtargetF")
        Dim formulaRange As Excel.Range = Nothing
        ' formulaRange might not exist
        Try : formulaRange = ExcelDnaUtil.Application.Range(targetExtentF) : Catch ex As Exception : End Try
        If Not IsNothing(formulaRange) AndAlso formulaRange.Parent.ProtectContents Then
            ErrorMsg("Worksheet " + formulaRange.Parent.Name + " is content protected, can't refresh DB Function !")
            Exit Function
        End If
        Dim DBMapperUnderlyingF As String = getDBModifNameFromRange(formulaRange)
        ' allow for avoidance of overwriting users changes with CUDFlags after an data error occurred
        If hadError Then
            If executedDBMappers.ContainsKey(DBMapperUnderlying) Then
                Dim retval = QuestionMsg(theMessage:="Error(s) occured during sequence, really refresh Targetrange? This could lead to loss of entries.", questionTitle:="Refresh of DB Functions in DB Sequence")
                If retval = vbCancel Then Exit Function
            End If
        End If
        ' prevent deadlock if we are in a transaction and want to refresh the area that was changed.
        If (modifiedDBMappers.ContainsKey(DBMapperUnderlying) Or modifiedDBMappers.ContainsKey(DBMapperUnderlyingF)) And TransactionIsOpen Then
            ErrorMsg("Transaction affecting the target area is still open, refreshing it would result in a deadlock on the database table. Skipping refresh, consider placing refresh outside of transaction.", "Refresh of DB Functions in DB Sequence")
            Exit Function
        End If
        ' reset query cache, so we really get new data !
        Dim callID As String
        Try
            ' get the callID of the underlying name of the target (key of the queryCache and StatusCollection)
            callID = "[" + caller.Parent.Parent.Name + "]" + caller.Parent.Name + "!" + caller.Address
        Catch ex As Exception
            ErrorMsg("Didn't find target of DBRefresh !", "Refresh of DB Functions")
            Exit Function
        End Try
        Try
            ' StatusCollection doesn't necessarily have the callID contained
            If Not StatusCollection.ContainsKey(callID) Then StatusCollection.Add(callID, New ContainedStatusMsg)
            Dim functionFormula As String = ExcelDnaUtil.Application.Range(srcExtent).Formula
            If UCase(Left(functionFormula, 12)) = "=DBLISTFETCH" Then
                LogInfo("Refresh DBListFetch: " + callID)
                Dim functionArgs = functionSplit(functionFormula, ",", """", "DBListFetch", "(", ")")
                If functionArgs.Length() < 3 Then
                    functionArgs = functionSplit(functionFormula, listSepLocal, """", "DBListFetch", "(", ")")
                End If
                Dim targetRangeName As String : targetRangeName = functionArgs(2)
                ' check if fetched argument targetRangeName is really a name or just a plain range address
                If Not existsNameInWb(targetRangeName, caller.Parent.Parent) And Not existsNameInSheet(targetRangeName, caller.Parent) Then targetRangeName = ""
                Dim formulaRangeName As String
                If UBound(functionArgs) > 2 Then
                    formulaRangeName = functionArgs(3)
                    If Not existsNameInWb(formulaRangeName, caller.Parent.Parent) And Not existsNameInSheet(formulaRangeName, caller.Parent) Then formulaRangeName = ""
                Else
                    formulaRangeName = ""
                End If
                Dim extendDataArea As Integer = 0
                If UBound(functionArgs) > 3 AndAlso functionArgs(4) <> "" Then
                    extendDataArea = Convert.ToInt16(functionArgs(4))
                End If
                Dim HeaderInfo As Boolean = False
                If UBound(functionArgs) > 4 AndAlso functionArgs(5) <> "" Then
                    HeaderInfo = convertToBool(functionArgs(5))
                End If
                Dim AutoFit As Boolean = False
                If UBound(functionArgs) > 5 AndAlso functionArgs(6) <> "" Then
                    AutoFit = convertToBool(functionArgs(6))
                End If
                Dim autoformat As Boolean = False
                If UBound(functionArgs) > 6 AndAlso functionArgs(7) <> "" Then
                    autoformat = convertToBool(functionArgs(7))
                End If
                Dim ShowRowNums As Boolean = False
                If UBound(functionArgs) > 7 AndAlso functionArgs(8) <> "" Then
                    ShowRowNums = convertToBool(functionArgs(8))
                End If
                ' call action procedure directly as we can avoid the external context required in the UDF
                DBListFetchAction(callID, getQuery(functionArgs(0), caller), caller, target, getConnString(functionArgs(1), caller, False), formulaRange, extendDataArea, HeaderInfo, AutoFit, autoformat, ShowRowNums, targetRangeName, formulaRangeName)
            ElseIf UCase(Left(functionFormula, 11)) = "=DBSETQUERY" Then
                LogInfo("Refresh DBSetQuery: " + callID)
                Dim functionArgs = functionSplit(functionFormula, ",", """", "DBSetQuery", "(", ")")
                If functionArgs.Length() < 3 Then
                    functionArgs = functionSplit(functionFormula, listSepLocal, """", "DBSetQuery", "(", ")")
                End If
                Dim targetRangeName As String : targetRangeName = functionArgs(2)
                If UBound(functionArgs) = 3 Then targetRangeName += "," + functionArgs(3)
                Functions.DBSetQueryAction(callID, getQuery(functionArgs(0), caller), target, getConnString(functionArgs(1), caller, True), caller, targetRangeName)
            ElseIf UCase(Left(functionFormula, 11)) = "=DBROWFETCH" Then
                LogInfo("Refresh DBRowFetch: " + callID)
                Dim functionArgs = functionSplit(functionFormula, ",", """", "DBRowFetch", "(", ")")
                If functionArgs.Length() < 3 Then
                    functionArgs = functionSplit(functionFormula, listSepLocal, """", "DBRowFetch", "(", ")")
                End If
                Dim HeaderInfo As Boolean = False
                Dim tempArray() As Excel.Range = Nothing
                If Boolean.TryParse(ExcelDnaUtil.Application.Evaluate(functionArgs(2)), HeaderInfo) Then
                    For i = 3 To UBound(functionArgs)
                        ReDim Preserve tempArray(i - 3)
                        tempArray(i - 3) = target.Parent.Range(functionArgs(i))
                    Next
                Else
                    For i = 2 To UBound(functionArgs)
                        ReDim Preserve tempArray(i - 2)
                        tempArray(i - 2) = target.Parent.Range(functionArgs(i))
                    Next
                End If
                Functions.DBRowFetchAction(callID, getQuery(functionArgs(0), caller), caller, tempArray, getConnString(functionArgs(1), caller, False), HeaderInfo)
            End If
        Catch ex As Exception
            ErrorMsg("Exception: " + ex.Message, "DBRefresh")
        End Try
        doDBRefresh = True
    End Function

    ''' <summary>get DBFunction's query from passed function argument</summary>
    ''' <param name="funcArg">function argument parsed from DBFunction formula</param>
    ''' <param name="caller">function caller range</param>
    ''' <returns></returns>
    Private Function getQuery(funcArg As String, caller As Excel.Range) As String
        Dim Query As Object
        Dim rangePart() As String = Split(funcArg, "!")
        If UBound(rangePart) = 1 Then
            Query = ExcelDnaUtil.Application.Evaluate(funcArg)
        Else
            ' avoid appending worksheet name if argument is a string (only references get appended)
            If Left(funcArg, 1) = """" Then
                Query = ExcelDnaUtil.Application.Evaluate(funcArg)
            Else
                Query = ExcelDnaUtil.Application.Evaluate(caller.Parent.Name + "!" + funcArg)
            End If
        End If
            If TypeName(Query) = "Range" Then Query = Query.Value.ToString
        getQuery = Query
    End Function

    ''' <summary>get connection string from passed function argument</summary>
    ''' <param name="funcArg">function argument parsed from DBFunction formula, can be empty, a number or a String</param>
    ''' <param name="caller">function caller range</param>
    ''' <returns>resolved connection string</returns>
    Private Function getConnString(funcArg As String, caller As Excel.Range, getConnStrForDBSet As Boolean) As String
        Dim ConnString As Object = Replace(funcArg, """", "")
        Dim testInt As Integer : Dim EnvPrefix As String = ""
        If CStr(ConnString) <> "" And Not Integer.TryParse(ConnString, testInt) Then
            Dim rangePart() As String = Split(funcArg, "!")
            If UBound(rangePart) = 1 Then
                ConnString = ExcelDnaUtil.Application.Evaluate(funcArg)
            Else
                ' avoid appending worksheet name if argument is a string (only references get appended)
                If Left(funcArg, 1) = """" Then
                    ConnString = ExcelDnaUtil.Application.Evaluate(funcArg)
                Else
                    ConnString = ExcelDnaUtil.Application.Evaluate(caller.Parent.Name + "!" + funcArg)
                End If
            End If
        End If
        If Integer.TryParse(ConnString, testInt) Then
            ConnString = Convert.ToDouble(testInt)
        End If
        Functions.resolveConnstring(ConnString, EnvPrefix, getConnStrForDBSet)
        getConnString = CStr(ConnString)
    End Function
End Class

''' <summary>DBMappers are used to store a range of data in Excel to the database. A special type of DBMapper is the CUD DBMapper for realizing the former DBSheets</summary>
Public Class DBMapper : Inherits DBModif

    ''' <summary>DBModif name of target range</summary>
    Private paramTargetName As String
    ''' <summary>Database to store to</summary>
    Private database As String
    ''' <summary>Database Table, where Data is to be stored</summary>
    Private tableName As String = ""
    ''' <summary>count of primary keys in datatable, starting from the leftmost column</summary>
    Private primKeysCount As Integer = 0
    ''' <summary>if set, then insert row into table if primary key is missing there. Default = False (only update)</summary>
    Private insertIfMissing As Boolean = False
    ''' <summary>additional stored procedure to be executed after saving</summary>
    Private executeAdditionalProc As String = ""
    ''' <summary>columns to be ignored (helper columns)</summary>
    Private ignoreColumns As String = ""
    ''' <summary>respect C/U/D Flags (DBSheet functionality)</summary>
    Public CUDFlags As Boolean = False
    ''' <summary>if set, don't notify error values in cells during update/insert</summary>
    Private IgnoreDataErrors As Boolean = False
    ''' <summary>used to pass whether changes were done</summary>
    Private changesDone As Boolean = False

    ''' <summary>legacy constructor for mapping existing DBMapper macro calls (copy in clipboard)</summary>
    ''' <param name="defkey"></param>
    ''' <param name="paramDefs"></param>
    ''' <param name="paramTarget"></param>
    Public Sub New(defkey As String, paramDefs As String, paramTarget As Excel.Range)
        dbmodifName = defkey
        ' if no target range is set, then no parameters can be found...
        If IsNothing(paramTarget) Then
            Throw New Exception("paramTarget is Nothing")
        End If
        paramTargetName = getDBModifNameFromRange(paramTarget)
        If Left(paramTargetName, 8) <> "DBMapper" Then
            Throw New Exception("target " + paramTargetName + " not matching passed DBModif type DBMapper for " + getTargetRangeAddress() + "/" + dbmodifName + "!")
        End If
        Dim paramText As String = paramDefs
        TargetRange = paramTarget

        Dim DBModifParams() As String = functionSplit(paramText, ",", """", "def", "(", ")")
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
        Try
            primKeysCount = DBModifParams(3).Split(",").Length
        Catch ex As Exception
            Throw New Exception("couldn't get primary key count given in DBMapper paramText (should be a comma separated list)!")
        End Try
        If DBModifParams.Length > 4 AndAlso DBModifParams(4) <> "" Then insertIfMissing = Convert.ToBoolean(DBModifParams(4))
        If DBModifParams.Length > 5 AndAlso DBModifParams(5) <> "" Then executeAdditionalProc = DBModifParams(5).Replace("""", "").Trim
        If DBModifParams.Length > 6 AndAlso DBModifParams(6) <> "" Then ignoreColumns = DBModifParams(6).Replace("""", "").Trim
        If DBModifParams.Length > 7 AndAlso DBModifParams(7) <> "" Then execOnSave = Convert.ToBoolean(DBModifParams(7))
    End Sub

    ''' <summary>normal constructor with definition XML</summary>
    ''' <param name="definitionXML"></param>
    ''' <param name="paramTarget"></param>
    Public Sub New(definitionXML As CustomXMLNode, paramTarget As Excel.Range)
        MyBase.New(definitionXML)
        Try
            ' if no target range is set, then no parameters can be found...
            If IsNothing(paramTarget) Then Throw New Exception("paramTarget is Nothing")
            paramTargetName = getDBModifNameFromRange(paramTarget)
            If Left(paramTargetName, 8) <> "DBMapper" Then Throw New Exception("target " + paramTargetName + " not matching passed DBModif type DBMapper for " + getTargetRangeAddress() + "/" + dbmodifName + "!")
            TargetRange = paramTarget
            ' fill parameters from definition
            env = getParamFromXML(definitionXML, "env")
            database = getParamFromXML(definitionXML, "database")
            If database = "" Then Throw New Exception("No database given in DBMapper definition!")
            tableName = getParamFromXML(definitionXML, "tableName")
            If tableName = "" Then Throw New Exception("No Tablename given in DBMapper definition!")
            Try
                primKeysCount = Convert.ToInt32(getParamFromXML(definitionXML, "primKeysStr"))
            Catch ex As Exception
                Throw New Exception("couldn't get primary key count given in DBMapper definition!")
            End Try
            insertIfMissing = Convert.ToBoolean(getParamFromXML(definitionXML, "insertIfMissing"))
            executeAdditionalProc = getParamFromXML(definitionXML, "executeAdditionalProc")
            ignoreColumns = getParamFromXML(definitionXML, "ignoreColumns")
            IgnoreDataErrors = Convert.ToBoolean(If(getParamFromXML(definitionXML, "IgnoreDataErrors") = "", "False", getParamFromXML(definitionXML, "IgnoreDataErrors")))
            CUDFlags = Convert.ToBoolean(getParamFromXML(definitionXML, "CUDFlags"))
            If Not IsNothing(TargetRange.ListObject) Then
                ' special grey table style for CUDFlags DBMapper
                If CUDFlags Then
                    TargetRange.ListObject.TableStyle = fetchSetting("DBMapperCUDFlagStyle", "TableStyleLight11")
                    ' otherwise blue
                Else
                    TargetRange.ListObject.TableStyle = fetchSetting("DBMapperStandardStyle", "TableStyleLight9")
                End If
            End If
            ' only allow CUDFlags only on DBMappers that have underlying Listobjects that were created with a query
            If CUDFlags And (IsNothing(TargetRange.ListObject) OrElse TargetRange.ListObject.SourceType <> Excel.XlListObjectSourceType.xlSrcQuery) Then
                CUDFlags = False
                definitionXML.SelectSingleNode("ns0:CUDFlags").Delete()
                definitionXML.AppendChildNode("CUDFlags", NamespaceURI:="DBModifDef", NodeValue:="False")
                Throw New Exception("CUDFlags only supported for DBMappers on ListObjects (created with DBSetQueryListObject)!")
            End If
        Catch ex As Exception
            ErrorMsg("Error when creating DBMapper '" + dbmodifName + "': " + ex.Message, "DBModifier Definitions Error")
        End Try
    End Sub

    Public Function hadChanges() As Boolean
        Return changesDone
    End Function

    ''' <summary>simply open a database connection, required for DBBegin Transaction (from next step)</summary>
    ''' <returns></returns>
    Public Overrides Function openDatabase() As Boolean
        Dim env As Integer = getEnv(Globals.selectedEnvironment + 1)
        openDatabase = True
        If IsNothing(dbcnn) Then
            Return openConnection(env, database)
        End If
    End Function

    ''' <summary>inserts CUD (Create/Update/Delete) Marks at the right end of the DBMapper range</summary>
    ''' <param name="changedRange">passed TargetRange by Change Event or delete button</param>
    ''' <param name="deleteFlag">if delete button was pressed, this is true</param>
    Public Sub doCUDMarks(changedRange As Excel.Range, Optional deleteFlag As Boolean = False)
        If Not CUDFlags Then Exit Sub
        ' sanity check for single cell DB Mappers..
        If TargetRange.Columns.Count = 1 And TargetRange.Rows.Count = 1 Then
            Dim retval As MsgBoxResult = QuestionMsg(theMessage:="DB Mapper Range with CUD Flags is only one cell, really set CUD Flags ?", questionTitle:="Set CUD Flags for DB Mapper")
            If retval = vbCancel Then Exit Sub
        End If
        Dim changedRangeRows As Integer = changedRange.Rows.Count
        If changedRangeRows > CInt(fetchSetting("maxRowCountCUD", "10000")) Then
            If Not QuestionMsg(theMessage:="A large range was changed (" + changedRange.Rows.Count.ToString + " > maxRowCountCUD:" + fetchSetting("maxRowCountCUD", "10000") + "), this will probably lead to CUD flag setting taking very long. Continue?", questionTitle:="Set CUD Flags for DB Mapper") Then Exit Sub
        End If
        If changedRange.Parent.ProtectContents Then
            ErrorMsg("Worksheet " + changedRange.Parent.Name + " is content protected, can't set CUD Flags !")
            Exit Sub
        End If
        ExcelDnaUtil.Application.AutoCorrect.AutoExpandListRange = False ' to prevent automatic creation of new column
        ' DBMapper ranges relative to start of TargetRange and respecting a header row, so CUDMarkRow = changedRange.Row - TargetRange.Row + 1 ...
        If deleteFlag Then
            For Each changedRow As Excel.Range In changedRange.Rows
                Dim CUDMarkRow As Integer = changedRow.Row - TargetRange.Row + 1
                TargetRange.Cells(CUDMarkRow, TargetRange.Columns.Count + 1).Value = "d"
                TargetRange.Rows(CUDMarkRow).Font.Strikethrough = True
            Next
        Else
            For Each changedRow As Excel.Range In changedRange.Rows
                Dim CUDMarkRow As Integer = changedRow.Row - TargetRange.Row + 1
                ' change only if not already set or
                If IsNothing(TargetRange.Cells(CUDMarkRow, TargetRange.Columns.Count + 1).Value) Then
                    Dim RowContainsData As Boolean = False
                    For Each containedCell As Excel.Range In TargetRange.Rows(CUDMarkRow).Cells
                        ' check without newly inserted/updated cells (copy paste) 
                        Dim possibleIntersection As Excel.Range = ExcelDnaUtil.Application.Intersect(containedCell, changedRange)
                        ' check if whole row is empty (except for the changedRange), formulas do not count as filled (automatically filled for lookups or other things)..
                        If Not IsNothing(containedCell.Value) AndAlso IsNothing(possibleIntersection) AndAlso Left(containedCell.Formula, 1) <> "=" Then
                            RowContainsData = True
                            Exit For
                        End If
                    Next
                    ' if Row Contains Data (not every cell empty except currently modified (changedRange), this is for adding rows below data range) then "u"pdate
                    If RowContainsData Then
                        TargetRange.Cells(CUDMarkRow, TargetRange.Columns.Count + 1).Value = "u"
                        TargetRange.Rows(CUDMarkRow).Font.Italic = True
                    Else ' else "i"nsert
                        TargetRange.Cells(CUDMarkRow, TargetRange.Columns.Count + 1).Value = "i"
                    End If
                    ExcelDnaUtil.Application.Statusbar = changedRow.Row.ToString + "/" + changedRangeRows.ToString
                End If
            Next
        End If
        ExcelDnaUtil.Application.Statusbar = False
        ExcelDnaUtil.Application.AutoCorrect.AutoExpandListRange = True
    End Sub

    ''' <summary>extend DataRange to "whole" DBMApper area (first row (field names) to the right and first column (primary key) down)</summary>
    Public Sub extendDataRange()
        Dim NamesList As Excel.Names = ExcelDnaUtil.Application.ActiveWorkbook.Names
        ' only extend like this if no CUD Flags present (may have non existing first (primary) columns -> auto identity columns !)
        If Not CUDFlags Then
            If IsNothing(TargetRange.Cells(2, 1).Value) Then Exit Sub ' only extend if there are multiple rows...
            preventChangeWhileFetching = True
            Dim rowEnd As Integer = TargetRange.Cells(1, 1).End(Excel.XlDirection.xlDown).Row
            ' unfortunately the above method to find the column extent doesn't work with hidden columns, so count the filled cells directly...
            Dim colEnd As Integer = 1
            While Not (IsNothing(TargetRange.Cells(1, colEnd + 1).Value) OrElse TargetRange.Cells(1, colEnd + 1).Value = "")
                colEnd += 1
            End While
            Try : NamesList.Add(Name:=paramTargetName, RefersTo:=TargetRange.Parent.Range(TargetRange.Cells(1, 1), TargetRange.Cells(rowEnd, colEnd)))
            Catch ex As Exception
                Throw New Exception("Error when reassigning name '" + paramTargetName + "' to DBMapper while extending DataRange: " + ex.Message)
            Finally
                preventChangeWhileFetching = False
            End Try
        Else
            ' CUD Flags are only allowed with DBMappers in ListObjects
            Try : NamesList.Add(Name:=paramTargetName, RefersTo:=TargetRange.ListObject.Range)
            Catch ex As Exception
                Throw New Exception("Error when reassigning name '" + paramTargetName + "' to DBMapper while extending DataRange: " + ex.Message)
            End Try
        End If
        ' the Data range might have been extended (by inserting rows), so reassign it to the TargetRange
        Try
            TargetRange = TargetRange.Parent.Range(paramTargetName)
        Catch ex As Exception
            Throw New Exception("Error when setting name '" + paramTargetName + "' to TargetRange while extending DataRange: " + ex.Message)
        End Try
    End Sub

    ''' <summary>reset CUD FLags, either after completion of doDBModif or on request (refresh)</summary>
    Public Sub resetCUDFlags()
        ' in case CUDFlags was set to a single cell DBMapper avoid resetting CUDFlags
        If CUDFlags And Not (TargetRange.Columns.Count = 1 And TargetRange.Rows.Count = 1) Then
            If TargetRange.Parent.ProtectContents Then
                ErrorMsg("Worksheet " + TargetRange.Parent.Name + " is content protected, can't reset CUD Flags !")
                Exit Sub
            End If
            ExcelDnaUtil.Application.AutoCorrect.AutoExpandListRange = False ' to prevent automatic creation of new column
            TargetRange.Columns(TargetRange.Columns.Count + 1).ClearContents
            TargetRange.Font.Italic = False
            TargetRange.Font.Strikethrough = False
            ExcelDnaUtil.Application.AutoCorrect.AutoExpandListRange = True ' to prevent automatic creation of new column
        End If
    End Sub

    Public Overrides Sub doDBModif(Optional WbIsSaving As Boolean = False, Optional calledByDBSeq As String = "", Optional TransactionOpen As Boolean = False)
        changesDone = False
        ' ask for saving only if a) is not done on WorkbookSave b) is set to ask and c) is not called by a DBSequence (asks already for saving)
        If Not WbIsSaving And askBeforeExecute And calledByDBSeq = "" Then
            If confirmText = "" Then confirmText = "Really execute DB Mapper " + dbmodifName + "?"
            Dim retval As MsgBoxResult = QuestionMsg(theMessage:=confirmText, questionTitle:="Execute DB Mapper")
            If retval = vbCancel Then Exit Sub
        End If
        ' if Environment is not existing, take from selectedEnvironment
        Dim env As Integer = getEnv(Globals.selectedEnvironment + 1)
        extendDataRange()
        ' check for mass changes and warn if necessary
        If CUDFlags Then
            Dim maxMassChanges As Integer = CInt(fetchSetting("maxNumberMassChange", "30"))
            Dim curWs As Excel.Worksheet = TargetRange.Parent ' this is necessary because using TargetRange directly deletes the content of the CUD area !!
            Dim changesToBeDone As Integer = ExcelDnaUtil.Application.WorksheetFunction.CountIf(curWs.Range(TargetRange.Columns(TargetRange.Columns.Count + 1).Address), "<>")
            If changesToBeDone > maxMassChanges Then
                Dim retval As MsgBoxResult = QuestionMsg(theMessage:="Modifying more rows (" + changesToBeDone + ") than defined warning limit (" + maxMassChanges + "), continue?", questionTitle:="Execute DB Mapper")
                If retval = vbCancel Then Exit Sub
            End If
        End If
        'now create/get a connection (dbcnn) for env(ironment) in case it was not already created by a step in the sequence before (transactions!)
        If Not TransactionOpen Then
            If Not openConnection(env, database) Then Exit Sub
        End If

        'checkrst is opened to get information about table schema (field types)
        Dim checkrst As ADODB.Recordset = New ADODB.Recordset
        Dim rst As ADODB.Recordset = New ADODB.Recordset
        Try
            checkrst.Open(tableName, dbcnn, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockReadOnly, CommandTypeEnum.adCmdTableDirect)
        Catch ex As Exception
            hadError = True
            ErrorMsg("Opening table '" + tableName + "' caused following error: " + ex.Message + " for DBMapper " + paramTargetName, "DBMapper Error")
            GoTo cleanup
        End Try

        ' check if all column names (except ignored) of DBMapper Range exist in table
        Dim colNum As Long = 1
        Do
            Dim fieldname As String = Trim(TargetRange.Cells(1, colNum).Value)
            ' only if not ignored...
            If InStr(1, LCase(ignoreColumns) + ",", LCase(fieldname) + ",") = 0 Then
                Try
                    Dim testExist As String = checkrst.Fields(fieldname).Name
                Catch ex As Exception
                    hadError = True
                    TargetRange.Parent.Activate
                    TargetRange.Cells(1, colNum).Select
                    ErrorMsg("Field '" + fieldname + "' does not exist in Table '" + tableName + "' and is not in ignoreColumns, Error in sheet " + TargetRange.Parent.Name, "DBMapper Error")
                    GoTo cleanup
                End Try
            End If
            colNum += 1
        Loop Until colNum > TargetRange.Columns.Count

        Dim rowNum As Long = 2
        dbcnn.CursorLocation = CursorLocationEnum.adUseServer

        Dim finishLoop As Boolean
        '''''''''''''''''''''''''''''''''''''''' walk through rows
        Do
            ' if CUDFlags are set, only insert/update/delete if CUDFlags column (right to DBMapper range) is filled...
            Dim rowCUDFlag As String = TargetRange.Cells(rowNum, TargetRange.Columns.Count + 1).Value
            If Not CUDFlags Or (CUDFlags And rowCUDFlag <> "") Then
                Dim AutoIncrement As Boolean = False
                ' try to find record for update, construct WHERE clause with primary key columns
                Dim primKeyCompound As String = " WHERE "
                Dim primKeyDisplay As String = ""
                For i As Integer = 1 To primKeysCount
                    Dim primKeyValue = TargetRange.Cells(rowNum, i).Value
                    If IsXLCVErr(primKeyValue) Then
                        If Not notifyUserOfDataError("Error in primary key value: " + CVErrText(primKeyValue) + ", cell (" + rowNum.ToString + "," + i.ToString + ") in sheet " + TargetRange.Parent.Name + " and row " + rowNum.ToString, rowNum, i) Then GoTo cleanup
                        GoTo nextRow
                    End If
                    If IsNothing(primKeyValue) Then primKeyValue = ""
                    Dim primKey = TargetRange.Cells(1, i).Value
                    ' if primKey is in ignoreColumns then the only reasonable reason is a lookup primary key in DBSHeets (CUDFlags only), so try with "real" (resolved key) instead...
                    If InStr(1, LCase(ignoreColumns) + ",", LCase(primKey) + ",") > 0 Then
                        If CUDFlags Then
                            primKey = Left(primKey, Len(primKey) - 2) ' correct the LU to real primary Key
                            primKeyValue = TargetRange.ListObject.ListColumns(primKey).Range(rowNum, 1).Value ' get the value from there
                        Else
                            If Not notifyUserOfDataError("Primary key '" + primKey + "' contained in ignoreColumns !", rowNum, i) Then GoTo cleanup
                            GoTo nextRow
                        End If
                    End If
                    Dim checkAutoIncrement As Boolean
                    Try : checkAutoIncrement = checkrst.Fields(primKey).Properties("IsAutoIncrement").Value
                    Catch ex As Exception
                        If Not notifyUserOfDataError("Primary key '" + primKey.ToString + "' not found in table '" + tableName + "' or IsAutoIncrement Property not provided:" + ex.Message, rowNum, i) Then GoTo cleanup
                        GoTo nextRow
                    End Try
                    If primKeysCount = 1 And CUDFlags And primKeyValue.ToString().Length = 0 And checkAutoIncrement Then
                        AutoIncrement = True
                        Exit For
                    End If
                    ' with CUDFlags there can be empty primary keys (auto identity columns), leave error checking to database in this case ...
                    If (Not CUDFlags Or (CUDFlags And rowCUDFlag = "u")) And primKeyValue.ToString().Length = 0 Then
                        If Not notifyUserOfDataError("Empty primary key value, cell (" + rowNum.ToString + "," + i.ToString + ") in sheet " + TargetRange.Parent.Name + " and row " + rowNum.ToString, rowNum, i) Then GoTo cleanup
                        GoTo nextRow
                    End If
                    ' now format the primary key value and construct the WHERE clause
                    Dim primKeyFormatted As String
                    If IsNothing(primKeyValue) Then
                        primKeyFormatted = "NULL"
                    ElseIf checkIsNumeric(checkrst.Fields(primKey).Type) Then ' only decimal points allowed in numeric data
                        primKeyFormatted = Replace(CStr(primKeyValue), ",", ".")
                    ElseIf checkIsDate(checkrst.Fields(primKey).Type) Then
                        If TypeName(primKeyValue) = "Date" Then ' received as a Date value already
                            primKeyFormatted = "'" + Format(primKeyValue, "yyyy-MM-dd") + "'" ' ISO 8601 standard SQL Date formatting
                        ElseIf TypeName(primKeyValue) = "Double" Then ' got a double
                            primKeyFormatted = "'" + Format(Date.FromOADate(primKeyValue), "yyyy-MM-dd") + "'" ' ISO 8601 standard SQL Date formatting
                        Else
                            notifyUserOfDataError("provided value neither Date nor Double, cannot convert into formatted primary key for lookup !", rowNum, i)
                            GoTo cleanup
                        End If
                    ElseIf checkIsTime(checkrst.Fields(primKey).Type) Then
                        If TypeName(primKeyValue) = "Date" Then
                            primKeyFormatted = "'" + Format(primKeyValue, "yyyy-MM-dd HH:mm:ss.fff") + "'" ' ISO 8601 standard SQL Date/time formatting, 24h format...
                        ElseIf TypeName(primKeyValue) = "Double" Then
                            primKeyFormatted = "'" + Format(Date.FromOADate(primKeyValue), "yyyy-MM-dd HH:mm:ss.fff") + "'" ' ISO 8601 standard SQL Date/time formatting, 24h format...
                        Else
                            notifyUserOfDataError("provided value neither Date nor Double, cannot convert into formatted primary key for lookup !", rowNum, i)
                            GoTo cleanup
                        End If
                    ElseIf TypeName(primKeyValue) = "Boolean" Then
                        primKeyFormatted = IIf(primKeyValue, "1", "0")
                    Else
                        primKeyFormatted = "'" + Replace(primKeyValue, "'", "''") + "'" ' quote quotes inside Strings and surround result with quotes
                    End If
                    primKeyCompound = primKeyCompound + primKey.ToString + " = " + primKeyFormatted + IIf(i = primKeysCount, "", " AND ")
                Next
                ' get the record for updating, however avoid opening recordset with empty primary key value if autoincrement is given...
                Dim getStmt As String = "SELECT * FROM " + tableName + primKeyCompound
                If Not AutoIncrement Then
                    Try
                        rst.Open(getStmt, dbcnn, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)
                        Dim check As Boolean = rst.EOF
                    Catch ex As Exception
                        notifyUserOfDataError("Problem getting recordset, Error: " + ex.Message + " in sheet " + TargetRange.Parent.Name + " and row " + rowNum.ToString + ", doing " + getStmt, rowNum)
                        GoTo cleanup
                    End Try
                    primKeyDisplay = Replace(Mid(primKeyCompound, 7), " AND ", ";")
                Else
                    ' just open the table if autoincrement set and empty primary key
                    rst.Open(tableName, dbcnn, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)
                End If

                ' If we have an autoincrementing primary key (empty primary key value !) or didn't find a record with the given primary key (rst.EOF),
                ' add a new record if insertIfMissing flag is set Or CUD Flag insert is given
                If AutoIncrement OrElse rst.EOF Then
                    Dim i As Integer
                    If insertIfMissing Or rowCUDFlag = "i" Then
                        ExcelDnaUtil.Application.StatusBar = Left("Inserting " + primKeyDisplay + " into " + tableName, 255)
                        rst.AddNew()
                        For i = 1 To primKeysCount
                            Dim primKeyValue As Object = TargetRange.Cells(rowNum, i).Value
                            If IsNothing(primKeyValue) Then primKeyValue = ""
                            Dim primKey = TargetRange.Cells(1, i).Value
                            ' if primKey is in ignoreColumns then the only reasonable reason is a lookup primary key in DBSHeets (CUDFlags only), so try with "real" (resolved key) instead...
                            If InStr(1, LCase(ignoreColumns) + ",", LCase(primKey) + ",") > 0 AndAlso CUDFlags Then
                                primKey = Left(primKey, Len(primKey) - 2) ' correct the LU to real primary Key
                                primKeyValue = TargetRange.ListObject.ListColumns(primKey).Range(rowNum, 1).Value ' get the value from there
                            End If
                            Try
                                ' skip empty primary field values for autoincrementing identity fields ..
                                If CStr(primKeyValue) <> "" Then rst.Fields(primKey).Value = primKeyValue
                            Catch ex As Exception
                                If Not notifyUserOfDataError("Error inserting primary key value into table " + tableName + ": " + dbcnn.Errors(0).Description, rowNum, i) Then GoTo cleanup
                            End Try
                        Next
                    Else
                        If Not notifyUserOfDataError("Did not find recordset with statement '" + getStmt + "', insertIfMissing = " + insertIfMissing.ToString() + " in sheet " + TargetRange.Parent.Name + " and row " + rowNum.ToString, rowNum, i) Then GoTo cleanup
                    End If
                    ExcelDnaUtil.Application.StatusBar = Left("Updating " + primKeyDisplay + " in " + tableName, 255)
                End If

                If Not CUDFlags Or (CUDFlags And (rowCUDFlag = "i" Or rowCUDFlag = "u")) Then
                    ' walk through non primary columns and fill fields
                    colNum = primKeysCount + 1
                    Do
                        Dim fieldname As String = TargetRange.Cells(1, colNum).Value
                        If InStr(1, LCase(ignoreColumns) + ",", LCase(fieldname) + ",") = 0 Then
                            Try
                                Dim fieldval As Object = TargetRange.Cells(rowNum, colNum).Value
                                If IsNothing(fieldval) Then
                                    rst.Fields(fieldname).Value = Nothing
                                Else
                                    If IsXLCVErr(fieldval) Then
                                        If IgnoreDataErrors Then
                                            rst.Fields(fieldname).Value = Nothing
                                        Else
                                            If Not notifyUserOfDataError("Field Value Update Error: " + CVErrText(fieldval) + " with Table: " + tableName + ", Field: " + fieldname + ", in sheet " + TargetRange.Parent.Name + " and row " + rowNum.ToString + ", col: " + colNum.ToString, rowNum, colNum) Then GoTo cleanup
                                        End If
                                    Else
                                        ' special treatment for time and date fields as they are
                                        If TypeName(fieldval) = "Date" Then
                                            Dim fieldValDate As DateTime = DirectCast(fieldval, DateTime)
                                            If fieldValDate.Hour() = 0 And fieldValDate.Minute() = 0 And fieldValDate.Second() = 0 And fieldValDate.Millisecond() = 0 Then
                                                fieldval = Format(fieldValDate, "yyyy-MM-dd")  ' ISO 8601 standard SQL Date formatting
                                            Else
                                                fieldval = Format(fieldValDate, "yyyy-MM-dd HH:mm:ss")  ' ISO 8601 standard SQL Date/time formatting, 24h format...
                                            End If
                                        End If
                                        rst.Fields(fieldname).Value = IIf(fieldval.ToString().Length = 0, Nothing, fieldval)
                                    End If
                                End If
                            Catch ex As Exception
                                Dim exMessage As String = ex.Message
                                rst.CancelUpdate()
                                If Not notifyUserOfDataError("Field Value Update Error: " + exMessage + " with Table: " + tableName + ", Field: " + fieldname + ", in sheet " + TargetRange.Parent.Name + " and row " + rowNum.ToString + ", col: " + colNum.ToString, rowNum, colNum) Then GoTo cleanup
                            End Try
                        End If
                        colNum += 1
                    Loop Until colNum > TargetRange.Columns.Count

                    ' now do the update/insert in the DB
                    Try
                        rst.Update()
                        changesDone = True
                        ' remove CUD Flag if present
                        If CUDFlags Then TargetRange.Cells(rowNum, TargetRange.Columns.Count + 1).Value = ""
                    Catch ex As Exception
                        Dim exMessage As String = ex.Message
                        rst.CancelUpdate()
                        If Not notifyUserOfDataError("Row Update Error, Table: " + rst.Source.ToString + ", Error: " + exMessage + " in sheet " + TargetRange.Parent.Name + " and row " + rowNum.ToString, rowNum) Then GoTo cleanup
                    End Try
                End If
                If (CUDFlags And rowCUDFlag = "d") Then
                    ExcelDnaUtil.Application.StatusBar = Left("Deleting " + primKeyDisplay + " in " + tableName, 255)
                    Try
                        rst.Delete(AffectEnum.adAffectCurrent)
                        changesDone = True
                        ' remove CUD Flag if present
                        TargetRange.Cells(rowNum, TargetRange.Columns.Count + 1).Value = ""
                    Catch ex As Exception
                        If Not notifyUserOfDataError("Error deleting row " + rowNum.ToString + " in sheet " + TargetRange.Parent.Name + ": " + ex.Message, rowNum) Then GoTo cleanup
                    End Try
                End If
                rst.Close()
nextRow:
                Try
                    If IsNothing(TargetRange.Cells(rowNum + 1, 1).Value) OrElse TargetRange.Cells(rowNum + 1, 1).Value.ToString().Length = 0 Then finishLoop = True
                Catch ex As Exception
                    If Not notifyUserOfDataError("Error in first primary column: Cells(" + (rowNum + 1).ToString + ",1): " + ex.Message, rowNum + 1) Then GoTo cleanup
                    'finishLoop = True '-> do not finish to allow erroneous data  !!
                End Try
            End If
            rowNum += 1
        Loop Until rowNum > TargetRange.Rows.Count Or (finishLoop And Not CUDFlags)

        ' any additional stored procedures to execute?
        If executeAdditionalProc.Length > 0 Then
            Try
                ExcelDnaUtil.Application.StatusBar = "executing stored procedure " + executeAdditionalProc
                dbcnn.Execute(executeAdditionalProc)
            Catch ex As Exception
                hadError = True
                ErrorMsg("Error in executing additional stored procedure: " + ex.Message, "DBMapper Error")
                GoTo cleanup
            End Try
        End If
cleanup:
        ExcelDnaUtil.Application.StatusBar = False
        ' close connection to return it to the pool (automatically closes recordset objects)...
        If calledByDBSeq = "" Then dbcnn.Close()
        ' DBSheet surrogate (CUDFlags), ask for refresh after DB Modification was done
        If changesDone Then
            Dim DBFunctionSrcExtent = getDBunderlyingNameFromRange(TargetRange)
            If DBFunctionSrcExtent <> "" Then
                If CUDFlags Then
                    ' also resetCUDFlags for CUDFlags DBMapper that do not ask before execute and were called by a DBSequence
                    Try
                        ' reset CUDFlags before refresh to avoid problems with reduced TargetRange due to deletions!
                        Me.resetCUDFlags()
                    Catch ex As Exception
                        ErrorMsg("Error in resetting CUD Flags: " + ex.Message, "DBMapper Error")
                    End Try
                    If askBeforeExecute AndAlso calledByDBSeq = "" Then
                        Dim retval As MsgBoxResult = QuestionMsg(theMessage:="Refresh Data Range of DB Mapper '" + dbmodifName + "' ?", questionTitle:="Refresh DB Mapper")
                        If retval = vbOK Then
                            doDBRefresh(Replace(DBFunctionSrcExtent, "DBFtarget", "DBFsource"))
                            ' clear CUD marks after completion is done with doDBRefresh/DBSetQueryAction/resizeDBMapperRange
                        End If
                    Else
                        LogWarn("no refresh took place for DBMapper " + dbmodifName)
                    End If
                End If
            End If
        End If
    End Sub

    Private Function notifyUserOfDataError(message As String, rowNum As Long, Optional colNum As Integer = -1) As Boolean
        hadError = True
        If Not nonInteractive Then
            TargetRange.Parent.Activate
            If colNum = -1 Then
                TargetRange.Rows(rowNum).Select
            Else
                TargetRange.Cells(rowNum, colNum).Select
            End If
        End If
        Dim retval As MsgBoxResult = QuestionMsg(message, MsgBoxStyle.OkCancel, "DBMapper Error", MsgBoxStyle.Critical)
        If retval = vbCancel Then Return False
        Return True
    End Function

    Public Overrides Sub setDBModifCreateFields(ByRef theDBModifCreateDlg As DBModifCreate)
        With theDBModifCreateDlg
            .envSel.SelectedIndex = getEnv() - 1
            .TargetRangeAddress.Text = getTargetRangeAddress()
            .Database.Text = database
            .execOnSave.Checked = execOnSave
            .Tablename.Text = tableName
            .PrimaryKeys.Text = primKeysCount.ToString()
            .insertIfMissing.Checked = insertIfMissing
            .addStoredProc.Text = executeAdditionalProc
            .IgnoreColumns.Text = ignoreColumns
            .CUDflags.Checked = CUDFlags
            .IgnoreDataErrors.Checked = IgnoreDataErrors
            .AskForExecute.Checked = askBeforeExecute
        End With
    End Sub

End Class

''' <summary>DBActions are used to issue DML commands defined in Cells against the database</summary>
Public Class DBAction : Inherits DBModif

    ''' <summary>Database to store to</summary>
    Private database As String
    ''' <summary>DBModif name of target range</summary>
    Private paramTargetName As String

    ''' <summary>normal constructor with definition xml</summary>
    ''' <param name="definitionXML"></param>
    ''' <param name="paramTarget"></param>
    Public Sub New(definitionXML As CustomXMLNode, paramTarget As Excel.Range)
        MyBase.New(definitionXML)
        Try
            ' if no target range is set, then no parameters can be found...
            If IsNothing(paramTarget) Then Exit Sub
            paramTargetName = getDBModifNameFromRange(paramTarget)
            If Left(paramTargetName, 8) <> "DBAction" Then Throw New Exception("target " + paramTargetName + " not matching passed DBModif type DBAction for " + getTargetRangeAddress() + "/" + dbmodifName + " !")
            TargetRange = paramTarget
            ' fill parameters from definition
            env = getParamFromXML(definitionXML, "env")
            database = getParamFromXML(definitionXML, "database")
            If database = "" Then Throw New Exception("No database given in DBAction definition!")
            ' AFTER setting TargetRange and all the rest check for defined action to have a decent getTargetRangeAddress for undefined actions...
            If paramTarget.Cells(1, 1).Text = "" Then Throw New Exception("No Action defined in " + paramTargetName + "(" + getTargetRangeAddress() + ")")
        Catch ex As Exception
            ErrorMsg("Error when creating DB Sequence '" + dbmodifName + "': " + ex.Message, "DBModifier Definitions Error")
        End Try
    End Sub

    ''' <summary>simply open a database connection, required for DBBegin Transaction (from next step)</summary>
    ''' <returns></returns>
    Public Overrides Function openDatabase() As Boolean
        openDatabase = True
        If IsNothing(dbcnn) Then
            Return openConnection(env, database)
        End If
    End Function

    Public Overrides Sub doDBModif(Optional WbIsSaving As Boolean = False, Optional calledByDBSeq As String = "", Optional TransactionOpen As Boolean = False)
        ' ask for saving only if a) is not done on WorkbookSave b) is set to ask and c) is not called by a DBSequence (asks already for saving)
        If Not WbIsSaving And askBeforeExecute And calledByDBSeq = "" Then
            If confirmText = "" Then confirmText = "Really execute DB Action " + dbmodifName + "?"
            Dim retval As MsgBoxResult = QuestionMsg(theMessage:=confirmText, questionTitle:="Execute DB Action")
            If retval = vbCancel Then Exit Sub
        End If
        ' if Environment is not existing, take from selectedEnvironment
        Dim env As Integer = getEnv(Globals.selectedEnvironment + 1)

        'Dim ds As DataSet = New DataSet()
        'Dim dataAdapter As SqlDataAdapter = New SqlDataAdapter()
        'Dim theConnString As String = fetchSetting("ConstConnString" + env, "")
        'Dim dbidentifier As String = fetchSetting("DBidentifierCCS" + env, "")
        'theConnString = Change(theConnString, dbidentifier, database, ";")
        'Dim cn As SqlConnection = New SqlConnection(theConnString)
        'cn.Open()

        'Dim trans As SqlTransaction = cn.BeginTransaction

        'dataAdapter.InsertCommand.Transaction = trans
        'dataAdapter.UpdateCommand.Transaction = trans
        'dataAdapter.DeleteCommand.Transaction = trans

        'Try
        '    dataAdapter.Update(ds)
        '    trans.Commit()
        'Catch ex As Exception
        '    trans.Rollback()
        'End Try
        'cn.Close()
        'Exit Sub

        'now create/get a connection (dbcnn) for env(ironment) in case it was not already created by the sequence (transactions!)
        If Not TransactionOpen Then
            If Not openConnection(env, database) Then Exit Sub
        End If
        Dim result As Long = 0
        Try
            ExcelDnaUtil.Application.StatusBar = "executing DBAction " + paramTargetName
            Dim executeText As String = ""
            For Each targetCell As Excel.Range In TargetRange
                executeText += targetCell.Text + " "
            Next
            dbcnn.Execute(executeText, result, Options:=CommandTypeEnum.adCmdText)
        Catch ex As Exception
            hadError = True
            ErrorMsg("Error: " + paramTargetName + ": " + ex.Message, "DBAction Error")
            ExcelDnaUtil.Application.StatusBar = False
            Exit Sub
        End Try
        If Not WbIsSaving And calledByDBSeq = "" Then
            ErrorMsg("DBAction " + paramTargetName + " executed, affected records: " + result.ToString, "DBAction executed", MsgBoxStyle.Information)
        End If
        ExcelDnaUtil.Application.StatusBar = False
        ' close connection to return it to the pool...
        If calledByDBSeq = "" Then dbcnn.Close()
    End Sub

    Public Overrides Sub setDBModifCreateFields(ByRef theDBModifCreateDlg As DBModifCreate)
        With theDBModifCreateDlg
            .envSel.SelectedIndex = getEnv() - 1
            .TargetRangeAddress.Text = getTargetRangeAddress()
            .Database.Text = database
            .execOnSave.Checked = execOnSave
            .AskForExecute.Checked = askBeforeExecute
        End With
    End Sub
End Class

''' <summary>DBSequences are used to group DBMappers and DBActions and run them in sequence together with refreshing DBFunctions and executing them in transaction brackets</summary>
Public Class DBSeqnce : Inherits DBModif

    ''' <summary>sequence of DB Mappers, DB Actions and DB Refreshes being executed in this sequence</summary>
    Private sequenceParams() As String = {}

    ''' <summary>normal constructor with definition xml</summary>
    ''' <param name="definitionXML"></param>
    Public Sub New(definitionXML As CustomXMLNode)
        MyBase.New(definitionXML)
        Try
            Dim seqSteps As Integer = definitionXML.SelectNodes("ns0:seqStep").Count
            If seqSteps = 0 Then
                Throw New Exception("no steps defined in DBSequence definition!")
            Else
                ReDim sequenceParams(seqSteps - 1)
                For i = 1 To seqSteps
                    sequenceParams(i - 1) = definitionXML.SelectNodes("ns0:seqStep")(i).Text
                Next
            End If
        Catch ex As Exception
            ErrorMsg("Error when creating DB Sequence '" + dbmodifName + "': " + ex.Message, "DBModifier Definitions Error")
        End Try
    End Sub

    Public Overrides Sub doDBModif(Optional WbIsSaving As Boolean = False, Optional calledByDBSeq As String = "", Optional TransactionOpen As Boolean = False)
        Dim TransactionIsOpen As Boolean = False
        hadError = False
        ' warning against recursions (should not happen...)
        If calledByDBSeq <> "" Then
            ErrorMsg("DB Sequence '" + dbmodifName + "' is being called by another DB Sequence (" + calledByDBSeq + "), this should not occur as infinite recursions are possible !", "Execute DB Sequence")
            Exit Sub
        End If
        ' ask for saving only if a) is not done on WorkbookSave b) is set to ask and c) is not called by a DBSequence (asks already for saving)
        If Not WbIsSaving And askBeforeExecute Then
            If confirmText = "" Then confirmText = "Really execute DB Sequence " + dbmodifName + "?"
            Dim retval As MsgBoxResult = QuestionMsg(theMessage:=confirmText, questionTitle:="Execute DB Sequence")
            If retval = vbCancel Then Exit Sub
        End If
        ' reset the db connection in any case to allow for new connections at DBBegin
        dbcnn = Nothing
        Dim executedDBMappers As Dictionary(Of String, Boolean) = New Dictionary(Of String, Boolean)
        Dim modifiedDBMappers As Dictionary(Of String, Boolean) = New Dictionary(Of String, Boolean)
        For i As Integer = 0 To UBound(sequenceParams)
            Dim definition() As String = Split(sequenceParams(i), ":")
            Dim DBModiftype As String = definition(0)
            Dim DBModifname As String = definition(1)
            Select Case DBModiftype
                Case "DBMapper", "DBAction"
                    LogInfo(DBModifname + "... ")
                    DBModifDefColl(DBModiftype).Item(DBModifname).doDBModif(WbIsSaving, calledByDBSeq:=MyBase.dbmodifName, TransactionOpen:=TransactionIsOpen)
                    If DBModiftype = "DBMapper" Then
                        If DirectCast(DBModifDefColl("DBMapper").Item(DBModifname), DBMapper).CUDFlags Then executedDBMappers(DBModifname) = True
                        If DirectCast(DBModifDefColl("DBMapper").Item(DBModifname), DBMapper).hadChanges Then modifiedDBMappers(DBModifname) = True
                    End If
                Case "DBBegin"
                    LogInfo("DBBeginTrans... ")
                    If IsNothing(dbcnn) Then
                        ' take database connection properties from next sequence step
                        Dim nextdefinition() As String = Split(sequenceParams(i + 1), ":")
                        If Not DBModifDefColl(nextdefinition(0)).Item(nextdefinition(1)).openDatabase() Then Exit Sub
                    End If
                    'TODO: migrate ADODB to ADO.net
                    dbcnn.BeginTrans()
                    TransactionIsOpen = True
                Case "DBCommitRollback"
                    If Not hadError Then
                        LogInfo("DBCommitTrans... ")
                        dbcnn.CommitTrans()
                    Else
                        LogInfo("DBRollbackTrans... ")
                        dbcnn.RollbackTrans()
                    End If
                    TransactionIsOpen = False
                Case Else
                    If Left(DBModiftype, 8) = "Refresh " Then
                        doDBRefresh(srcExtent:=DBModifname, executedDBMappers:=executedDBMappers, modifiedDBMappers:=modifiedDBMappers, TransactionIsOpen:=TransactionIsOpen)
                    Else
                        ErrorMsg("Unknown type of sequence step '" + DBModiftype + "' being called in DB Sequence (" + calledByDBSeq + ") !", "Execute DB Sequence")
                    End If
            End Select
        Next
    End Sub

    Public Function getSequenceSteps() As String()
        Return sequenceParams
    End Function


    Public Overrides Sub setDBModifCreateFields(ByRef theDBModifCreateDlg As DBModifCreate)
        With theDBModifCreateDlg
            .execOnSave.Checked = execOnSave
            .AskForExecute.Checked = askBeforeExecute
            For i As Integer = 0 To UBound(sequenceParams)
                .DBSeqenceDataGrid.Rows.Add(sequenceParams(i))
                ' prepare sequence for repairing in case there are entries that do not match the existing definitions 
                ' (especially for DBrefresh of DBfunctions this might happen)
                .RepairDBSeqnce.Text += sequenceParams(i) + IIf(i = UBound(sequenceParams), "", vbCrLf)
            Next
        End With
    End Sub

End Class


''' <summary>global helper functions for DBModifiers</summary>
Public Module DBModifs

    ''' <summary>main db connection For mapper</summary>
    Public dbcnn As ADODB.Connection
    ''' <summary>avoid entering Application.SheetChange Event handler during listfetch/setquery</summary>
    Public preventChangeWhileFetching As Boolean = False
    ''' <summary>indicates an error in execution, used for commit/rollback</summary>
    Public hadError As Boolean
    ''' <summary>used to work around the fact that when started by Application.Run, Formulas are sometimes returned as local</summary>
    Public listSepLocal As String = ExcelDnaUtil.Application.International(Excel.XlApplicationInternational.xlListSeparator)

    ''' <summary>opens a database connection</summary>
    ''' <param name="env">number of the environment as given in the settings</param>
    ''' <param name="database">database to replace database selection parameter in connection string of environment</param>
    ''' <returns>True on success</returns>
    Public Function openConnection(env As Integer, database As String) As Boolean
        openConnection = False

        Dim theConnString As String = fetchSetting("ConstConnString" + env.ToString, "")
        If theConnString = "" Then
            ErrorMsg("No Connectionstring given for environment: " + env.ToString + ", please correct and rerun.", "Open Connection Error")
            Exit Function
        End If
        Dim dbidentifier As String = fetchSetting("DBidentifierCCS" + env.ToString, "")
        If dbidentifier = "" Then
            ErrorMsg("No DB identifier given for environment: " + env.ToString + ", please correct and rerun.", "Open Connection Error")
            Exit Function
        End If

        ' connections are pooled by ADO depending on the connection string:
        dbcnn = New Connection
        theConnString = Change(theConnString, dbidentifier, database, ";")
        LogInfo("open connection with " + theConnString)
        ExcelDnaUtil.Application.StatusBar = "Trying " + Globals.CnnTimeout.ToString + " sec. with connstring: " + theConnString
        Try
            dbcnn.ConnectionString = theConnString
            dbcnn.ConnectionTimeout = Globals.CnnTimeout
            dbcnn.CommandTimeout = Globals.CmdTimeout
            dbcnn.Open()
            openConnection = True
        Catch ex As Exception
            ErrorMsg("Error connecting to DB: " + ex.Message + ", connection string: " + theConnString, "Open Connection Error")
            dbcnn = Nothing
        End Try
        ExcelDnaUtil.Application.StatusBar = False
    End Function

    ''' <summary>in case there is a defined DBMapper underlying the DBListFetch/DBSetQuery target area then change the extent of that to the new area given in theRange</summary>
    ''' <param name="theRange"></param>
    Public Sub resizeDBMapperRange(theRange As Excel.Range, oldRange As Excel.Range)
        ' only do this for the active workbook...
        If theRange.Parent.Parent Is ExcelDnaUtil.Application.Activeworkbook Then
            ' getDBModifNameFromRange gets any DBModifName (starting with DBMapper/DBAction...) intersecting theRange, so we can reassign it to the changed range with this...
            Dim dbMapperRangeName As String = getDBModifNameFromRange(theRange)
            ' only allow resizing of dbMapperRange if it was EXACTLY matching the FORMER target range of the DB Function
            If Left(dbMapperRangeName, 8) = "DBMapper" AndAlso oldRange.Address = ExcelDnaUtil.Application.Activeworkbook.Names(dbMapperRangeName).RefersToRange.Address Then
                ' (re)assign db mapper range name to the passed (changed) DBListFetch/DBSetQuery function target range
                Try : theRange.Name = dbMapperRangeName
                Catch ex As Exception
                    Throw New Exception("Error when assigning name '" + dbMapperRangeName + "' to DBListFetch/DBSetQuery target range: " + ex.Message)
                End Try
                ' notify associated DBMapper object of new target range
                Try
                    Dim extendedMapper As DBMapper = Globals.DBModifDefColl("DBMapper").Item(dbMapperRangeName)
                    extendedMapper.setTargetRange(theRange)
                Catch ex As Exception
                    Throw New Exception("Error notifying the associated DBMapper object when extending '" + dbMapperRangeName + "' to DBListFetch/DBSetQuery target range: " + ex.Message)
                End Try
            End If
        End If
    End Sub

    ''' <summary>creates a DBModif at the current active cell or edits an existing one defined in targetDefName (after being called in defined range or from ribbon + Ctrl + Shift)</summary>
    Sub createDBModif(createdDBModifType As String, Optional targetDefName As String = "")
        ' clipboard helper for legacy definitions:
        ' if saveRangeToDB<Single> macro calls were copied into clipboard, 1st parameter (datarange) removed (empty), connid moved to 2nd place as database name (remove MSSQL)
        'mapper.saveRangeToDBSingle(Range("DB_DefName"), "tableName", "primKey1,primKey2,primKey3", "MSSQLDB_NAME", True) -> def(, "DB_NAME", "tableName", "primKey1,primKey2,primKey3", True)    DBMapperName = DB_DefName
        'mapper.saveRangeToDBSingle(Range("DB_DefName"), "tableName", "primKey1,primKey2,primKey3", "MSSQLDB_NAME")       -> def(, "DB_NAME", "tableName", "primKey1,primKey2,primKey3")          DBMapperName = DB_DefName
        '
        ' for saveRangeToDB(DataRange, tableNamesStr, primKeysStr, primKeyColumnsStr, startDataColumn, connid, ParamArray optionalArray())
        ' remove primKeyColumnsStr and startDataColumn before copying to clipboard...
        Dim existingDBModif As DBModif = Nothing
        Dim existingDefName As String = targetDefName
        Dim createdDBMapperFromClipboard As Boolean = False
        If Clipboard.ContainsText() And createdDBModifType = "DBMapper" Then
            Dim cpbdtext As String = Clipboard.GetText()
            If InStr(cpbdtext.ToLower(), "saverangetodb") > 0 Then
                Dim firstBracket As Integer = InStr(cpbdtext, "(")
                Dim firstComma As Integer = InStr(cpbdtext, ",")
                Dim connDefStart As Integer = InStrRev(cpbdtext, """" + fetchSetting("connIDPrefixDBtype", "MSSQL"))
                Dim commaBeforeConnDef As Integer = InStrRev(cpbdtext, ",", connDefStart)
                ' after conndef, all parameters are optional, so in case there is no comma afterwards, set this to end of whole definition string
                Dim commaAfterConnDef As Integer = IIf(InStr(connDefStart, cpbdtext, ",") > 0, InStr(connDefStart, cpbdtext, ","), Len(cpbdtext))
                Dim DB_DefName, newDefString, RangeDefName As String
                RangeDefName = Mid(cpbdtext, firstBracket + 1, firstComma - firstBracket - 1)
                Try : DB_DefName = "DBMapper" + Replace(Replace(Mid(RangeDefName, InStr(RangeDefName, "Range(""") + 7), """)", ""), ":", "")
                Catch ex As Exception
                    ErrorMsg("Error when retrieving DB_DefName from clipboard: " + ex.Message, "DBMapper Legacy Creation Error") : Exit Sub
                End Try
                Try : newDefString = "def(" + Replace(Mid(cpbdtext, commaBeforeConnDef, commaAfterConnDef - commaBeforeConnDef), "MSSQL", "") + Mid(cpbdtext, firstComma, commaBeforeConnDef - firstComma - 1) + Mid(cpbdtext, commaAfterConnDef - 1)
                Catch ex As Exception
                    ErrorMsg("Error when building new definition from clipboard: " + ex.Message, "DBMapper Legacy Creation Error") : Exit Sub
                End Try
                ' assign new name to active cell
                ' Add doesn't work directly with ExcelDnaUtil.Application.ActiveWorkbook.Names (late binding), so create an object here...
                Dim NamesList As Excel.Names = ExcelDnaUtil.Application.ActiveWorkbook.Names
                Try : NamesList.Add(Name:=DB_DefName, RefersTo:=ExcelDnaUtil.Application.ActiveCell)
                Catch ex As Exception
                    ErrorMsg("Error when assigning name '" + DB_DefName + "' to active cell: " + ex.Message, "DBMapper Legacy Creation Error")
                    Exit Sub
                End Try

                ' build a new DBMapper with legacy constructor
                existingDBModif = New DBMapper(defkey:=DB_DefName, paramDefs:=newDefString, paramTarget:=ExcelDnaUtil.Application.ActiveCell)
                existingDefName = DB_DefName
                createdDBMapperFromClipboard = True
                Clipboard.Clear()
            End If
        End If

        ' fetch parameters if there is an existing definition...
        If DBModifDefColl.ContainsKey(createdDBModifType) AndAlso DBModifDefColl(createdDBModifType).ContainsKey(existingDefName) Then
            existingDBModif = DBModifDefColl(createdDBModifType).Item(existingDefName)
            ' reset the target range to a potentially changed area
            If createdDBModifType <> "DBSeqnce" Then existingDBModif.setTargetRange(ExcelDnaUtil.Application.Range(existingDefName))
        End If

        ' prepare DBModifier Create Dialog
        Dim theDBModifCreateDlg As DBModifCreate = New DBModifCreate()
        With theDBModifCreateDlg
            ' store DBModification type in tag for validation purposes...
            .Tag = createdDBModifType
            .envSel.DataSource = Globals.environdefs
            .envSel.SelectedIndex = -1
            .DBModifName.Text = Replace(existingDefName, createdDBModifType, "")
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
                .IgnoreDataErrors.Hide()
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
                ' fill Datagridview for DBSequence
                Dim cb As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn With {
                    .HeaderText = "Sequence Step",
                    .ReadOnly = False
                }
                cb.ValueType() = GetType(String)
                Dim ds As List(Of String) = New List(Of String)

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
                For Each ws As Excel.Worksheet In ExcelDnaUtil.Application.ActiveWorkbook.Worksheets
                    For Each theFunc As String In {"DBListFetch(", "DBSetQuery(", "DBRowFetch("}
                        searchCell = ws.Cells.Find(What:=theFunc, After:=ws.Range("A1"), LookIn:=Excel.XlFindLookIn.xlFormulas, LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlNext, MatchCase:=False)
                        Dim firstFoundAddress As String = ""
                        If Not IsNothing(searchCell) Then firstFoundAddress = searchCell.Address
                        While Not IsNothing(searchCell)
                            Dim underlyingName As String = getDBunderlyingNameFromRange(searchCell)
                            ds.Add("Refresh " + theFunc + searchCell.Parent.Name + "!" + searchCell.Address + "):" + underlyingName)
                            searchCell = ws.Cells.FindNext(searchCell)
                            If searchCell.Address = firstFoundAddress Then Exit While
                        End While
                    Next
                    ' reset the cell find dialog....
                    searchCell = Nothing
                    searchCell = ws.Cells.Find(What:="", After:=ws.Range("A1"), LookIn:=Excel.XlFindLookIn.xlFormulas, LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlNext, MatchCase:=False)
                Next
                ' at last add special items DBBeginTrans and DBCommitTrans for setting DB Transaction brackets
                ds.Add("DBBegin:Begins DB Transaction")
                ds.Add("DBCommitRollback:Commits or Rolls back DB Transaction")
                ' and bind the dataset to the combobox
                cb.DataSource() = ds
                .DBSeqenceDataGrid.Columns.Add(cb)
                .DBSeqenceDataGrid.Columns(0).Width = 400
            Else
                theDBModifCreateDlg.FormBorderStyle = FormBorderStyle.FixedDialog
                If createdDBModifType = "DBAction" Then
                    theDBModifCreateDlg.MinimumSize = New Drawing.Size(width:=490, height:=160)
                    theDBModifCreateDlg.Size = New Drawing.Size(width:=490, height:=160)
                    .execOnSave.Top = .Tablename.Top
                    .AskForExecute.Top = .Tablename.Top
                    .envSel.Top = .Tablename.Top
                    '.TargetRangeAddress.Location = .PrimaryKeysLabel.Location
                Else
                    theDBModifCreateDlg.MinimumSize = New Drawing.Size(width:=490, height:=290)
                    theDBModifCreateDlg.Size = New Drawing.Size(width:=490, height:=290)
                End If
                ' hide controls irrelevant for DBMapper and DBAction
                .DBSeqenceDataGrid.Hide()
            End If

            ' delegate filling of dialog fields to created DBModif object
            If Not IsNothing(existingDBModif) Then existingDBModif.setDBModifCreateFields(theDBModifCreateDlg)

            ' display dialog for parameters
            If theDBModifCreateDlg.ShowDialog() = DialogResult.Cancel Then
                ' remove targetRange Name created in clipboard helper
                If createdDBMapperFromClipboard Then
                    Try : ExcelDnaUtil.Application.ActiveWorkbook.Names(existingDefName).Delete : Catch ex As Exception : End Try
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
                Try : ExcelDnaUtil.Application.ActiveWorkbook.Names(existingDefName).Delete : Catch ex As Exception : End Try
                ' then (re)set name
                Dim NamesList As Excel.Names = ExcelDnaUtil.Application.ActiveWorkbook.Names
                Dim alreadyExists As Boolean = False
                Try
                    Dim testExist As String = NamesList.Item(createdDBModifType + .DBModifName.Text).ToString
                Catch ex As Exception
                    alreadyExists = True
                End Try
                If Not alreadyExists Then
                    ErrorMsg("Error adding DBModifier '" + createdDBModifType + .DBModifName.Text + "', Name already exists in Workbook!", "DBModifier Creation Error")
                    Exit Sub
                End If
                Try
                    NamesList.Add(Name:=createdDBModifType + .DBModifName.Text, RefersTo:=targetRange)
                Catch ex As Exception
                    ErrorMsg("Error when assigning name '" + createdDBModifType + .DBModifName.Text + "' to active cell: " + ex.Message, "DBModifier Creation Error")
                    Exit Sub
                End Try
            End If

            Dim CustomXmlParts As Object = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
            If CustomXmlParts.Count = 0 Then ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.Add("<root xmlns=""DBModifDef""></root>")
            CustomXmlParts = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
            ' remove old node in case of renaming DBModifier
            ' Elements have names of DBModif types, attribute Name is given name (<DBMapper Name=existingDefName>)
            If Not IsNothing(CustomXmlParts(1).SelectSingleNode("/ns0:root/ns0:" + createdDBModifType + "[@Name='" + Replace(existingDefName, createdDBModifType, "") + "']")) Then CustomXmlParts(1).SelectSingleNode("/ns0:root/ns0:" + createdDBModifType + "[@Name='" + Replace(existingDefName, createdDBModifType, "") + "']").Delete
            ' NamespaceURI:="DBModifDef" is required to avoid adding a xmlns attribute to each element.
            CustomXmlParts(1).SelectSingleNode("/ns0:root").AppendChildNode(createdDBModifType, NamespaceURI:="DBModifDef")
            ' new appended elements are last, get it to append further child elements
            Dim dbModifNode As CustomXMLNode = CustomXmlParts(1).SelectSingleNode("/ns0:root").LastChild
            ' append the detailed settings to the definition element
            dbModifNode.AppendChildNode("Name", NodeType:=MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, NodeValue:= .DBModifName.Text)
            dbModifNode.AppendChildNode("execOnSave", NamespaceURI:="DBModifDef", NodeValue:= .execOnSave.Checked.ToString())
            dbModifNode.AppendChildNode("askBeforeExecute", NamespaceURI:="DBModifDef", NodeValue:= .AskForExecute.Checked.ToString())
            If createdDBModifType = "DBMapper" Then
                dbModifNode.AppendChildNode("env", NamespaceURI:="DBModifDef", NodeValue:=IIf(.envSel.SelectedIndex = -1, "", (.envSel.SelectedIndex + 1).ToString()))
                dbModifNode.AppendChildNode("database", NamespaceURI:="DBModifDef", NodeValue:= .Database.Text)
                dbModifNode.AppendChildNode("tableName", NamespaceURI:="DBModifDef", NodeValue:= .Tablename.Text)
                dbModifNode.AppendChildNode("primKeysStr", NamespaceURI:="DBModifDef", NodeValue:= .PrimaryKeys.Text)
                dbModifNode.AppendChildNode("insertIfMissing", NamespaceURI:="DBModifDef", NodeValue:= .insertIfMissing.Checked.ToString())
                dbModifNode.AppendChildNode("executeAdditionalProc", NamespaceURI:="DBModifDef", NodeValue:= .addStoredProc.Text)
                dbModifNode.AppendChildNode("ignoreColumns", NamespaceURI:="DBModifDef", NodeValue:= .IgnoreColumns.Text)
                dbModifNode.AppendChildNode("CUDFlags", NamespaceURI:="DBModifDef", NodeValue:= .CUDflags.Checked.ToString())
                dbModifNode.AppendChildNode("IgnoreDataErrors", NamespaceURI:="DBModifDef", NodeValue:= .IgnoreDataErrors.Checked.ToString())
            ElseIf createdDBModifType = "DBAction" Then
                dbModifNode.AppendChildNode("env", NamespaceURI:="DBModifDef", NodeValue:=IIf(.envSel.SelectedIndex = -1, "", (.envSel.SelectedIndex + 1).ToString()))
                dbModifNode.AppendChildNode("database", NamespaceURI:="DBModifDef", NodeValue:= .Database.Text)
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
            If Not IsNothing(existingDBModif) Then existingDBModif.addHiddenFeatureDefs(dbModifNode)
            ' refresh mapper definitions to reflect changes immediately...
            getDBModifDefinitions()
            ' extend Datarange for new DBMappers immediately after definition...
            If createdDBModifType = "DBMapper" Then
                DirectCast(Globals.DBModifDefColl("DBMapper").Item(createdDBModifType + .DBModifName.Text), DBMapper).extendDataRange()
            End If
        End With
    End Sub

    ''' <summary>gets defined names for DBModifier (DBMapper/DBAction/DBSeqnce) invocation in the current workbook and updates Ribbon with it</summary>
    Public Sub getDBModifDefinitions()
        ' load DBModifier definitions (objects) into Global collection DBModifDefColl
        Try
            Globals.DBModifDefColl.Clear()
            Dim CustomXmlParts As Object = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
            If CustomXmlParts.Count = 1 Then
                ' read definitions from CustomXMLParts
                ' get DBModifier definitions from docproperties
                For Each customXMLNodeDef As CustomXMLNode In CustomXmlParts(1).SelectSingleNode("/ns0:root").ChildNodes
                    Dim DBModiftype As String = Left(customXMLNodeDef.BaseName, 8)
                    If DBModiftype = "DBSeqnce" Or DBModiftype = "DBMapper" Or DBModiftype = "DBAction" Then
                        Dim nodeName As String
                        If customXMLNodeDef.Attributes.Count > 0 Then
                            nodeName = DBModiftype + customXMLNodeDef.Attributes(1).Text
                        Else
                            nodeName = customXMLNodeDef.BaseName
                        End If

                        Dim targetRange As Excel.Range = Nothing
                        ' for DBMappers and DBActions the data of the DBModification is stored in Ranges, so check for those and get the Range
                        If DBModiftype = "DBMapper" Or DBModiftype = "DBAction" Then
                            For Each rangename As Excel.Name In ExcelDnaUtil.Application.ActiveWorkbook.Names
                                Dim rangenameName As String = Replace(rangename.Name, rangename.Parent.Name + "!", "")
                                If rangenameName = nodeName And InStr(rangename.RefersTo, "#REF!") > 0 Then
                                    ErrorMsg(DBModiftype + " definitions range [" + rangename.Parent.Name + "]" + rangename.Name + " contains #REF!", "DBModifier Definitions Error")
                                    Exit For
                                ElseIf rangenameName = nodeName Then
                                    ' might fail...
                                    Try : targetRange = rangename.RefersToRange : Catch ex As Exception : End Try
                                    Exit For
                                End If
                            Next
                            If IsNothing(targetRange) Then
                                ErrorMsg("Error, required target range named '" + nodeName + "' not existing for " + DBModiftype + "." + vbCrLf + "either create target range or delete CustomXML definition node named '" + nodeName + "' (Ctrl-Shift-Click on Store DBModif Data dialogBox launcher)!", "DBModifier Definitions Error")
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
                            ErrorMsg("Error, not supported DBModiftype: " + DBModiftype, "DBModifier Definitions Error")
                        End If
                        ' ... and add it to the collection DBModifDefColl
                        Dim defColl As Dictionary(Of String, DBModif) ' definition lookup collection for DBModifiername -> object
                        If Not IsNothing(newDBModif) Then
                            If Not DBModifDefColl.ContainsKey(DBModiftype) Then
                                ' add to new DBModiftype "menu"
                                defColl = New Dictionary(Of String, DBModif) From {
                                    {nodeName, newDBModif}
                                }
                                DBModifDefColl.Add(DBModiftype, defColl)
                            Else
                                ' add definition to existing DBModiftype "menu"
                                defColl = DBModifDefColl(DBModiftype)
                                defColl.Add(nodeName, newDBModif)
                            End If
                        End If
                    End If
                Next
            ElseIf CustomXmlParts.Count > 1 Then
                ErrorMsg("Multiple CustomXmlParts for DBModifDef existing!", "get DBModif Definitions")
            End If
            Globals.theRibbon.Invalidate()
        Catch ex As Exception
            ErrorMsg("Exception: " + ex.Message, "get DBModif Definitions")
        End Try
    End Sub

    ''' <summary>gets DB Modification Name (DBMapper or DBAction) from theRange</summary>
    ''' <param name="theRange"></param>
    ''' <returns>the retrieved name as a string (not name object !)</returns>
    Public Function getDBModifNameFromRange(theRange As Excel.Range) As String
        Dim nm As Excel.Name
        Dim rng, testRng As Excel.Range

        getDBModifNameFromRange = ""
        If IsNothing(theRange) Then Exit Function
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
            ErrorMsg("Exception: " + ex.Message, "get DBModif Name From Range")
        End Try
    End Function

    ''' <summary>To check for errors in passed range obj, makes use of the fact that Range.Value never passes Integer Values back except for Errors</summary>
    ''' <param name="rangeval">Range.Value to be checked for errors</param>
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
        hadError = False : nonInteractiveErrMsgs = "" : nonInteractive = headLess
        Dim DBModiftype As String = Left(DBModifName, 8)
        If DBModiftype = "DBSeqnce" Or DBModiftype = "DBMapper" Or DBModiftype = "DBAction" Then
            If Not Globals.DBModifDefColl(DBModiftype).ContainsKey(DBModifName) Then
                Dim DBModifavailable As String = ""
                For Each DBMtype As String In {"DBMapper", "DBAction", "DBSeqnce"}
                    For Each DBMkey As String In Globals.DBModifDefColl(DBMtype).Keys : DBModifavailable += "," + DBMkey : Next
                Next
                nonInteractive = False
                Return "DB Modifier '" + DBModifName + "' not existing, available: " + DBModifavailable
            End If
            Try
                Globals.DBModifDefColl(DBModiftype).Item(DBModifName).doDBModif()
            Catch ex As Exception
                nonInteractive = False
                Return "DB Modifier '" + DBModifName + "' doDBModif had following error(s): " + ex.Message
            End Try
            If hadError Then
                nonInteractive = False
                Return nonInteractiveErrMsgs
            Else
                nonInteractive = False
                Return ""
            End If
        Else
            nonInteractive = False
            Return "No valid type (" + DBModiftype + ") in passed DB Modifier '" + DBModifName + "', DB Modifier name must start with 'DBSeqnce', 'DBMapper' Or 'DBAction' !"
        End If
    End Function

    ''' <summary>marks a row in a DBMapper for deletion</summary>
    <ExcelCommand(Name:="deleteRow", ShortCut:="^D")>
    Public Sub deleteRow()
        Dim targetName As String = getDBModifNameFromRange(ExcelDnaUtil.Application.Selection)
        If Left(targetName, 8) = "DBMapper" Then DirectCast(Globals.DBModifDefColl("DBMapper").Item(targetName), DBMapper).doCUDMarks(ExcelDnaUtil.Application.Selection, True)
    End Sub

    ''' <summary>inserts a row in a DBMapper</summary>
    <ExcelCommand(Name:="insertRow", ShortCut:="^I")>
    Public Sub insertRow()
        Dim targetName As String = getDBModifNameFromRange(ExcelDnaUtil.Application.Selection)
        If Left(targetName, 8) = "DBMapper" Then
            ' get the target range for the DBMapper to get the ListObject
            Dim insertTarget As Excel.Range = DirectCast(Globals.DBModifDefColl("DBMapper").Item(targetName), DBMapper).getTargetRange
            ' calculate insert row from selection and top row of insert target
            Dim insertRow As Integer = ExcelDnaUtil.Application.Selection.Row - insertTarget.Row
            insertTarget.ListObject.ListRows.Add(insertRow)
        End If
    End Sub
End Module

