Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports Microsoft.Office.Core
Imports System.Collections
Imports System.Data
Imports System.Text
Imports System.Linq

''' <summary>Abstraction of a DB Modification Object (concrete classes DBMapper, DBAction or DBSeqnce)</summary>
Public MustInherit Class DBModif

    '''<summary>unique key of DBModif</summary>
    Protected dbmodifName As String
    ''' <summary>Range where DBMapper data is located (only DBMapper and DBAction; paramText is stored in custom doc properties having the same Name)</summary>
    Protected TargetRange As Excel.Range
    ''' <summary>DBModif name of target range</summary>
    Protected paramTargetName As String
    ''' <summary>Database to store to, not available to DB Sequences</summary>
    Protected database As String
    '''<summary>should DBMap be saved / DBAction be done on Excel Saving? (default no)</summary>
    Public execOnSave As Boolean = False
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
        execOnSave = Convert.ToBoolean(getParamFromXML(definitionXML, "execOnSave", "Boolean"))
        askBeforeExecute = Convert.ToBoolean(getParamFromXML(definitionXML, "askBeforeExecute", "Boolean"))
        confirmText = getParamFromXML(definitionXML, "confirmText")
    End Sub

    ''' <summary>needed for legacy DBmapper constructor</summary>
    Public Sub New()

    End Sub

    ''' <summary>asks user the confirmation question, in case it is required by the DB Modifier</summary>
    ''' <returns>Yes, No or Cancel (only possible when saving to finish questions)</returns>
    Public Function confirmExecution(Optional WbIsSaving As Boolean = False) As MsgBoxResult
        ' when saving always ask for DBModifiers that are to be executed, otherwise ask only if required....
        If (WbIsSaving And execOnSave) Or (Not WbIsSaving And askBeforeExecute) Then
            Dim executeTitle As String = "Execute " + TypeName(Me) + IIf(WbIsSaving, " on save", "")
            ' special confirm text or standard text?
            Dim executeMessage As String = IIf(confirmText <> "", confirmText, "Really execute " + dbmodifName + "?")
            ' for confirmation on Workbook saving provide cancel possibility as shortcut.
            Dim setQuestionType As MsgBoxStyle = IIf(WbIsSaving, MsgBoxStyle.YesNoCancel, MsgBoxStyle.YesNo)
            ' only ask if set to ask...
            If askBeforeExecute Then Return QuestionMsg(theMessage:=executeMessage, questionType:=setQuestionType, questionTitle:=executeTitle)
        End If
        Return MsgBoxResult.No
    End Function

    ''' <summary>public accessor function</summary>
    ''' <returns>the targetRangeAddress nicely formatted</returns>
    Public Function getTargetRangeAddress() As String
        getTargetRangeAddress = ""
        If TypeName(Me) <> "DBSeqnce" Then
            Dim addRefersToFormula As String = ""
            If InStr(1, ExcelDnaUtil.Application.ActiveWorkbook.Names.Item(dbmodifName).RefersTo, "=OFFSET(") > 0 Then
                addRefersToFormula = " (" + ExcelDnaUtil.Application.ActiveWorkbook.Names.Item(dbmodifName).RefersTo
            End If
            getTargetRangeAddress = TargetRange.Parent.Name + "!" + TargetRange.Address + addRefersToFormula
        End If
    End Function

    ''' <summary>public accessor function</summary>
    ''' <returns>the targetRange itself</returns>
    Public Function getTargetRange() As Excel.Range
        getTargetRange = Nothing
        If TypeName(Me) <> "DBSeqnce" Then
            getTargetRange = TargetRange
        End If
    End Function

    ''' <summary>public accessor function: get Environment (integer) where connection id should be taken from (if not existing, take from selectedEnvironment being passed in defaultEnv)</summary>
    ''' <param name="defaultEnv">optionally passed selected Environment</param>
    ''' <returns>the Environment of the DBModif</returns>
    Protected Function getEnv(Optional defaultEnv As Integer = 0) As Integer
        getEnv = defaultEnv
        If TypeName(Me) = "DBSeqnce" Then Throw New NotImplementedException()
        ' set environment on DBModif overrides selected environment. Could be empty or 0 to indicate a not set environment...
        If env <> "" AndAlso env <> "0" Then getEnv = Convert.ToInt16(env)
    End Function

    ''' <summary>checks whether DBModifier needs saving, usually because execOnSave is set (in case of CUD DBMappers if any i/u/d Flages are present)</summary>
    ''' <returns>true if save needed</returns>
    Public Overridable Function DBModifSaveNeeded() As Boolean
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
    Protected Function getParamFromXML(definitionXML As CustomXMLNode, nodeName As String, Optional ReturnType As String = "") As String
        Dim nodeCount As Integer = definitionXML.SelectNodes("ns0:" + nodeName).Count
        If nodeCount = 0 Then
            getParamFromXML = "" ' optional nodes become empty strings
        Else
            getParamFromXML = definitionXML.SelectSingleNode("ns0:" + nodeName).Text
        End If
        If ReturnType = "Boolean" And getParamFromXML = "" Then getParamFromXML = "False"
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

    ''' <summary>open a database specific connection, not available to DB Sequences</summary>
    ''' <returns></returns>
    Public Function openDatabase(Optional DBSequenceEnv As String = "") As Boolean
        If TypeName(Me) = "DBSeqnce" Then Throw New NotImplementedException() ' DB Sequences have no database
        Dim setEnv As Integer = getEnv()
        If DBSequenceEnv = "" Then
            ' if Environment is not existing (default environment = 0), take from selectedEnvironment
            If setEnv = 0 Then setEnv = Globals.selectedEnvironment + 1
        Else
            ' if Environment is not existing (default environment = 0), take from sequence environment
            If setEnv = 0 Then
                setEnv = CInt(DBSequenceEnv)
            Else
                ' otherwise check if set environment is different from sequence environment and set to sequence environment after warning message.
                If setEnv <> CInt(DBSequenceEnv) Then
                    Globals.UserMsg("Environment in " + TypeName(Me) + " (" + env + ") different than given environment of DB Sequence (" + DBSequenceEnv + "), overruling with sequence environment !")
                    setEnv = CInt(DBSequenceEnv)
                End If
            End If
        End If
        openDatabase = openIdbConnection(setEnv, database)
    End Function

    ''' <summary>refresh a DB Function (currently only DBListFetch and DBSetQuery) by invoking its respective DB*Action procedure
    ''' It is necessary to prepare the inputs to this DB*Action procedure as the UDF cannot be invoked from here</summary>
    ''' <param name="srcExtent">the unique hidden name of the DB Function cell (DBFsource(GUID))</param>
    ''' <param name="executedDBMappers">in a DB Sequence, this parameter notifies of DBMappers that were executed before to allow avoidance of refreshing changes</param>
    ''' <param name="modifiedDBMappers">in a DB Sequence, this parameter notifies of a DBMapper that had changes, necessary to avoid deadlocks</param>
    ''' <param name="TransactionIsOpen">in a DB Sequence, this parameter notifies of an open transaction, necessary to avoid deadlocks</param>
    ''' <returns></returns>
    Protected Function doDBRefresh(srcExtent As String, Optional executedDBMappers As Dictionary(Of String, Boolean) = Nothing, Optional modifiedDBMappers As Dictionary(Of String, Boolean) = Nothing, Optional TransactionIsOpen As Boolean = False) As Boolean
        If executedDBMappers Is Nothing Then executedDBMappers = New Dictionary(Of String, Boolean)
        If modifiedDBMappers Is Nothing Then modifiedDBMappers = New Dictionary(Of String, Boolean)
        doDBRefresh = False
        ' refresh DBFunction in sequence, invoke this "manually", simulating the call of the user defined function by excel
        Dim caller As Excel.Range
        Try : caller = ExcelDnaUtil.Application.Range(srcExtent)
        Catch ex As Exception
            Globals.UserMsg("Didn't find caller cell of DBRefresh using srcExtent " + srcExtent + "!", "Refresh of DB Functions")
            Exit Function
        End Try
        If caller.Parent.ProtectContents Then
            Globals.UserMsg("Worksheet " + caller.Parent.Name + " is content protected, can't refresh DB Function !")
            Exit Function
        End If
        Dim targetExtent = Replace(srcExtent, "DBFsource", "DBFtarget")
        Dim target As Excel.Range
        Try : target = ExcelDnaUtil.Application.Range(targetExtent)
        Catch ex As Exception
            Globals.UserMsg("Didn't find target of DBRefresh using targetExtent " + targetExtent + "!", "Refresh of DB Functions")
            Exit Function
        End Try
        If target.Parent.ProtectContents Then
            Globals.UserMsg("Worksheet " + target.Parent.Name + " is content protected, can't refresh DB Function !")
            Exit Function
        End If
        Dim DBMapperUnderlying As String = getDBModifNameFromRange(target)
        Dim targetExtentF = Replace(srcExtent, "DBFsource", "DBFtargetF")
        Dim formulaRange As Excel.Range = Nothing
        ' formulaRange might not exist
        Try : formulaRange = ExcelDnaUtil.Application.Range(targetExtentF) : Catch ex As Exception : End Try
        If formulaRange IsNot Nothing AndAlso formulaRange.Parent.ProtectContents Then
            Globals.UserMsg("Worksheet " + formulaRange.Parent.Name + " is content protected, can't refresh DB Function !")
            Exit Function
        End If
        Dim DBMapperUnderlyingF As String = getDBModifNameFromRange(formulaRange)
        ' allow for avoidance of overwriting users changes with CUDFlags after a data error occurred
        If hadError Then
            If executedDBMappers.ContainsKey(DBMapperUnderlying) Then
                Dim retval = QuestionMsg(theMessage:="Error(s) occured during sequence, really refresh Targetrange? This could lead to loss of entries.", questionTitle:="Refresh of DB Functions in DB Sequence")
                If retval = vbCancel Then Exit Function
            End If
        End If
        ' prevent deadlock if we are in a transaction and want to refresh the area that was changed.
        If (modifiedDBMappers.ContainsKey(DBMapperUnderlying) Or modifiedDBMappers.ContainsKey(DBMapperUnderlyingF)) And TransactionIsOpen Then
            Globals.UserMsg("Transaction affecting the target area is still open, refreshing it would result in a deadlock on the database table. Skipping refresh, consider placing refresh outside of transaction.", "Refresh of DB Functions in DB Sequence")
            Exit Function
        End If
        ' reset query cache, so we really get new data !
        Dim callID As String
        Try
            ' get the callID of the underlying name of the target (key of the queryCache and StatusCollection)
            callID = "[" + caller.Parent.Parent.Name + "]" + caller.Parent.Name + "!" + caller.Address
        Catch ex As Exception
            Globals.UserMsg("Didn't find target of DBRefresh !", "Refresh of DB Functions")
            Exit Function
        End Try
        Try
            ' StatusCollection doesn't necessarily have the callID contained
            If Not StatusCollection.ContainsKey(callID) Then StatusCollection.Add(callID, New ContainedStatusMsg)
            Dim functionFormula As String = ExcelDnaUtil.Application.Range(srcExtent).Formula
            If UCase(Left(functionFormula, 12)) = "=DBLISTFETCH" Then
                Globals.LogInfo("Refresh DBListFetch: " + callID)
                Dim functionArgs = Globals.functionSplit(functionFormula, ",", """", "DBListFetch", "(", ")")
                If functionArgs.Length() < 3 Then
                    functionArgs = Globals.functionSplit(functionFormula, listSepLocal, """", "DBListFetch", "(", ")")
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
                Globals.LogInfo("Refresh DBSetQuery: " + callID)
                Dim functionArgs = Globals.functionSplit(functionFormula, ",", """", "DBSetQuery", "(", ")")
                If functionArgs.Length() < 3 Then
                    functionArgs = Globals.functionSplit(functionFormula, listSepLocal, """", "DBSetQuery", "(", ")")
                End If
                Dim targetRangeName As String : targetRangeName = functionArgs(2)
                If UBound(functionArgs) = 3 Then targetRangeName += "," + functionArgs(3)
                Functions.DBSetQueryAction(callID, getQuery(functionArgs(0), caller), target, getConnString(functionArgs(1), caller, True), caller, targetRangeName)
            ElseIf UCase(Left(functionFormula, 11)) = "=DBROWFETCH" Then
                Globals.LogInfo("Refresh DBRowFetch: " + callID)
                Dim functionArgs = Globals.functionSplit(functionFormula, ",", """", "DBRowFetch", "(", ")")
                If functionArgs.Length() < 3 Then
                    functionArgs = Globals.functionSplit(functionFormula, listSepLocal, """", "DBRowFetch", "(", ")")
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
            Globals.UserMsg("Exception: " + ex.Message, "DBRefresh")
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
            ' funcArg is already a reference to a parent sheet
            Query = ExcelDnaUtil.Application.Evaluate(funcArg)
        Else
            ' avoid adding parent WS name if argument is a string (only plain range references need adding)
            If Left(funcArg, 1) = """" Then
                Query = ExcelDnaUtil.Application.Evaluate(funcArg)
            Else
                ' add parent name, otherwise evaluate will fail
                Query = ExcelDnaUtil.Application.Evaluate("'" + caller.Parent.Name + "'!" + funcArg)
            End If
        End If
        ' either funcArg is already a string (direct contained/chained in the function) or it is a reference to a range. 
        ' In the latter case derive actual Query from range...
        If TypeName(Query) = "Range" Then Query = Query.Value.ToString()
        getQuery = Query
    End Function

    ''' <summary>get connection string from passed function argument</summary>
    ''' <param name="funcArg">function argument parsed from DBFunction formula, can be empty, a number (as a string) or a String</param>
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
    '''<summary>first columnn is treated as an autoincrementing key column, needed to ignore empty values there (otherwise DBMappers stop here)</summary>
    Private AutoIncFlag As Boolean = False
    ''' <summary>prevent filling of whole table during execution of DB Mappers, this is useful for very large tables that are incrementally filled and would take unnecessary long time to start the DB Mapper. If set to true then each record is searched independently by going to the database. If the records to be stored are not too many, then this is more efficient than loading a very large table.</summary>
    Private avoidFill As Boolean = False

    ''' <summary>legacy constructor for mapping existing DBMapper macro calls (copy in clipboard)</summary>
    ''' <param name="defkey"></param>
    ''' <param name="paramDefs"></param>
    ''' <param name="paramTarget"></param>
    Public Sub New(defkey As String, paramDefs As String, paramTarget As Excel.Range)
        dbmodifName = defkey
        ' if no target range is set, then no parameters can be found...
        If paramTarget Is Nothing Then
            Throw New Exception("paramTarget is Nothing")
        End If
        paramTargetName = getDBModifNameFromRange(paramTarget)
        If Left(paramTargetName, 8) <> "DBMapper" Then
            Throw New Exception("target " + paramTargetName + " not matching passed DBModif type DBMapper for " + getTargetRangeAddress() + "/" + dbmodifName + "!")
        End If
        Dim paramText As String = paramDefs
        TargetRange = paramTarget

        Dim DBModifParams() As String = Globals.functionSplit(paramText, ",", """", "def", "(", ")")
        If DBModifParams Is Nothing Then Exit Sub
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
            If paramTarget Is Nothing Then Throw New Exception("paramTarget is Nothing")
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
                Throw New Exception("couldn't get primary key count given in DBMapper definition:" + ex.Message)
            End Try
            insertIfMissing = Convert.ToBoolean(getParamFromXML(definitionXML, "insertIfMissing", "Boolean"))
            executeAdditionalProc = getParamFromXML(definitionXML, "executeAdditionalProc")
            ignoreColumns = getParamFromXML(definitionXML, "ignoreColumns")
            IgnoreDataErrors = Convert.ToBoolean(getParamFromXML(definitionXML, "IgnoreDataErrors", "Boolean"))
            CUDFlags = Convert.ToBoolean(getParamFromXML(definitionXML, "CUDFlags", "Boolean"))
            AutoIncFlag = Convert.ToBoolean(getParamFromXML(definitionXML, "AutoIncFlag", "Boolean"))
            avoidFill = Convert.ToBoolean(getParamFromXML(definitionXML, "avoidFill", "Boolean"))
            If TargetRange.ListObject IsNot Nothing Then
                ' special grey table style for CUDFlags DBMapper
                If CUDFlags Then
                    TargetRange.ListObject.TableStyle = fetchSetting("DBMapperCUDFlagStyle", "TableStyleLight11")
                    ' otherwise blue
                Else
                    TargetRange.ListObject.TableStyle = fetchSetting("DBMapperStandardStyle", "TableStyleLight9")
                End If
            End If
            ' allow CUDFlags only on DBMappers with underlying Listobjects that were created with a query
            If CUDFlags And (TargetRange.ListObject Is Nothing OrElse TargetRange.ListObject.SourceType <> Excel.XlListObjectSourceType.xlSrcQuery) Then
                CUDFlags = False
                definitionXML.SelectSingleNode("ns0:CUDFlags").Delete()
                definitionXML.AppendChildNode("CUDFlags", NamespaceURI:="DBModifDef", NodeValue:="False")
                Throw New Exception("CUDFlags only supported for DBMappers on ListObjects (created with DBSetQueryListObject)!")
            End If
        Catch ex As Exception
            Globals.UserMsg("Error when creating DBMapper '" + dbmodifName + "': " + ex.Message, "DBModifier Definitions Error")
        End Try
    End Sub

    ''' <summary>checks whether DBModifier needs saving, either because execOnSave is set or in case of CUD DBMappers if any i/u/d Flages are present</summary>
    ''' <returns>true if save needed</returns>
    Public Overrides Function DBModifSaveNeeded() As Boolean
        If CUDFlags Then
            ' use TargetRange.ListObject.Range instead of TargetRange here because TargetRange is updated only after Saving/executing the DBMapper...
            Dim testRange As Excel.Range = TargetRange.ListObject.Range.Columns(TargetRange.ListObject.Range.Columns.Count).Offset(0, 1)
            ' check whether i/u/d column (one to the right of DBMapper range) is empty..
            Dim existingCUDFlags As Integer = ExcelDnaUtil.Application.WorksheetFunction.CountIfs(testRange, "<>")
            Return existingCUDFlags > 0
        Else
            Return MyBase.DBModifSaveNeeded()
        End If
    End Function

    Public Function hadChanges() As Boolean
        Return changesDone
    End Function

    ''' <summary>inserts CUD (Create/Update/Delete) Marks at the right end of the DBMapper range</summary>
    ''' <param name="changedRange">passed TargetRange by Change Event or delete button</param>
    ''' <param name="deleteFlag">if delete button was pressed, this is true</param>
    Public Sub insertCUDMarks(changedRange As Excel.Range, Optional deleteFlag As Boolean = False)
        If Not CUDFlags Then Exit Sub
        ' sanity check for single cell DB Mappers..
        If TargetRange.Columns.Count = 1 And TargetRange.Rows.Count = 1 Then
            Dim retval As MsgBoxResult = QuestionMsg(theMessage:="DB Mapper Range with CUD Flags is only one cell, really set CUD Flags ?", questionTitle:="Set CUD Flags for DB Mapper")
            If retval = vbCancel Then Exit Sub
        End If
        ' sanity check for whole range change (this happens when the table is autofilled down by dragging while being INSIDE the table)..
        ' in this case excel extends the change to the whole table and additionally the dragged area...
        If TargetRange.Columns.Count = changedRange.Columns.Count And TargetRange.Rows.Count <= changedRange.Rows.Count Then
            Dim retval As MsgBoxResult = QuestionMsg(theMessage:="Change affects whole DB Mapper Range, this might lead to erroneous behaviour, really set CUD Flags ?", questionTitle:="Set CUD Flags for DB Mapper")
            If retval = vbCancel Then Exit Sub
        End If
        Dim changedRangeRows As Integer = changedRange.Rows.Count
        If changedRangeRows > CInt(fetchSetting("maxRowCountCUD", "10000")) Then
            If Not QuestionMsg(theMessage:="A large range was changed (" + changedRange.Rows.Count.ToString() + " > maxRowCountCUD:" + fetchSetting("maxRowCountCUD", "10000") + "), this will probably lead to CUD flag setting taking very long. Continue?", questionTitle:="Set CUD Flags for DB Mapper") Then Exit Sub
        End If
        If changedRange.Parent.ProtectContents Then
            Globals.UserMsg("Worksheet " + changedRange.Parent.Name + " is content protected, can't set CUD Flags !")
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
            Dim countRow As Integer = 1
            For Each changedRow As Excel.Range In changedRange.Rows
                Dim CUDMarkRow As Integer = changedRow.Row - TargetRange.Row + 1
                ' change only if not already set or
                If TargetRange.Cells(CUDMarkRow, TargetRange.Columns.Count + 1).Value Is Nothing Then
                    Dim RowContainsData As Boolean = False
                    For Each containedCell As Excel.Range In TargetRange.Rows(CUDMarkRow).Cells
                        ' check without newly inserted/updated cells (copy paste) 
                        Dim possibleIntersection As Excel.Range = ExcelDnaUtil.Application.Intersect(containedCell, changedRange)
                        ' check if whole row is empty (except for the changedRange), formulas do not count as filled (automatically filled for lookups or other things)..
                        If containedCell.Value IsNot Nothing AndAlso possibleIntersection Is Nothing AndAlso Left(containedCell.Formula, 1) <> "=" Then
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
                    ExcelDnaUtil.Application.Statusbar = "Create/Update/Delete mark for row " + countRow.ToString() + "/" + changedRange.Rows.Count.ToString()
                    countRow += 1
                End If
            Next
        End If
        ExcelDnaUtil.Application.Statusbar = False
        ExcelDnaUtil.Application.AutoCorrect.AutoExpandListRange = True
    End Sub

    ''' <summary>extend DataRange to "whole" DBMApper area (first row (header/field names) to the right and first column (first primary key) down)</summary>
    Public Sub extendDataRange()
        Dim NamesList As Excel.Names = ExcelDnaUtil.Application.ActiveWorkbook.Names
        ' only extend like this if no CUD Flags or AutoIncFlag present (may have non existing first (primary) columns -> auto identity columns !)
        If Not (CUDFlags Or AutoIncFlag) Then
            If TargetRange.Cells(2, 1).Value Is Nothing Then Exit Sub ' only extend if there are multiple rows...
            preventChangeWhileFetching = True
            Dim rowCount As Integer = TargetRange.Cells(1, 1).End(Excel.XlDirection.xlDown).Row - TargetRange.Cells(1, 1).Row + 1
            ' unfortunately the above method to find the column extent doesn't work with hidden columns, so count the filled cells directly...
            Dim colCount As Integer = 1
            While Not (TargetRange.Cells(1, colCount + 1).Value Is Nothing OrElse TargetRange.Cells(1, colCount + 1).Value.ToString() = "")
                colCount += 1
            End While
            Try
                ' only if the referral is to a real range (not an offset formula !)
                If InStr(1, NamesList.Item(paramTargetName).RefersTo, "=OFFSET(") = 0 Then
                    NamesList.Item(paramTargetName).RefersTo = NamesList.Item(paramTargetName).RefersToRange.Resize(rowCount, colCount)
                    ' this has lead to a strange replacing of the rightdownmost cell's formula by the cell's value
                    'NamesList.Add(Name:=paramTargetName, RefersTo:=TargetRange.Parent.Range(TargetRange.Cells(1, 1), TargetRange.Cells(rowEnd, colEnd)))
                End If
            Catch ex As Exception
                Throw New Exception("Error when resizing name '" + paramTargetName + "' to DBMapper while extending DataRange: " + ex.Message)
            Finally
                preventChangeWhileFetching = False
            End Try
        End If
        ' reassign extended area to TargetRange
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
                Globals.UserMsg("Worksheet " + TargetRange.Parent.Name + " is content protected, can't reset CUD Flags !")
                Exit Sub
            End If
            ExcelDnaUtil.Application.AutoCorrect.AutoExpandListRange = False ' to prevent automatic creation of new column
            TargetRange.Columns(TargetRange.Columns.Count + 1).ClearContents
            ' remove updated rows cell style
            TargetRange.Font.Italic = False
            ' remove deleted rows cell style
            TargetRange.Font.Strikethrough = False
            ' remove automatically created hyperlink formatting (link is removed by refresh, but formats stay)
            TargetRange.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic
            TargetRange.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone
            ExcelDnaUtil.Application.AutoCorrect.AutoExpandListRange = True ' to prevent automatic creation of new column
        End If
    End Sub

    ''' <summary>to re-add hidden features only available in XML</summary>
    ''' <param name="definitionXML">the definition node of the DB Modifier where the hidden features should be added</param>
    Public Overrides Sub addHiddenFeatureDefs(definitionXML As CustomXMLNode)
        MyBase.addHiddenFeatureDefs(definitionXML)
        definitionXML.AppendChildNode("avoidFill", NamespaceURI:="DBModifDef", NodeValue:=avoidFill.ToString())
    End Sub

    ''' <summary>execute the modifications for the DB Mapper by storing the data modifications in the DBMapper range to the database</summary>
    ''' <param name="WbIsSaving">flag for being called during Workbook saving</param>
    ''' <param name="calledByDBSeq">the name of the DB Sequence that called the DBMapper</param>
    ''' <param name="TransactionOpen">flag whether a transaction is open during the DB Sequence</param>
    Public Overrides Sub doDBModif(Optional WbIsSaving As Boolean = False, Optional calledByDBSeq As String = "", Optional TransactionOpen As Boolean = False)
        changesDone = False
        ' ask for saving only if a) is not done on WorkbookSave b) is set to ask and c) is not called by a DBSequence (asks already for saving) and d) is in interactive mode
        If Not WbIsSaving And askBeforeExecute And calledByDBSeq = "" And Not nonInteractive Then
            If Not confirmExecution() = MsgBoxResult.Yes Then Exit Sub
        End If
        extendDataRange()
        ' check for mass changes and warn if necessary
        If CUDFlags Then
            Dim maxMassChanges As Integer = CInt(fetchSetting("maxNumberMassChange", "30"))
            Dim curWs As Excel.Worksheet = TargetRange.Parent ' this is necessary because using TargetRange directly deletes the content of the CUD area !!
            Dim changesToBeDone As Integer = ExcelDnaUtil.Application.WorksheetFunction.CountIf(curWs.Range(TargetRange.Columns(TargetRange.Columns.Count + 1).Address), "<>")
            If changesToBeDone > maxMassChanges Then
                Dim retval As MsgBoxResult = QuestionMsg(theMessage:="Modifying more rows (" + changesToBeDone.ToString() + ") than defined warning limit (" + maxMassChanges.ToString() + "), continue?", questionTitle:="Execute DB Mapper")
                If retval = vbCancel Then Exit Sub
            End If
        End If
        'now create/get a connection (dbcnn) for env(ironment) in case it was not already created by a step in the sequence before (transactions!)
        If Not TransactionOpen Then
            ExcelDnaUtil.Application.StatusBar = "opening database connection for " + database
            If Not openDatabase() Then Exit Sub
        End If

        ' set up data adapter and data set for checking DBMapper columns
        Dim da As Common.DbDataAdapter = Nothing
        Dim dscheck As DataSet = New DataSet()
        ExcelDnaUtil.Application.StatusBar = "initialising the Data Adapter"
        Try
            If TypeName(idbcnn) = "SqlConnection" Then
                ' decent behaviour for SQL Server
                Using comm As SqlClient.SqlCommand = New SqlClient.SqlCommand("SET ARITHABORT ON", idbcnn, DBModifs.trans)
                    comm.ExecuteNonQuery()
                End Using
                da = New SqlClient.SqlDataAdapter(New SqlClient.SqlCommand("select * from " + tableName, idbcnn, DBModifs.trans))
                da.UpdateBatchSize = 20
            ElseIf TypeName(idbcnn) = "OleDbConnection" Then
                da = New OleDb.OleDbDataAdapter(New OleDb.OleDbCommand("select * from " + tableName, idbcnn, DBModifs.trans))
                da.UpdateBatchSize = 20
            Else
                da = New Odbc.OdbcDataAdapter(New Odbc.OdbcCommand("select * from " + tableName, idbcnn, DBModifs.trans))
            End If
            da.SelectCommand.CommandTimeout = Globals.CmdTimeout
        Catch ex As Exception
            Globals.UserMsg("Error in initializing Data Adapter for " + tableName + ": " + ex.Message, "DBMapper Error")
        End Try

        ExcelDnaUtil.Application.StatusBar = "retrieving the schema for " + tableName
        Try
            da.FillSchema(dscheck, SchemaType.Source, tableName)
        Catch ex As Exception
            Globals.UserMsg("Error in retrieving Schema for " + tableName + ": " + ex.Message, "DBMapper Error")
        End Try

        ' get the DataTypeName from the database if it exists, so we have a more accurate type information for the parameterized commands (select/insert/update/delete)
        Dim schemaReader As Common.DbDataReader = Nothing
        Dim schemaDataTypeCollection As Collection = New Collection
        Try
            schemaReader = da.SelectCommand.ExecuteReader()
            For Each schemaRow As DataRow In schemaReader.GetSchemaTable().Rows
                Try : schemaDataTypeCollection.Add(schemaRow("DataTypeName"), schemaRow("ColumnName")) : Catch ex As Exception : End Try
            Next
            ' cancel command to finish datareader (otherwise close takes very long until timeout)
            da.SelectCommand.Cancel()
            schemaReader.Close()
        Catch ex As Exception
            If Not IsNothing(schemaReader) Then
                da.SelectCommand.Cancel()
                schemaReader.Close()
            End If
        End Try

        ' first check if all column names (except ignored) of DBMapper Range exist in table and collect fieldnames
        Dim allColumnsStr(TargetRange.Columns.Count - 1) As String
        Dim colNum As Integer = 1
        Dim fieldColNum As Integer = 0
        Do
            Dim fieldname As String = Trim(TargetRange.Cells(1, colNum).Value)
            ' only if not ignored...
            If InStr(1, LCase(ignoreColumns) + ",", LCase(fieldname) + ",") = 0 Then
                If Not dscheck.Tables(0).Columns.Contains(fieldname) Then
                    hadError = True
                    TargetRange.Parent.Activate
                    TargetRange.Cells(1, colNum).Select
                    Globals.UserMsg("Field '" + fieldname + "' does not exist in Table '" + tableName + "' and is not in ignoreColumns, Error in sheet " + TargetRange.Parent.Name, "DBMapper Error")
                    GoTo cleanup
                Else
                    allColumnsStr(fieldColNum) = fieldname
                    fieldColNum += 1
                End If
            End If
            colNum += 1
        Loop Until colNum > TargetRange.Columns.Count
        ' keep only those that were filled...
        ReDim Preserve allColumnsStr(fieldColNum - 1)

        ' before setting the commands for the adapter, we need to have the primary key information, or update/delete command builder will fail...
        Dim primKeyColumnsStr(primKeysCount - 1) As String
        Dim primKeyCompound As String = " WHERE " ' try to find record for update in dataset based on primary key where clause for avoidFill
        For i As Integer = 1 To primKeysCount
            Dim primKey As String = TargetRange.Cells(1, i).Value.ToString()
            If primKey Is Nothing OrElse primKey = "" Then
                notifyUserOfDataError("Primary key field name in column " + i.ToString() + " is blank !", 1, i)
                GoTo cleanup
            End If
            ' if primKey is in ignoreColumns then the only reasonable reason is a lookup primary key in DBSHeets (CUDFlags only), so try with "real" (resolved key) instead...
            If InStr(1, LCase(ignoreColumns) + ",", LCase(primKey) + ",") > 0 Then
                If CUDFlags Then
                    primKey = Left(primKey, Len(primKey) - 2) ' correct the LU to real primary Key
                Else
                    notifyUserOfDataError("Primary key '" + primKey + "' is contained in ignoreColumns !", 1, i)
                    GoTo cleanup
                End If
            End If
            If Not dscheck.Tables(0).Columns.Contains(primKey) Then
                notifyUserOfDataError("Primary key '" + primKey + "' not found in table '" + tableName, 1, i)
                GoTo cleanup
            End If
            primKeyColumnsStr(i - 1) = primKey
            primKeyCompound += primKey + " = @" + primKey + IIf(i = primKeysCount, "", " AND ")
        Next

        ' the actual dataset
        Dim ds As DataSet = New DataSet()
        ' replace the select command to avoid columns that are default filled but non-nullable (will produce errors if not assigned to new row)
        da.SelectCommand.CommandText = "SELECT " + String.Join(",", allColumnsStr) + " FROM " + tableName
        ' fill schema again to reflect the changed columns
        Try
            da.FillSchema(ds, SchemaType.Source, tableName)
        Catch ex As Exception
            notifyUserOfDataError("Error in getting schema information for " + tableName + " (" + da.SelectCommand.CommandText + "): " + ex.Message, 1)
            GoTo cleanup
        End Try

        ' for avoidFill mode (no uploading of whole table) replace select with parameterized query (params are added below)
        If avoidFill Then da.SelectCommand.CommandText = "SELECT " + String.Join(",", allColumnsStr) + " FROM " + tableName + primKeyCompound

        Dim allColumns(UBound(allColumnsStr)) As DataColumn
        For i As Integer = 0 To UBound(allColumnsStr)
            allColumns(i) = ds.Tables(0).Columns(allColumnsStr(i))
        Next

        ' assign primary key columns externally in case there are none defined on the table (might be)
        Dim primKeyColumns(primKeysCount - 1) As DataColumn
        For i As Integer = 0 To UBound(primKeyColumnsStr)
            primKeyColumns(i) = ds.Tables(0).Columns(primKeyColumnsStr(i))
            ' for avoidFill mode (no uploading of whole table) set up params for parameterized query from primary keys
            If avoidFill Then
                Dim param As Common.DbParameter
                If TypeName(idbcnn) = "SqlConnection" Then
                    param = DirectCast(da.SelectCommand, SqlClient.SqlCommand).CreateParameter()
                ElseIf TypeName(idbcnn) = "OleDbConnection" Then
                    param = DirectCast(da.SelectCommand, OleDb.OleDbCommand).CreateParameter()
                Else
                    param = DirectCast(da.SelectCommand, Odbc.OdbcCommand).CreateParameter()
                End If
                With param
                    .ParameterName = "@" + primKeyColumnsStr(i)
                    .SourceColumn = primKeyColumnsStr(i)
                    .DbType = TypeToDbType(ds.Tables(0).Columns(i).DataType(), primKeyColumnsStr(i), schemaDataTypeCollection)
                End With
                da.SelectCommand.Parameters.Add(param)
            End If
        Next
        Try
            ds.Tables(0).PrimaryKey = primKeyColumns
        Catch ex As Exception
            notifyUserOfDataError("Error in setting primary keys for " + tableName + ": " + ex.Message, 1)
            GoTo cleanup
        End Try
        ' for faster loading of data
        ds.Tables(0).BeginLoadData()
        ' fill the dataset in normal mode (needed to find records in memory)
        If Not avoidFill Then
            ExcelDnaUtil.Application.StatusBar = "filling the table data into dataset"
            Try
                da.Fill(ds.Tables(0))
            Catch ex As Exception
                If InStr(LCase(ex.Message()), "timeout") > 0 Or TypeOf ex Is System.OutOfMemoryException Then
                    notifyUserOfDataError("Timeout/OutOfMemoryException in retrieving Data for " + tableName + ": " + ex.Message + vbCrLf + vbCrLf + "You can usually resolve this problem by adding <avoidFill>True</avoidFill> to the DB Mappers definition!", 1)
                Else
                    notifyUserOfDataError("Error in retrieving Data for " + tableName + ": " + ex.Message + vbCrLf + "Following primary keys are defined (check whether enough): " + String.Join(Of DataColumn)(", ", primKeyColumns), 1)
                End If
                GoTo cleanup
            End Try
        End If
        ' set up the insert/update/delete CommandBuilders
        Try
            Dim custCmdBuilder As CustomCommandBuilder
            If TypeName(idbcnn) = "SqlConnection" Then
                custCmdBuilder = New CustomSqlCommandBuilder(ds.Tables(0), idbcnn, allColumns, schemaDataTypeCollection)
            ElseIf TypeName(idbcnn) = "OleDbConnection" Then
                custCmdBuilder = New CustomOleDbCommandBuilder(ds.Tables(0), idbcnn, allColumns, schemaDataTypeCollection)
            Else
                custCmdBuilder = New CustomOdbcCommandBuilder(ds.Tables(0), idbcnn, allColumns, schemaDataTypeCollection)
            End If
            da.UpdateCommand = custCmdBuilder.UpdateCommand()
            da.UpdateCommand.CommandTimeout = Globals.CmdTimeout
            da.DeleteCommand = custCmdBuilder.DeleteCommand()
            da.DeleteCommand.CommandTimeout = Globals.CmdTimeout
            da.InsertCommand = custCmdBuilder.InsertCommand()
            da.InsertCommand.CommandTimeout = Globals.CmdTimeout
        Catch ex As Exception
            notifyUserOfDataError("Error in setting Insert/Update/Delete Commands for Data Adapter for " + tableName + ": " + ex.Message, 1)
            GoTo cleanup
        End Try
        ExcelDnaUtil.Application.StatusBar = "Assigning transaction to CommandBuilders"
        Try
            da.UpdateCommand.UpdatedRowSource = UpdateRowSource.None
            da.UpdateCommand.Transaction = DBModifs.trans
            da.DeleteCommand.UpdatedRowSource = UpdateRowSource.None
            da.DeleteCommand.Transaction = DBModifs.trans
            da.InsertCommand.UpdatedRowSource = UpdateRowSource.None
            da.InsertCommand.Transaction = DBModifs.trans
        Catch ex As Exception
            notifyUserOfDataError("Error in setting Transaction for Insert/Update/Delete Commands for Data Adapter for " + tableName + ": " + ex.Message, 1)
            GoTo cleanup
        End Try

        Dim rowNum As Long = 2
        ' walk through all rows in DBMapper Targetrange to store in data set
        Dim finishLoop As Boolean
        Do
            ' if CUDFlags are set, only insert/update/delete if CUDFlags column (right to DBMapper range) is filled...
            Dim rowCUDFlag As String = TargetRange.Cells(rowNum, TargetRange.Columns.Count + 1).Value
            If Not CUDFlags Or (CUDFlags And rowCUDFlag <> "") Then
                Dim AutoIncrement As Boolean = False
                Dim primKeyValues(primKeysCount - 1) As Object
                Dim primKeyValueStr As String = ""
                For i As Integer = 1 To primKeysCount
                    Dim primKey As String = TargetRange.Cells(1, i).Value.ToString()
                    Dim primKeyValue = TargetRange.Cells(rowNum, i).Value
                    If IsXLCVErr(primKeyValue) Then
                        notifyUserOfDataError("Error in primary key value: " + CVErrText(primKeyValue) + " in sheet " + TargetRange.Parent.Name + ", row " + rowNum.ToString() + ", col " + i.ToString(), rowNum, i)
                        GoTo cleanup
                    End If
                    If primKeyValue Is Nothing Then primKeyValue = ""
                    ' if primKey is in ignoreColumns then the only reasonable reason is a lookup primary key in DBSHeets (CUDFlags only), so try with "real" (resolved key) instead...
                    If InStr(1, LCase(ignoreColumns) + ",", LCase(primKey) + ",") > 0 Then
                        If CUDFlags Then
                            primKey = Left(primKey, Len(primKey) - 2) ' correct the LU to real primary Key
                            If Left(ds.Tables(0).Columns(primKey).DataType.Name, 4) = "Date" Then
                                primKeyValue = TargetRange.ListObject.ListColumns(primKey).Range(rowNum, 1).Value
                            Else
                                primKeyValue = TargetRange.ListObject.ListColumns(primKey).Range(rowNum, 1).Value2
                            End If
                        End If
                    End If
                    ' special treatment for date(time) fields, try to convert from double (OLE Automation standard: julian datetime values) if not properly formatted
                    If Left(ds.Tables(0).Columns(primKey).DataType.Name, 4) = "Date" Then
                        If TypeName(primKeyValue) = "Double" Then primKeyValue = Date.FromOADate(primKeyValue)
                    End If
                    ' empty primary keys are valid if primary key has autoincrement property defined, so pass DBNull Value here to avoid exception in finding record below (record is not found of course)...
                    primKeyValues(i - 1) = IIf(IsNothing(primKeyValue) OrElse primKeyValue.ToString() = "", DBNull.Value, primKeyValue)
                    If avoidFill Then
                        da.SelectCommand.Parameters.Item("@" + primKey).Value = primKeyValues(i - 1)
                    End If
                    Dim checkAutoIncrement As Boolean = ds.Tables(0).Columns(primKey).AutoIncrement
                    If Not checkAutoIncrement And Len(primKeyValue) = 0 Then
                        If Not notifyUserOfDataError("AutoIncrement property for primary key '" + primKey + "' is false and primary key value is empty!", 1, i) Then GoTo cleanup
                        GoTo nextRow
                    End If
                    ' single primary key, AutoIncFlag is to indicate first column might be left empty in such cases. CUD DBMappers can have multiple empty primary keys as long as they have auto identity.
                    If primKeysCount = 1 And (CUDFlags Or AutoIncFlag) And primKeyValue.ToString().Length = 0 And checkAutoIncrement Then
                        AutoIncrement = True ' needed to avoid searching for primary key(s) that are empty because of auto identity
                        Exit For
                    End If
                    ' with CUDFlags and multiple primary keys there can be empty primary keys (auto identity columns), leave error checking to database in this case ...
                    If (Not CUDFlags Or (CUDFlags And rowCUDFlag = "u")) And primKeyValue.ToString().Length = 0 Then
                        If Not notifyUserOfDataError("Empty primary key value in sheet " + TargetRange.Parent.Name + ", row " + rowNum.ToString() + ", col " + i.ToString(), rowNum, i) Then GoTo cleanup
                        GoTo nextRow
                    End If
                    primKeyValueStr += IIf(primKeyValueStr <> "", ",", "") + primKeyValue.ToString()
                Next
                ' if we avoid the full table fill at the beginning, select the single rows to be updated here...
                If avoidFill Then
                    Try
                        da.Fill(ds.Tables(0))
                    Catch ex As Exception
                        If Not notifyUserOfDataError("Error in retrieving Data for " + tableName + ": " + ex.Message, rowNum) Then GoTo cleanup
                        GoTo nextRow
                    End Try
                End If

                ' get the record for updating, however avoid finding record with empty primary key value if autoincrement is given...
                Dim foundRow As DataRow = Nothing
                If Not AutoIncrement Then
                    Try
                        foundRow = ds.Tables(0).Rows.Find(primKeyValues)
                    Catch ex As Exception
                        If Not notifyUserOfDataError("Problem getting record, Error: " + ex.Message + " in sheet " + TargetRange.Parent.Name + ", row " + rowNum.ToString(), rowNum) Then GoTo cleanup
                        GoTo nextRow
                    End Try
                End If

                Dim insertRecord As Boolean = False
                ' If we have an autoincrementing primary key (empty primary key value !) or didn't find a record with the given primary key (rst.EOF) ...
                If AutoIncrement OrElse IsNothing(foundRow) Then
                    If insertIfMissing Or rowCUDFlag = "i" Then
                        insertRecord = True
                        ' ... add a new record if insertIfMissing flag is set Or CUD Flag insert is given
                        foundRow = ds.Tables(0).NewRow()
                        For i As Integer = 1 To primKeysCount
                            Dim primKey As String = TargetRange.Cells(1, i).Value.ToString()
                            Dim primKeyValue As Object = TargetRange.Cells(rowNum, i).Value
                            If primKeyValue Is Nothing Then primKeyValue = ""
                            ' if primKey is in ignoreColumns then the only reasonable reason is a lookup primary key in DBSheets (CUDFlags only), so try with "real" (resolved key) instead...
                            If InStr(1, LCase(ignoreColumns) + ",", LCase(primKey) + ",") > 0 AndAlso CUDFlags Then
                                primKey = Left(primKey, Len(primKey) - 2) ' correct the LU to real primary Key
                                If Left(ds.Tables(0).Columns(primKey).DataType.Name, 4) = "Date" Then
                                    primKeyValue = TargetRange.ListObject.ListColumns(primKey).Range(rowNum, 1).Value
                                Else
                                    primKeyValue = TargetRange.ListObject.ListColumns(primKey).Range(rowNum, 1).Value2
                                End If
                            End If
                            ' special treatment for date(time) fields, try to convert from double (OLE Automation standard: julian datetime values) if not properly formatted
                            If Left(ds.Tables(0).Columns(primKey).DataType.Name, 4) = "Date" Then
                                If TypeName(primKeyValue) = "Double" Then primKeyValue = Date.FromOADate(primKeyValue)
                            End If
                            Try
                                ' skip empty primary field values for autoincrementing identity fields ..
                                If CStr(primKeyValue) <> "" Then foundRow(primKey) = primKeyValue
                            Catch ex As Exception
                                If Not notifyUserOfDataError("Error inserting primary key value into table " + tableName + ": " + ex.Message, rowNum, i) Then GoTo cleanup
                                GoTo nextRow
                            End Try
                        Next
                    Else
                        If Not notifyUserOfDataError("Did Not find record with primary keys '" + primKeyValueStr + "', insertIfMissing = " + insertIfMissing.ToString() + " in sheet " + TargetRange.Parent.Name + ", row " + rowNum.ToString(), rowNum) Then GoTo cleanup
                        GoTo nextRow
                    End If
                    ExcelDnaUtil.Application.StatusBar = Left("Inserting " + IIf(AutoIncrement, "new autoincremented key", primKeyValueStr) + " into " + tableName, 255)
                End If
                ' fill non primary key fields to prepare record for insert or update
                If Not CUDFlags Or (CUDFlags And (rowCUDFlag = "i" Or rowCUDFlag = "u")) Then
                    colNum = primKeysCount + 1
                    Do
                        Dim fieldname As String = TargetRange.Cells(1, colNum).Value
                        If InStr(1, LCase(ignoreColumns) + ",", LCase(fieldname) + ",") = 0 Then
                            Try
                                Dim fieldval As Object = TargetRange.Cells(rowNum, colNum).Value
                                If Left(ds.Tables(0).Columns(fieldname).DataType.Name, 4) = "Date" Then
                                    fieldval = TargetRange.Cells(rowNum, colNum).Value
                                Else
                                    fieldval = TargetRange.Cells(rowNum, colNum).Value2
                                End If
                                If fieldval Is Nothing Then
                                    foundRow(fieldname) = DBNull.Value ' explicitly set DBNull Value, Nothing or null doesn't work here
                                Else
                                    If IsXLCVErr(fieldval) Then
                                        If IgnoreDataErrors Then
                                            foundRow(fieldname) = DBNull.Value ' if data errors are ignored, set DBNull Value
                                        Else
                                            If Not notifyUserOfDataError("Field Value Update Error: " + CVErrText(fieldval) + " with Table: " + tableName + ", Field: " + fieldname + ", in sheet " + TargetRange.Parent.Name + ", row " + rowNum.ToString() + ", col " + colNum.ToString(), rowNum, colNum) Then GoTo cleanup
                                            GoTo nextRow
                                        End If
                                    Else
                                        ' special treatment for date(time) fields, try to convert from double (OLE Automation standard: julian datetime values) if not properly formatted
                                        If Left(ds.Tables(0).Columns(fieldname).DataType.Name, 4) = "Date" Then
                                            If TypeName(fieldval) = "Double" Then fieldval = Date.FromOADate(fieldval)
                                        End If
                                        Try
                                            foundRow(fieldname) = IIf(fieldval.ToString().Length = 0, DBNull.Value, fieldval)
                                        Catch ex As Exception
                                            notifyUserOfDataError("Error: " + ex.Message + " in assigning value in sheet " + TargetRange.Parent.Name + ", row " + rowNum.ToString() + ", col " + colNum.ToString(), rowNum, colNum)
                                        End Try
                                    End If
                                End If
                            Catch ex As Exception
                                Dim exMessage As String = ex.Message
                                foundRow.CancelEdit()
                                If Not notifyUserOfDataError("Field Value Insert or Update Error: " + exMessage + " with Table: " + tableName + ", Field: " + fieldname + ", in sheet " + TargetRange.Parent.Name + ", row " + rowNum.ToString() + ", col " + colNum.ToString(), rowNum, colNum) Then GoTo cleanup
                                GoTo nextRow
                            End Try
                        End If
                        colNum += 1
                    Loop Until colNum > TargetRange.Columns.Count
                    ExcelDnaUtil.Application.StatusBar = Left("Filled fields for " + primKeyValueStr + " in " + tableName, 255)
                    If insertRecord Then
                        Try
                            ds.Tables(0).Rows.Add(foundRow)
                        Catch ex As Exception
                            If Not notifyUserOfDataError("Error inserting row " + rowNum.ToString() + " in sheet " + TargetRange.Parent.Name + ": " + ex.Message, rowNum) Then GoTo cleanup
                        End Try
                    End If
                End If

                ' delete only with CUDFlags...
                If (CUDFlags And rowCUDFlag = "d") Then
                    Try
                        foundRow.Delete()
                    Catch ex As Exception
                        If Not notifyUserOfDataError("Error deleting row " + rowNum.ToString() + " in sheet " + TargetRange.Parent.Name + ": " + ex.Message, rowNum) Then GoTo cleanup
                    End Try
                    ExcelDnaUtil.Application.StatusBar = Left("Deleting " + primKeyValueStr + " in " + tableName, 255)
                End If

nextRow:
                Try
                    If TargetRange.Cells(rowNum + 1, 1).Value Is Nothing OrElse TargetRange.Cells(rowNum + 1, 1).Value.ToString().Length = 0 Then
                        ' avoid for CUD DBMappers and auto incrementing situations (empty primary keys)
                        If Not (CUDFlags Or AutoIncFlag) Then finishLoop = True
                    End If
                Catch ex As Exception
                    If Not notifyUserOfDataError("Error in first primary column: Cells(" + (rowNum + 1).ToString() + ",1): " + ex.Message, rowNum + 1) Then GoTo cleanup
                    'finishLoop = True '-> do not finish to allow erroneous data  !!
                End Try
            End If
            rowNum += 1
        Loop Until rowNum > TargetRange.Rows.Count Or finishLoop

        ' now update the changes in the DB
        ExcelDnaUtil.Application.StatusBar = Left("Doing modifications (inserts/updates/deletes) in Database for " + tableName, 255)
        Try
            da.Update(ds.Tables(0))
            changesDone = True
        Catch ex As Exception
            Dim exMessage As String = ex.Message
            If Not notifyUserOfDataError("Row Update Error, Table: " + tableName + ", Error: " + exMessage + " in sheet " + TargetRange.Parent.Name, rowNum) Then GoTo cleanup
        End Try

        ' any additional stored procedures to execute?
        If executeAdditionalProc.Length > 0 Then
            Dim result As Integer
            Try
                ExcelDnaUtil.Application.StatusBar = "executing stored procedure " + executeAdditionalProc
                Dim storedProcCmd As IDbCommand
                If TypeName(idbcnn) = "SqlConnection" Then
                    storedProcCmd = New SqlClient.SqlCommand(executeAdditionalProc, idbcnn, trans)
                ElseIf TypeName(idbcnn) = "OleDbConnection" Then
                    storedProcCmd = New OleDb.OleDbCommand(executeAdditionalProc, idbcnn, trans)
                Else
                    storedProcCmd = New Odbc.OdbcCommand(executeAdditionalProc, idbcnn, trans)
                End If
                storedProcCmd.CommandText = executeAdditionalProc
                result = storedProcCmd.ExecuteNonQuery()
            Catch ex As Exception
                hadError = True
                Globals.UserMsg("Error in executing additional stored procedure: " + ex.Message, "DBMapper Error")
                GoTo cleanup
            End Try
            LogInfo("executed " + executeAdditionalProc + ", affected rows: " + result.ToString())
        End If
cleanup:
        ExcelDnaUtil.Application.StatusBar = False
        ' close connection to return it to the pool (automatically closes recordset objects, so no need for checkrst.Close() or rst.Close())...
        If calledByDBSeq = "" Then idbcnn.Close()
        ' DBSheet surrogate (CUDFlags), ask for refresh after DB Modification was done
        If changesDone Then
            Dim DBFunctionSrcExtent = getUnderlyingDBNameFromRange(TargetRange)
            If DBFunctionSrcExtent <> "" Then
                If CUDFlags Then
                    ' also resetCUDFlags for CUDFlags DBMapper that do not ask before execute and were called by a DBSequence
                    Try
                        ' reset CUDFlags before refresh to avoid problems with reduced TargetRange due to deletions!
                        Me.resetCUDFlags()
                    Catch ex As Exception
                        Globals.UserMsg("Error in resetting CUD Flags: " + ex.Message, "DBMapper Error")
                    End Try
                    If calledByDBSeq = "" Then
                        Dim retval As MsgBoxResult
                        ' only ask when DBModifier not done on Workbook save, in this case refresh automatically...
                        If Not WbIsSaving Then retval = QuestionMsg(theMessage:="Refresh Data Range of DB Mapper '" + dbmodifName + "' ?", questionTitle:="Refresh DB Mapper")
                        If WbIsSaving Or retval = vbOK Then
                            doDBRefresh(Replace(DBFunctionSrcExtent, "DBFtarget", "DBFsource"))
                            ' clear CUD marks after completion is done with doDBRefresh/DBSetQueryAction/resizeDBMapperRange
                        End If
                    Else
                        Globals.LogWarn("no refresh took place for DBMapper " + dbmodifName)
                    End If
                End If
            End If
        End If
    End Sub

    ''' <summary>notfication of error for user including selection of error cell</summary>
    ''' <param name="message">error message</param>
    ''' <param name="rowNum">error cell row</param>
    ''' <param name="colNum">error cell column</param>
    ''' <returns></returns>
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

    ''' <summary>set the fields in the DB Modifier Create Dialog with attributes of object</summary>
    ''' <param name="theDBModifCreateDlg"></param>
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
            .AutoIncFlag.Checked = AutoIncFlag
            .IgnoreDataErrors.Checked = IgnoreDataErrors
            .AskForExecute.Checked = askBeforeExecute
        End With
    End Sub

End Class

''' <summary>DBActions are used to issue DML commands defined in Cells against the database</summary>
Public Class DBAction : Inherits DBModif

    ''' <summary>normal constructor with definition xml</summary>
    ''' <param name="definitionXML"></param>
    ''' <param name="paramTarget"></param>
    Public Sub New(definitionXML As CustomXMLNode, paramTarget As Excel.Range)
        MyBase.New(definitionXML)
        Try
            ' if no target range is set, then no parameters can be found...
            If paramTarget Is Nothing Then Exit Sub
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
            Globals.UserMsg("Error when creating DB Action '" + dbmodifName + "': " + ex.Message, "DBModifier Definitions Error")
        End Try
    End Sub

    Public Overrides Sub doDBModif(Optional WbIsSaving As Boolean = False, Optional calledByDBSeq As String = "", Optional TransactionOpen As Boolean = False)
        ' ask for saving only if a) is not done on WorkbookSave b) is set to ask and c) is not called by a DBSequence (asks already for saving) and d) is in interactive mode
        If Not WbIsSaving And askBeforeExecute And calledByDBSeq = "" And Not nonInteractive Then
            If Not confirmExecution() = MsgBoxResult.Yes Then Exit Sub
        End If

        'now create/get a database specific connection (idbcnn) in case it was not already created by the sequence (transactions!)
        If Not TransactionOpen Then
            If Not openDatabase() Then Exit Sub
        End If

        Dim result As Integer
        Try
            ExcelDnaUtil.Application.StatusBar = "executing DBAction " + paramTargetName
            Dim executeText As String = ""
            For Each targetCell As Excel.Range In TargetRange
                executeText += targetCell.Text + " "
            Next
            Dim DmlCmd As IDbCommand
            If TypeName(idbcnn) = "SqlConnection" Then
                DmlCmd = New SqlClient.SqlCommand(executeText, idbcnn, trans)
            ElseIf TypeName(idbcnn) = "OleDbConnection" Then
                DmlCmd = New OleDb.OleDbCommand(executeText, idbcnn, trans)
            Else
                DmlCmd = New Odbc.OdbcCommand(executeText, idbcnn, trans)
            End If
            DmlCmd.CommandTimeout = Globals.CmdTimeout
            DmlCmd.CommandType = CommandType.Text
            result = DmlCmd.ExecuteNonQuery()
        Catch ex As Exception
            hadError = True
            Globals.UserMsg("Error in executing DB Action " + paramTargetName + ": " + ex.Message, "DBAction Error")
            ExcelDnaUtil.Application.StatusBar = False
            Exit Sub
        End Try
        If Not WbIsSaving And calledByDBSeq = "" Then
            Globals.UserMsg("DBAction " + paramTargetName + " executed, affected records: " + result.ToString(), "DBAction executed", MsgBoxStyle.Information)
        End If
        ExcelDnaUtil.Application.StatusBar = False
        ' close connection to return it to the pool...
        If calledByDBSeq = "" Then idbcnn.Close()
    End Sub

    ''' <summary>set the fields in the DB Modifier Create Dialog with attributes of object</summary>
    ''' <param name="theDBModifCreateDlg"></param>
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
            Globals.UserMsg("Error when creating DB Sequence '" + dbmodifName + "': " + ex.Message, "DBModifier Definitions Error")
        End Try
    End Sub

    Public Overrides Sub doDBModif(Optional WbIsSaving As Boolean = False, Optional calledByDBSeq As String = "", Optional TransactionOpen As Boolean = False)
        Dim TransactionIsOpen As Boolean = False
        hadError = False
        ' warning against recursions (should not happen...)
        If calledByDBSeq <> "" Then
            Globals.UserMsg("DB Sequence '" + dbmodifName + "' is being called by another DB Sequence (" + calledByDBSeq + "), this should not occur as infinite recursions are possible !", "Execute DB Sequence")
            Exit Sub
        End If
        ' ask for saving only if a) is not done on WorkbookSave b) is set to ask and c) is in interactive mode
        If Not WbIsSaving And askBeforeExecute And Not nonInteractive Then
            If Not confirmExecution() = MsgBoxResult.Yes Then Exit Sub
        End If

        ' reset the db connection in any case to allow for new connections at DBBegin
        idbcnn = Nothing
        Dim executedDBMappers As Dictionary(Of String, Boolean) = New Dictionary(Of String, Boolean)
        Dim modifiedDBMappers As Dictionary(Of String, Boolean) = New Dictionary(Of String, Boolean)
        For i As Integer = 0 To UBound(sequenceParams)
            Dim definition() As String = Split(sequenceParams(i), ":")
            Dim DBModiftype As String = definition(0)
            Dim DBModifname As String = definition(1)
            Select Case DBModiftype
                Case "DBMapper", "DBAction"
                    If Not hadError Then
                        Globals.LogInfo(DBModifname + "... ")
                        DBModifDefColl(DBModiftype).Item(DBModifname).doDBModif(WbIsSaving, calledByDBSeq:=MyBase.dbmodifName, TransactionOpen:=TransactionIsOpen)
                        If DBModiftype = "DBMapper" Then
                            If DirectCast(DBModifDefColl("DBMapper").Item(DBModifname), DBMapper).CUDFlags Then executedDBMappers(DBModifname) = True
                            If DirectCast(DBModifDefColl("DBMapper").Item(DBModifname), DBMapper).hadChanges Then modifiedDBMappers(DBModifname) = True
                        End If
                    End If
                Case "DBBegin"
                    Globals.LogInfo("DBBeginTrans... ")
                    If idbcnn Is Nothing Then
                        ' take database connection properties from sequence step after DBBegin. (checked) requirement: all steps have the same connection in this case!
                        Dim nextdefinition() As String = Split(sequenceParams(i + 1), ":")
                        If Not DBModifDefColl(nextdefinition(0)).Item(nextdefinition(1)).openDatabase(env) Then Exit Sub
                    End If
                    DBModifs.trans = idbcnn.BeginTransaction()
                    TransactionIsOpen = True
                Case "DBCommitRollback"
                    If Not hadError Then
                        Globals.LogInfo("DBCommitTrans... ")
                        DBModifs.trans.Commit()
                    Else
                        Globals.LogInfo("DBRollbackTrans... ")
                        DBModifs.trans.Rollback()
                    End If
                    TransactionIsOpen = False
                Case Else
                    If Not hadError Then
                        If Left(DBModiftype, 8) = "Refresh " Then
                            doDBRefresh(srcExtent:=DBModifname, executedDBMappers:=executedDBMappers, modifiedDBMappers:=modifiedDBMappers, TransactionIsOpen:=TransactionIsOpen)
                        Else
                            Globals.UserMsg("Unknown type of sequence step '" + DBModiftype + "' being called in DB Sequence (" + calledByDBSeq + ") !", "Execute DB Sequence")
                        End If
                    End If
            End Select
        Next
    End Sub

    ''' <summary>required for creation of DB Sequences after finishing dialog with OK button</summary>
    ''' <returns></returns>
    Public Function getSequenceSteps() As String()
        Return sequenceParams
    End Function

    ''' <summary>set the fields in the DB Modifier Create Dialog with attributes of object</summary>
    ''' <param name="theDBModifCreateDlg"></param>
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

    ''' <summary>main db connection for DB modifiers</summary>
    Public idbcnn As System.Data.IDbConnection
    ''' <summary>avoid entering Application.SheetChange Event handler during listfetch/setquery</summary>
    Public preventChangeWhileFetching As Boolean = False
    ''' <summary>indicates an error in execution of DBModifiers, used for commit/rollback and for noninteractive message return</summary>
    Public hadError As Boolean
    ''' <summary>used to work around the fact that when started by Application.Run, Formulas are sometimes returned as local</summary>
    Public listSepLocal As String = ExcelDnaUtil.Application.International(Excel.XlApplicationInternational.xlListSeparator)
    ''' <summary>common transaction, needed for DBSequence and all other DB Modifiers</summary>
    Public trans As Common.DbTransaction = Nothing

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

        Dim theConnString As String = Globals.fetchSetting("ConstConnString" + env.ToString(), "")
        If theConnString = "" Then
            Globals.UserMsg("No Connectionstring given for environment: " + env.ToString() + ", please correct and rerun.", "Open Connection Error")
            Exit Function
        End If
        Dim dbidentifier As String = fetchSetting("DBidentifierCCS" + env.ToString(), "")
        If dbidentifier = "" Then
            Globals.UserMsg("No DB identifier given for environment: " + env.ToString() + ", please correct and rerun.", "Open Connection Error")
            Exit Function
        End If
        ' change the database in the connection string
        theConnString = Globals.Change(theConnString, dbidentifier, database, ";")
        ' need to change/set the connection timeout in the connection string as the property is readonly then...
        If InStr(theConnString, "Connection Timeout=") > 0 Then
            theConnString = Globals.Change(theConnString, "Connection Timeout=", Globals.CnnTimeout.ToString(), ";")
        ElseIf InStr(theConnString, "Connect Timeout=") > 0 Then
            theConnString = Globals.Change(theConnString, "Connect Timeout=", Globals.CnnTimeout.ToString(), ";")
        Else
            theConnString += ";Connection Timeout=" + Globals.CnnTimeout.ToString()
        End If

        Try
            If InStr(theConnString.ToLower, "provider=sqloledb") Or InStr(theConnString.ToLower, "driver=sql server") Then
                ' remove provider=SQLOLEDB; (or whatever is in ConnStringSearch<>) for sql server as this is not allowed for ado.net (legacy adodb)
                theConnString = Replace(theConnString, Globals.fetchSetting("ConnStringSearch" + Globals.env(), "provider=SQLOLEDB") + ";", "")
                idbcnn = New SqlClient.SqlConnection(theConnString)
            ElseIf InStr(theConnString.ToLower, "oledb") Then
                idbcnn = New OleDb.OleDbConnection(theConnString)
            Else
                idbcnn = New Odbc.OdbcConnection(theConnString)
            End If
        Catch ex As Exception
            Globals.UserMsg("Error creating connection object: " + ex.Message + ", connection string: " + theConnString, "Open Connection Error")
            idbcnn = Nothing
            ExcelDnaUtil.Application.StatusBar = False
            Exit Function
        End Try

        Globals.LogInfo("open connection with " + theConnString)
        ExcelDnaUtil.Application.StatusBar = "Trying " + Globals.CnnTimeout.ToString() + " sec. with connstring: " + theConnString
        Try
            idbcnn.Open()
            openIdbConnection = True
        Catch ex As Exception
            Globals.UserMsg("Error connecting to DB: " + ex.Message + ", connection string: " + theConnString, "Open Connection Error")
            idbcnn = Nothing
        End Try
        ExcelDnaUtil.Application.StatusBar = False
    End Function

    ''' <summary>in case there is a defined DBMapper underlying the DBListFetch/DBSetQuery target area then change the extent of it (oldRange) to the new area given in theRange</summary>
    ''' <param name="theRange">new extent after resfresh of DBListFetch/DBSetQuery function</param>
    ''' <param name="oldRange">extent before resfresh of DBListFetch/DBSetQuery function</param>
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
                ' pass the associated DBMapper the new target range
                Try
                    Dim extendedMapper As DBMapper = Globals.DBModifDefColl("DBMapper").Item(dbMapperRangeName)
                    extendedMapper.setTargetRange(theRange)
                Catch ex As Exception
                    Throw New Exception("Error passing new Range to the associated DBMapper object when extending '" + dbMapperRangeName + "' to DBListFetch/DBSetQuery target range: " + ex.Message)
                End Try
            End If
        End If
    End Sub

    ''' <summary>creates a DBModif at the current active cell or edits an existing one defined in targetDefName (after being called in defined range or from ribbon + Ctrl + Shift)</summary>
    Public Sub createDBModif(createdDBModifType As String, Optional targetDefName As String = "")
        ' clipboard helper for legacy definitions:
        ' if saveRangeToDB<Single> macro calls were copied into clipboard, 1st parameter (datarange) removed (empty), connid moved to 2nd place as database name (remove MSSQL)
        'mapper.saveRangeToDBSingle(Range("DB_DefName"), "tableName", "primKey1,primKey2,primKey3", "MSSQLDB_NAME", True) 
        '       -> def(, "DB_NAME", "tableName", "primKey1,primKey2,primKey3", True)    DBMapperName = DB_DefName
        'mapper.saveRangeToDBSingle(Range("DB_DefName"), "tableName", "primKey1,primKey2,primKey3", "MSSQLDB_NAME")       
        '       -> def(, "DB_NAME", "tableName", "primKey1,primKey2,primKey3")          DBMapperName = DB_DefName
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
                    Globals.UserMsg("Error when retrieving DB_DefName from clipboard: " + ex.Message, "DBMapper Legacy Creation Error") : Exit Sub
                End Try
                Try : newDefString = "def(" + Replace(Mid(cpbdtext, commaBeforeConnDef, commaAfterConnDef - commaBeforeConnDef), "MSSQL", "") + Mid(cpbdtext, firstComma, commaBeforeConnDef - firstComma - 1) + Mid(cpbdtext, commaAfterConnDef - 1)
                Catch ex As Exception
                    Globals.UserMsg("Error when building new definition from clipboard: " + ex.Message, "DBMapper Legacy Creation Error") : Exit Sub
                End Try
                ' assign new name to active cell
                ' Add doesn't work directly with ExcelDnaUtil.Application.ActiveWorkbook.Names (late binding), so create an object here...
                Dim NamesList As Excel.Names = ExcelDnaUtil.Application.ActiveWorkbook.Names
                Try : NamesList.Add(Name:=DB_DefName, RefersTo:=ExcelDnaUtil.Application.ActiveCell)
                Catch ex As Exception
                    Globals.UserMsg("Error when assigning name '" + DB_DefName + "' to active cell: " + ex.Message, "DBMapper Legacy Creation Error")
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
            If createdDBModifType <> "DBSeqnce" Then
                Dim existingDefRange As Excel.Range = Nothing
                Try
                    existingDefRange = ExcelDnaUtil.Application.Range(existingDefName)
                Catch ex As Exception
                    ' if target name relates to an invalid (offset) formula, getting a range fails  ...
                    If InStr(ExcelDnaUtil.Application.ActiveWorkbook.Names.Item(existingDefName).RefersTo, "OFFSET(") > 0 Then
                        Globals.UserMsg("Offset formula that '" + existingDefName + "' refers to, did not return a valid range." + vbCrLf + "Please check the offset formula to return a valid range !", "DBModifier Definitions Error")
                        ExcelDnaUtil.Application.Dialogs(Excel.XlBuiltInDialog.xlDialogNameManager).Show()
                        Exit Sub
                    End If
                End Try
                existingDBModif.setTargetRange(existingDefRange)
            End If
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
                .AutoIncFlag.Hide()
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
                Else
                    theDBModifCreateDlg.MinimumSize = New Drawing.Size(width:=490, height:=290)
                    theDBModifCreateDlg.Size = New Drawing.Size(width:=490, height:=290)
                End If
                ' hide controls irrelevant for DBMapper and DBAction
                .DBSeqenceDataGrid.Hide()
            End If

            ' delegate filling of dialog fields to created DBModif object
            If existingDBModif IsNot Nothing Then existingDBModif.setDBModifCreateFields(theDBModifCreateDlg)

            ' display dialog for parameters
            If theDBModifCreateDlg.ShowDialog() = DialogResult.Cancel Then
                ' remove targetRange Name created in clipboard helper
                If createdDBMapperFromClipboard Then
                    Try : ExcelDnaUtil.Application.ActiveWorkbook.Names(existingDefName).Delete : Catch ex As Exception : End Try
                End If
                Exit Sub
            End If

            ' only for DBMapper or DBAction: change or add target range name
            If createdDBModifType <> "DBSeqnce" Then
                Dim targetRange As Excel.Range
                If existingDBModif Is Nothing Then
                    targetRange = ExcelDnaUtil.Application.ActiveCell
                Else
                    targetRange = existingDBModif.getTargetRange()
                End If

                Dim NamesList As Excel.Names = ExcelDnaUtil.Application.ActiveWorkbook.Names
                If existingDefName = "" Then
                    Try
                        NamesList.Add(Name:=createdDBModifType + .DBModifName.Text, RefersTo:=targetRange)
                    Catch ex As Exception
                        Globals.UserMsg("Error when assigning name '" + createdDBModifType + .DBModifName.Text + "' to active cell: " + ex.Message, "DBModifier Creation Error")
                        Exit Sub
                    End Try
                Else
                    NamesList.Item(existingDefName).Name = createdDBModifType + .DBModifName.Text
                End If
            End If

            Dim CustomXmlParts As Object = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
            If CustomXmlParts.Count = 0 Then
                ' in case no CustomXmlPart in Namespace DBModifDef exists in the workbook, add one
                ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.Add("<root xmlns=""DBModifDef""></root>")
                CustomXmlParts = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
            End If

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
        Globals.LogInfo("reading DBModifier Definitions for Workbook: " + ExcelDnaUtil.Application.ActiveWorkbook.Name)
        Try
            Globals.DBModifDefColl.Clear()
            Dim CustomXmlParts As Object = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
            If CustomXmlParts.Count = 1 Then
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
                        Globals.LogInfo("reading DBModifier Definition for " + nodeName)
                        Dim targetRange As Excel.Range = Nothing
                        ' for DBMappers and DBActions the data of the DBModification is stored in Ranges, so check for those and get the Range
                        If DBModiftype = "DBMapper" Or DBModiftype = "DBAction" Then
                            For Each rangename As Excel.Name In ExcelDnaUtil.Application.ActiveWorkbook.Names
                                Dim rangenameName As String = Replace(rangename.Name, rangename.Parent.Name + "!", "")
                                If rangenameName = nodeName And InStr(rangename.RefersTo, "#REF!") > 0 Then
                                    Globals.UserMsg(DBModiftype + " definitions range [" + rangename.Parent.Name + "]" + rangename.Name + " contains #REF!", "DBModifier Definitions Error")
                                    Exit For
                                ElseIf rangenameName = nodeName Then
                                    ' might fail...
                                    Try
                                        targetRange = rangename.RefersToRange
                                    Catch ex As Exception
                                        ' if target name relates to an invalid (offset) formula, referstorange fails  ...
                                        If InStr(rangename.RefersTo, "OFFSET(") > 0 Then
                                            Globals.UserMsg("Offset formula that '" + nodeName + "' refers to, did not return a valid range." + vbCrLf + "Please check the offset formula to return a valid range !", "DBModifier Definitions Error")
                                            ExcelDnaUtil.Application.Dialogs(Excel.XlBuiltInDialog.xlDialogNameManager).Show()
                                            GoTo EndOuterLoop
                                        End If
                                    End Try
                                    Exit For
                                End If
                            Next
                            If targetRange Is Nothing Then
                                Dim answer As MsgBoxResult = Globals.QuestionMsg("Required target range named '" + nodeName + "' cannot be found for this " + DBModiftype + " definition." + vbCrLf + "Should the target range name and definition be removed (If you still need the " + DBModiftype + ", (re)create the target range with this name again)?", , "DBModifier Definitions Error", MsgBoxStyle.Critical)
                                If answer = MsgBoxResult.Ok Then
                                    ' remove name, in case it still exists
                                    Try : ExcelDnaUtil.Application.ActiveWorkbook.Names(nodeName).Delete() : Catch ex As Exception : End Try
                                    ' remove node
                                    If Not IsNothing(CustomXmlParts(1).SelectSingleNode("/ns0:root/ns0:" + DBModiftype + "[@Name='" + Replace(nodeName, DBModiftype, "") + "']")) Then
                                        Try : CustomXmlParts(1).SelectSingleNode("/ns0:root/ns0:" + DBModiftype + "[@Name='" + Replace(nodeName, DBModiftype, "") + "']").Delete : Catch ex As Exception
                                            Globals.UserMsg("Error removing node in DBModif definitions: " + ex.Message)
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
                            Globals.UserMsg("Not supported DBModiftype: " + DBModiftype, "DBModifier Definitions Error")
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
                                    Globals.UserMsg("DBModifier " + nodeName + " added twice, this indicates legacy definitions that were modified!" + vbCrLf + "To fix, convert all other definitions in the same way and then remove the legacy definitions by editing the raw DB Modif definitions.", "get DBModif Definitions")
                                Else
                                    defColl.Add(nodeName, newDBModif)
                                End If
                            End If
                        End If
                    End If
EndOuterLoop:
                Next
            ElseIf CustomXmlParts.Count > 1 Then
                Globals.UserMsg("Multiple CustomXmlParts for DBModifDef existing!", "get DBModif Definitions")
            End If
            Globals.theRibbon.Invalidate()
        Catch ex As Exception
            Globals.UserMsg("Exception:  " + ex.Message, "get DBModif Definitions")
        End Try
    End Sub

    ''' <summary>gets DB Modification Name (DBMapper or DBAction) from theRange</summary>
    ''' <param name="theRange"></param>
    ''' <returns>the retrieved name as a string (not name object !)</returns>
    Public Function getDBModifNameFromRange(theRange As Excel.Range) As String
        Dim nm As Excel.Name
        Dim rng, testRng As Excel.Range

        getDBModifNameFromRange = ""
        If theRange Is Nothing Then Exit Function
        Try
            ' try all names in workbook
            For Each nm In theRange.Parent.Parent.Names
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
            Globals.UserMsg("Exception: " + ex.Message, "get DBModif Name From Range")
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
        hadError = False : Globals.nonInteractive = headLess
        Globals.nonInteractiveErrMsgs = "" ' reset noninteractive messages
        Dim DBModiftype As String = Left(DBModifName, 8)
        If DBModiftype = "DBSeqnce" Or DBModiftype = "DBMapper" Or DBModiftype = "DBAction" Then
            If Not Globals.DBModifDefColl(DBModiftype).ContainsKey(DBModifName) Then
                If Globals.DBModifDefColl(DBModiftype).Count = 0 Then
                    Globals.nonInteractive = False
                    Return "No DBModifier contained in Workbook at all!"
                End If
                Dim DBModifavailable As String = ""
                For Each DBMtype As String In {"DBMapper", "DBAction", "DBSeqnce"}
                    For Each DBMkey As String In Globals.DBModifDefColl(DBMtype).Keys
                        DBModifavailable += "," + DBMkey
                    Next
                Next
                Globals.nonInteractive = False
                Return "DB Modifier '" + DBModifName + "' not existing, available: " + DBModifavailable
            End If
            Globals.LogInfo("Doing DBModifier '" + DBModifName + "' ...")
            Try
                Globals.DBModifDefColl(DBModiftype).Item(DBModifName).doDBModif()
            Catch ex As Exception
                Globals.nonInteractive = False
                Return "DB Modifier '" + DBModifName + "' doDBModif had following error(s): " + ex.Message
            End Try
            Globals.nonInteractive = False
            If hadError Then Return Globals.nonInteractiveErrMsgs
            Return "" ' no error, no message
        Else
            Globals.nonInteractive = False
            Return "No valid type (" + DBModiftype + ") in passed DB Modifier '" + DBModifName + "', DB Modifier name must start with 'DBSeqnce', 'DBMapper' Or 'DBAction' !"
        End If
    End Function

    ''' <summary>marks a row in a DBMapper for deletion, used as a ExcelCommand to have a keyboard shortcut</summary>
    <ExcelCommand(Name:="deleteRow", ShortCut:="^D")>
    Public Sub deleteRow()
        Dim targetName As String = getDBModifNameFromRange(ExcelDnaUtil.Application.Selection)
        If Left(targetName, 8) = "DBMapper" Then DirectCast(Globals.DBModifDefColl("DBMapper").Item(targetName), DBMapper).insertCUDMarks(ExcelDnaUtil.Application.Selection, True)
    End Sub

    ''' <summary>inserts a row in a DBMapper, used as a ExcelCommand to have a keyboard shortcut</summary>
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


Public Class CustomCommandBuilder

    ''' <summary>Creates Insert command with support for Autoincrement (Identity) fields</summary>
    ''' <returns>Command for inserting</returns>
    Public Overridable Function InsertCommand() As Common.DbCommand
        Throw New NotImplementedException()
    End Function

    ''' <summary>Creates Delete command</summary>
    ''' <returns>Command for deleting</returns>
    Public Overridable Function DeleteCommand() As Common.DbCommand
        Throw New NotImplementedException()
    End Function

    ''' <summary>Creates Update command</summary>
    ''' <returns>Command for updating</returns>
    Public Overridable Function UpdateCommand() As Common.DbCommand
        Throw New NotImplementedException()
    End Function

End Class

''' <summary>Custom Command builder for SQLServer to avoid primary key problems with builtin ones
''' derived (transposed into VB.NET) from https://www.cogin.com/articles/CustomCommandBuilder.php
''' Copyright by Dejan Grujic 2004. http://www.cogin.com
''' </summary>
Public Class CustomSqlCommandBuilder
    Inherits CustomCommandBuilder

    Private dataTable As DataTable
    Private connection As SqlClient.SqlConnection
    Private allColumns As DataColumn()
    Private schemaDataTypeCollection As Collection

    Public Sub New(dataTable As DataTable, connection As SqlClient.SqlConnection, allColumns As DataColumn(), schemaDataTypeCollection As Collection)
        Me.dataTable = dataTable
        Me.connection = connection
        Me.allColumns = allColumns
        Me.schemaDataTypeCollection = schemaDataTypeCollection
    End Sub

    ''' <summary>Creates Insert command with support for Autoincrement (Identity) fields</summary>
    ''' <returns>SqlCommand for inserting</returns>
    Public Overrides Function InsertCommand() As Common.DbCommand
        Dim command As SqlClient.SqlCommand = GetTextCommand("")
        Dim intoString As StringBuilder = New StringBuilder()
        Dim valuesString As StringBuilder = New StringBuilder()
        Dim autoincrementColumns As ArrayList = AutoincrementKeyColumns()
        For Each column As DataColumn In allColumns
            If Not autoincrementColumns.Contains(column) Then
                If (intoString.Length > 0) Then
                    intoString.Append(", ")
                    valuesString.Append(", ")
                End If
                intoString.Append(column.ColumnName)
                valuesString.Append("@").Append(column.ColumnName)
                command.Parameters.Add(CreateParam(column))
            End If
        Next
        Dim commandText As String = "INSERT INTO " + TableName() + "(" + intoString.ToString() + ") VALUES (" + valuesString.ToString() + "); "
        If autoincrementColumns.Count > 0 Then
            commandText += "SELECT SCOPE_IDENTITY() AS " + DirectCast(autoincrementColumns(0), DataColumn).ColumnName()
        End If
        command.CommandText = commandText
        Return command
    End Function

    Private Function AutoincrementKeyColumns() As ArrayList
        AutoincrementKeyColumns = New ArrayList
        For Each primaryKeyColumn As DataColumn In dataTable.PrimaryKey
            If primaryKeyColumn.AutoIncrement Then
                AutoincrementKeyColumns.Add(primaryKeyColumn)
            End If
        Next
    End Function


    ''' <summary>Creates Delete command</summary>
    ''' <returns>SqlCommand for deleting</returns>
    Public Overrides Function DeleteCommand() As Common.DbCommand
        Dim command As SqlClient.SqlCommand = GetTextCommand("")
        Dim whereString As StringBuilder = New StringBuilder()
        For Each column As DataColumn In dataTable.PrimaryKey
            If (whereString.Length > 0) Then
                whereString.Append(" AND ")
            End If
            whereString.Append(column.ColumnName).Append(" = @").Append(column.ColumnName)
            command.Parameters.Add(CreateParam(column))
        Next
        Dim commandText As String = "DELETE FROM " + TableName() + " WHERE " + whereString.ToString()
        command.CommandText = commandText
        Return command
    End Function

    ''' <summary>Creates Update command</summary>
    ''' <returns>SqlCommand for updating</returns>
    Public Overrides Function UpdateCommand() As Common.DbCommand
        Dim command As SqlClient.SqlCommand = GetTextCommand("")
        Dim setString As StringBuilder = New StringBuilder()
        Dim whereString As StringBuilder = New StringBuilder()

        Dim primaryKeyColumns() As DataColumn = dataTable.PrimaryKey
        For Each column As DataColumn In allColumns
            If (System.Array.IndexOf(primaryKeyColumns, column) <> -1) Then
                ' primary key
                If (whereString.Length > 0) Then
                    whereString.Append(" AND ")
                End If
                whereString.Append(column.ColumnName).Append("= @old").Append(column.ColumnName)
            Else
                If (setString.Length > 0) Then
                    setString.Append(", ")
                End If
                setString.Append(column.ColumnName).Append(" = @").Append(column.ColumnName)
            End If
            command.Parameters.Add(CreateParam(column))
            command.Parameters.Add(CreateOldParam(column))
        Next
        Dim commandText As String = "UPDATE " + TableName() + " SET " + setString.ToString() + " WHERE " + whereString.ToString()
        command.CommandText = commandText
        Return command
    End Function

    Private Function CreateOldParam(column As DataColumn) As SqlClient.SqlParameter
        Dim sqlParam As SqlClient.SqlParameter = New SqlClient.SqlParameter()
        Dim columnName As String = column.ColumnName
        sqlParam.ParameterName = "@old" + columnName
        sqlParam.SourceColumn = columnName
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam.DbType = TypeToDbType(column.DataType(), columnName, schemaDataTypeCollection)
        Return sqlParam
    End Function

    Private Function CreateParam(column As DataColumn) As SqlClient.SqlParameter
        Dim sqlParam As SqlClient.SqlParameter = New SqlClient.SqlParameter()
        Dim columnName As String = column.ColumnName
        sqlParam.ParameterName = "@" + columnName
        sqlParam.SourceColumn = columnName
        sqlParam.DbType = TypeToDbType(column.DataType(), columnName, schemaDataTypeCollection)
        Return sqlParam
    End Function

    Private Function GetTextCommand(text As String) As SqlClient.SqlCommand
        Dim command As SqlClient.SqlCommand = New SqlClient.SqlCommand With {
            .CommandType = CommandType.Text,
            .CommandText = text,
            .Connection = connection
        }
        Return command
    End Function

    Private Function TableName() As String
        Return "[" + dataTable.TableName + "]"
    End Function
End Class

''' <summary>Custom Command builder for ODBC to avoid primary key problems with builtin ones
''' derived (transposed into VB.NET) from https://www.cogin.com/articles/CustomCommandBuilder.php
''' Copyright by Dejan Grujic 2004. http://www.cogin.com
''' </summary>
Public Class CustomOdbcCommandBuilder
    Inherits CustomCommandBuilder

    Private dataTable As DataTable
    Private connection As Odbc.OdbcConnection
    Private allColumns As DataColumn()
    Private schemaDataTypeCollection As Collection

    Public Sub New(dataTable As DataTable, connection As Odbc.OdbcConnection, allColumns As DataColumn(), schemaDataTypeCollection As Collection)
        Me.dataTable = dataTable
        Me.connection = connection
        Me.allColumns = allColumns
        Me.schemaDataTypeCollection = schemaDataTypeCollection
    End Sub

    ''' <summary>Creates Insert command with support for Autoincrement (Identity) fields</summary>
    ''' <returns>OdbcCommand for inserting</returns>
    Public Overrides Function InsertCommand() As Common.DbCommand
        Dim command As Odbc.OdbcCommand = GetTextCommand("")
        Dim intoString As StringBuilder = New StringBuilder()
        Dim valuesString As StringBuilder = New StringBuilder()
        Dim autoincrementColumns As ArrayList = AutoincrementKeyColumns()
        For Each column As DataColumn In allColumns
            If Not autoincrementColumns.Contains(column) Then
                If (intoString.Length > 0) Then
                    intoString.Append(", ")
                    valuesString.Append(", ")
                End If
                intoString.Append(column.ColumnName)
                valuesString.Append("@").Append(column.ColumnName)
                command.Parameters.Add(CreateParam(column))
            End If
        Next
        Dim commandText As String = "INSERT INTO " + TableName() + "(" + intoString.ToString() + ") VALUES (" + valuesString.ToString() + "); "
        If autoincrementColumns.Count > 0 Then
            commandText += "SELECT SCOPE_IDENTITY() AS " + DirectCast(autoincrementColumns(0), DataColumn).ColumnName()
        End If
        command.CommandText = commandText
        Return command
    End Function

    Private Function AutoincrementKeyColumns() As ArrayList
        AutoincrementKeyColumns = New ArrayList
        For Each primaryKeyColumn As DataColumn In dataTable.PrimaryKey
            If primaryKeyColumn.AutoIncrement Then
                AutoincrementKeyColumns.Add(primaryKeyColumn)
            End If
        Next
    End Function

    ''' <summary>Creates Delete command</summary>
    ''' <returns>OdbcCommand for deleting</returns>
    Public Overrides Function DeleteCommand() As Common.DbCommand
        Dim command As Odbc.OdbcCommand = GetTextCommand("")
        Dim whereString As StringBuilder = New StringBuilder()
        For Each column As DataColumn In dataTable.PrimaryKey
            If (whereString.Length > 0) Then
                whereString.Append(" AND ")
            End If
            whereString.Append(column.ColumnName).Append(" = @").Append(column.ColumnName)
            command.Parameters.Add(CreateParam(column))
        Next
        Dim commandText As String = "DELETE FROM " + TableName() + " WHERE " + whereString.ToString()
        command.CommandText = commandText
        Return command
    End Function

    ''' <summary>Creates Update command</summary>
    ''' <returns>OdbcCommand for updating</returns>
    Public Overrides Function UpdateCommand() As Common.DbCommand
        Dim command As Odbc.OdbcCommand = GetTextCommand("")
        Dim setString As StringBuilder = New StringBuilder()
        Dim whereString As StringBuilder = New StringBuilder()

        Dim primaryKeyColumns() As DataColumn = dataTable.PrimaryKey
        For Each column As DataColumn In allColumns
            If (System.Array.IndexOf(primaryKeyColumns, column) <> -1) Then
                ' primary key
                If (whereString.Length > 0) Then
                    whereString.Append(" AND ")
                End If
                whereString.Append(column.ColumnName).Append("= @old").Append(column.ColumnName)
            Else
                If (setString.Length > 0) Then
                    setString.Append(", ")
                End If
                setString.Append(column.ColumnName).Append(" = @").Append(column.ColumnName)
            End If
            command.Parameters.Add(CreateParam(column))
            command.Parameters.Add(CreateOldParam(column))
        Next
        Dim commandText As String = "UPDATE " + TableName() + " SET " + setString.ToString() + " WHERE " + whereString.ToString()
        command.CommandText = commandText
        Return command
    End Function

    Private Function CreateOldParam(column As DataColumn) As Odbc.OdbcParameter
        Dim sqlParam As Odbc.OdbcParameter = New Odbc.OdbcParameter()
        Dim columnName As String = column.ColumnName
        sqlParam.ParameterName = "@old" + columnName
        sqlParam.SourceColumn = columnName
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam.DbType = TypeToDbType(column.DataType(), columnName, schemaDataTypeCollection)
        Return sqlParam
    End Function

    Private Function CreateParam(column As DataColumn) As Odbc.OdbcParameter
        Dim sqlParam As Odbc.OdbcParameter = New Odbc.OdbcParameter()
        Dim columnName As String = column.ColumnName
        sqlParam.ParameterName = "@" + columnName
        sqlParam.SourceColumn = columnName
        sqlParam.DbType = TypeToDbType(column.DataType(), columnName, schemaDataTypeCollection)
        Return sqlParam
    End Function

    Private Function GetTextCommand(text As String) As Odbc.OdbcCommand
        Dim command As Odbc.OdbcCommand = New Odbc.OdbcCommand With {
            .CommandType = CommandType.Text,
            .CommandText = text,
            .Connection = connection
        }
        Return command
    End Function

    Private Function TableName() As String
        Return "[" + dataTable.TableName + "]"
    End Function
End Class

''' <summary>Custom Command builder for OleDB to avoid primary key problems with builtin ones
''' derived (transposed into VB.NET) from https://www.cogin.com/articles/CustomCommandBuilder.php
''' Copyright by Dejan Grujic 2004. http://www.cogin.com
''' </summary>
Public Class CustomOleDbCommandBuilder
    Inherits CustomCommandBuilder

    Private dataTable As DataTable
    Private connection As OleDb.OleDbConnection
    Private allColumns As DataColumn()
    Private schemaDataTypeCollection As Collection

    Public Sub New(dataTable As DataTable, connection As OleDb.OleDbConnection, allColumns As DataColumn(), schemaDataTypeCollection As Collection)
        Me.dataTable = dataTable
        Me.connection = connection
        Me.allColumns = allColumns
        Me.schemaDataTypeCollection = schemaDataTypeCollection
    End Sub

    ''' <summary>Creates Insert command with support for Autoincrement (Identity) fields</summary>
    ''' <returns>OleDbCommand for inserting</returns>
    Public Overrides Function InsertCommand() As Common.DbCommand
        Dim command As OleDb.OleDbCommand = GetTextCommand("")
        Dim intoString As StringBuilder = New StringBuilder()
        Dim valuesString As StringBuilder = New StringBuilder()
        Dim autoincrementColumns As ArrayList = AutoincrementKeyColumns()
        For Each column As DataColumn In allColumns
            If Not autoincrementColumns.Contains(column) Then
                If (intoString.Length > 0) Then
                    intoString.Append(", ")
                    valuesString.Append(", ")
                End If
                intoString.Append(column.ColumnName)
                valuesString.Append("@").Append(column.ColumnName)
                command.Parameters.Add(CreateParam(column))
            End If
        Next
        Dim commandText As String = "INSERT INTO " + TableName() + "(" + intoString.ToString() + ") VALUES (" + valuesString.ToString() + "); "
        If autoincrementColumns.Count > 0 Then
            commandText += "SELECT SCOPE_IDENTITY() AS " + DirectCast(autoincrementColumns(0), DataColumn).ColumnName()
        End If
        command.CommandText = commandText
        Return command
    End Function

    Private Function AutoincrementKeyColumns() As ArrayList
        AutoincrementKeyColumns = New ArrayList
        For Each primaryKeyColumn As DataColumn In dataTable.PrimaryKey
            If primaryKeyColumn.AutoIncrement Then
                AutoincrementKeyColumns.Add(primaryKeyColumn)
            End If
        Next
    End Function

    ''' <summary>Creates Delete command</summary>
    ''' <returns>OleDbCommand for deleting</returns>
    Public Overrides Function DeleteCommand() As Common.DbCommand
        Dim command As OleDb.OleDbCommand = GetTextCommand("")
        Dim whereString As StringBuilder = New StringBuilder()
        For Each column As DataColumn In dataTable.PrimaryKey
            If (whereString.Length > 0) Then
                whereString.Append(" AND ")
            End If
            whereString.Append(column.ColumnName).Append(" = @").Append(column.ColumnName)
            command.Parameters.Add(CreateParam(column))
        Next
        Dim commandText As String = "DELETE FROM " + TableName() + " WHERE " + whereString.ToString()
        command.CommandText = commandText
        Return command
    End Function

    ''' <summary>Creates Update command</summary>
    ''' <returns>OleDbCommand for updating</returns>
    Public Overrides Function UpdateCommand() As Common.DbCommand
        Dim command As OleDb.OleDbCommand = GetTextCommand("")
        Dim setString As StringBuilder = New StringBuilder()
        Dim whereString As StringBuilder = New StringBuilder()

        Dim primaryKeyColumns() As DataColumn = dataTable.PrimaryKey
        For Each column As DataColumn In allColumns
            If (System.Array.IndexOf(primaryKeyColumns, column) <> -1) Then
                ' primary key
                If (whereString.Length > 0) Then
                    whereString.Append(" AND ")
                End If
                whereString.Append(column.ColumnName).Append("= @old").Append(column.ColumnName)
            Else
                If (setString.Length > 0) Then
                    setString.Append(", ")
                End If
                setString.Append(column.ColumnName).Append(" = @").Append(column.ColumnName)
            End If
            command.Parameters.Add(CreateParam(column))
            command.Parameters.Add(CreateOldParam(column))
        Next
        Dim commandText As String = "UPDATE " + TableName() + " SET " + setString.ToString() + " WHERE " + whereString.ToString()
        command.CommandText = commandText
        Return command
    End Function

    Private Function CreateOldParam(column As DataColumn) As OleDb.OleDbParameter
        Dim sqlParam As OleDb.OleDbParameter = New OleDb.OleDbParameter()
        Dim columnName As String = column.ColumnName
        sqlParam.ParameterName = "@old" + columnName
        sqlParam.SourceColumn = columnName
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam.DbType = TypeToDbType(column.DataType(), columnName, schemaDataTypeCollection)
        Return sqlParam
    End Function

    Private Function CreateParam(column As DataColumn) As OleDb.OleDbParameter
        Dim sqlParam As OleDb.OleDbParameter = New OleDb.OleDbParameter()
        Dim columnName As String = column.ColumnName
        sqlParam.ParameterName = "@" + columnName
        sqlParam.SourceColumn = columnName
        sqlParam.DbType = TypeToDbType(column.DataType(), columnName, schemaDataTypeCollection)
        Return sqlParam
    End Function

    Private Function GetTextCommand(text As String) As OleDb.OleDbCommand
        Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand With {
            .CommandType = CommandType.Text,
            .CommandText = text,
            .Connection = connection
        }
        Return command
    End Function

    Private Function TableName() As String
        Return "[" + dataTable.TableName + "]"
    End Function
End Class
