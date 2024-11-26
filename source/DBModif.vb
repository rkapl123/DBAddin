Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Core
Imports System.Collections
Imports System.Data
Imports System.Data.Common
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Data.SqlClient
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
    ''' <summary>ask for confirmation before execution of DBModif</summary>
    Protected askBeforeExecute As Boolean = True
    ''' <summary>environment specific for the DBModif object, if left empty then set to default environment (either 0 or currently selected environment)</summary>
    Protected env As String = ""
    ''' <summary>Text displayed for confirmation before doing dbModif instead of standard text</summary>
    Protected confirmText As String = ""

    Public Sub New(definitionXML As CustomXMLNode)
        ' allow empty definition for DBModifDummy
        If TypeName(Me) = "DBModifDummy" Then Exit Sub
        Try
            If definitionXML.Attributes.Count > 0 Then
                dbmodifName = definitionXML.BaseName + definitionXML.Attributes(1).Text
            Else
                dbmodifName = definitionXML.BaseName
            End If
        Catch ex As Exception
            UserMsg("error in creating DBmodifier " + TypeName(Me) + " with passed definitionXML: " + ex.Message)
        End Try
        execOnSave = Convert.ToBoolean(getParamFromXML(definitionXML, "execOnSave", "Boolean"))
        askBeforeExecute = Convert.ToBoolean(getParamFromXML(definitionXML, "askBeforeExecute", "Boolean"))
        confirmText = getParamFromXML(definitionXML, "confirmText")
    End Sub
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
        If Not IsNothing(TargetRange) Then Marshal.ReleaseComObject(TargetRange)
    End Sub

    ''' <summary>asks user the confirmation question, in case it is required by the DB Modifier</summary>
    ''' <returns>Yes, No or Cancel (only possible when saving to finish questions)</returns>
    Public Function confirmExecution(Optional WbIsSaving As Boolean = False) As MsgBoxResult
        ' when saving always ask for DBModifiers that are to be executed, otherwise ask only if required....
        If (WbIsSaving And execOnSave) Or (Not WbIsSaving And askBeforeExecute) Then
            Dim executeTitle As String = "Execute " + TypeName(Me) + IIf(WbIsSaving, " on save", "")
            ' special confirm text or standard text?
            Dim executeMessage As String = IIf(confirmText <> "", confirmText, "Really execute " + IIf(Strings.LCase(dbmodifName) = Strings.LCase(TypeName(Me)), "Unnamed " + TypeName(Me), dbmodifName) + "?")
            ' for confirmation on Workbook saving provide cancel possibility as shortcut.
            Dim setQuestionType As MsgBoxStyle = IIf(WbIsSaving, MsgBoxStyle.YesNoCancel, MsgBoxStyle.YesNo)
            ' only ask if set to ask...
            If askBeforeExecute Then Return QuestionMsg(theMessage:=executeMessage, questionType:=setQuestionType, questionTitle:=executeTitle)
            Return MsgBoxResult.Yes
        Else
            Return MsgBoxResult.No
        End If
    End Function

    ''' <summary>public accessor function</summary>
    ''' <returns>the targetRangeAddress nicely formatted</returns>
    Public Function getTargetRangeAddress() As String
        getTargetRangeAddress = ""
        Dim actWbNames As Excel.Names
        Try : actWbNames = ExcelDnaUtil.Application.ActiveWorkbook.Names : Catch ex As Exception
            LogWarn("Exception when trying to get the active workbook names for getting Target Range Address: " + ex.Message + ", this might be either due to errors in the VBA-IDE (missing references) or due to opening this workbook from an MS-Office hyperlink, starting up Excel (timing issue). Switch to another workbook and back to fix.")
            Exit Function
        End Try
        If TypeName(Me) <> "DBSeqnce" Then
            Dim addRefersToFormula As String = ""
            If InStr(1, actWbNames.Item(dbmodifName).RefersTo, "=OFFSET(") > 0 Then
                addRefersToFormula = " (" + actWbNames.Item(dbmodifName).RefersTo
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

    ''' <summary>public accessor function: get Environment (integer) where connection id should be taken from</summary>
    ''' <returns>the Environment of the DBModif, 0 to indicate a not set environment</returns>
    Protected Function getEnv() As Integer
        getEnv = 0
        If TypeName(Me) = "DBSeqnce" Then Throw New NotImplementedException()
        ' set environment on DBModif overrides selected environment. Could be empty or 0 to indicate a not set environment
        If env <> "" AndAlso env <> "0" Then getEnv = Convert.ToInt16(env)
    End Function

    ''' <summary>checks whether DBModifier needs saving, usually because execOnSave is set (in case of CUD DBMappers if any i/u/d flags are present)</summary>
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
    ''' <param name="nodeName">the definition element's name (e.g "env")</param>
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
            If setEnv = 0 Then setEnv = selectedEnvironment + 1
        Else
            ' if Environment is not existing (default environment = 0), take from environment of enclosing sequence 
            If setEnv = 0 Then
                setEnv = CInt(DBSequenceEnv)
            Else
                ' otherwise check if set environment is different from sequence environment and set to sequence environment after warning message.
                If setEnv <> CInt(DBSequenceEnv) Then
                    UserMsg("Environment in " + TypeName(Me) + " (" + env + ") different than given environment of DB Sequence (" + DBSequenceEnv + "), overruling with sequence environment !")
                    setEnv = CInt(DBSequenceEnv)
                End If
            End If
        End If
        openDatabase = openIdbConnection(setEnv, database)
    End Function

    ''' <summary>refresh a DB Function (DBListFetch, DBRowFetch and DBSetQuery) by invoking its respective DB*Action procedure (the UDFs cannot be directly invoked from VB.NET code)
    ''' additionally preparing the inputs to the DB*Action procedure by extracting them from the DB Functions parameters</summary>
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
            UserMsg("Didn't find caller cell of DBRefresh using srcExtent " + srcExtent + "!", "Refresh of DB Functions")
            Exit Function
        End Try
        If caller.Parent.ProtectContents Then
            UserMsg("Worksheet " + caller.Parent.Name + " is content protected, can't refresh DB Function !")
            Exit Function
        End If
        Dim targetExtent = Replace(srcExtent, "DBFsource", "DBFtarget")
        Dim target As Excel.Range
        Try : target = ExcelDnaUtil.Application.Range(targetExtent)
        Catch ex As Exception
            UserMsg("Didn't find target of DBRefresh using targetExtent " + targetExtent + "!", "Refresh of DB Functions")
            Exit Function
        End Try

        If target.Parent.ProtectContents Then
            UserMsg("Worksheet " + target.Parent.Name + " is content protected, can't refresh DB Function !")
            Exit Function
        End If
        Dim DBMapperUnderlying As String = getDBModifNameFromRange(target)
        Dim targetExtentF = Replace(srcExtent, "DBFsource", "DBFtargetF")
        Dim formulaRange As Excel.Range = Nothing
        ' formulaRange might not exist
        Try : formulaRange = ExcelDnaUtil.Application.Range(targetExtentF) : Catch ex As Exception : End Try
        If formulaRange IsNot Nothing AndAlso formulaRange.Parent.ProtectContents Then
            UserMsg("Worksheet " + formulaRange.Parent.Name + " is content protected, can't refresh DB Function !")
            Exit Function
        End If
        Dim DBMapperUnderlyingF As String = getDBModifNameFromRange(formulaRange)

        ' allow for avoidance of overwriting users changes with CUDFlags after a data error occurred
        If hadError Then
            If executedDBMappers.ContainsKey(DBMapperUnderlying) Then
                Dim retval = QuestionMsg(theMessage:="Error(s) occurred during sequence, really refresh Target-range? This could lead to loss of entries.", questionTitle:="Refresh of DB Functions in DB Sequence")
                If retval = vbCancel Then Exit Function
            End If
        End If
        ' prevent deadlock if we are in a transaction and want to refresh the area that was changed.
        If (modifiedDBMappers.ContainsKey(DBMapperUnderlying) Or modifiedDBMappers.ContainsKey(DBMapperUnderlyingF)) And TransactionIsOpen Then
            UserMsg("Transaction affecting the target area is still open, refreshing it would result in a deadlock on the database table. Skipping refresh, consider placing refresh outside of transaction.", "Refresh of DB Functions in DB Sequence")
            Exit Function
        End If

        ' reset query cache, so we really get new data !
        Dim callID As String
        Try
            ' get the callID of the underlying name of the target (key of the queryCache and StatusCollection)
            callID = "[" + caller.Parent.Parent.Name + "]" + caller.Parent.Name + "!" + caller.Address
        Catch ex As Exception
            UserMsg("Didn't find target of DBRefresh !", "Refresh of DB Functions")
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
            UserMsg("Exception: " + ex.Message, "DBRefresh")
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
        resolveConnstring(ConnString, EnvPrefix, getConnStrForDBSet)
        getConnString = CStr(ConnString)
    End Function

End Class

''' <summary>DBMappers are used to store a range of data in Excel to the database. A special type of DBMapper is the CUD DBMapper for realizing the former DBSheets</summary>
Public Class DBMapper : Inherits DBModif

    ''' <summary>Database Table, where Data is to be stored</summary>
    Private ReadOnly tableName As String = ""
    ''' <summary>count of primary keys in data-table, starting from the leftmost column</summary>
    Private ReadOnly primKeysCount As Integer = 0
    ''' <summary>if set, then insert row into table if primary key is missing there. Default = False (only update)</summary>
    Private ReadOnly insertIfMissing As Boolean = False
    ''' <summary>additional stored procedure to be executed after saving</summary>
    Private ReadOnly executeAdditionalProc As String = ""
    ''' <summary>columns to be ignored (helper columns)</summary>
    Private ReadOnly ignoreColumns As String = ""
    ''' <summary>respect C/U/D Flags (DBSheet functionality)</summary>
    Public CUDFlags As Boolean = False
    ''' <summary>if set, don't notify error values in cells during update/insert</summary>
    Private ReadOnly IgnoreDataErrors As Boolean = False
    ''' <summary>used to pass whether changes were done</summary>
    Private changesDone As Boolean = False
    '''<summary>first column is treated as an auto-incrementing key column, needed to ignore empty values there (otherwise DBMappers stop here)</summary>
    Private ReadOnly AutoIncFlag As Boolean = False
    ''' <summary>prevent filling of whole table during execution of DB Mappers, this is useful for very large tables that are incrementally filled and would take unnecessary long time to start the DB Mapper. If set to true then each record is searched independently by going to the database. If the records to be stored are not too many, then this is more efficient than loading a very large table.</summary>
    Private ReadOnly avoidFill As Boolean = False
    ''' <summary>flag to prevent automatic resizing of columns</summary>
    Private ReadOnly preventColResize As Boolean = False
    ''' <summary>if set, delete table before inserting the contents of DBMapper</summary>
    Private ReadOnly deleteBeforeMapperInsert As Boolean = False
    ''' <summary>contains cells in DBSheetLookups with lookups that should only be refreshed after DB modification was done. If empty, all lookups are refreshed for that DB Modifier/Sheet</summary>
    Private ReadOnly onlyRefreshTheseDBSheetLookups As String

    ''' <summary>constructor with definition XML</summary>
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
            If tableName = "" Then Throw New Exception("No Table-name given in DBMapper definition!")
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
            ' from here on, all properties are set only in XML...
            avoidFill = Convert.ToBoolean(getParamFromXML(definitionXML, "avoidFill", "Boolean"))
            preventColResize = Convert.ToBoolean(getParamFromXML(definitionXML, "preventColResize", "Boolean"))
            deleteBeforeMapperInsert = Convert.ToBoolean(getParamFromXML(definitionXML, "deleteBeforeMapperInsert", "Boolean"))
            onlyRefreshTheseDBSheetLookups = getParamFromXML(definitionXML, "onlyRefreshTheseDBSheetLookups")
            ' set table styles for DBMappers having a list-object underneath
            Dim DBmapperListObj As Excel.ListObject = Nothing
            Try : DBmapperListObj = TargetRange.ListObject : Catch ex As Exception : End Try
            If DBmapperListObj IsNot Nothing Then
                ' special gray table style for CUDFlags DBMapper
                If CUDFlags Then
                    DBmapperListObj.TableStyle = fetchSetting("DBMapperCUDFlagStyle", "TableStyleLight11")
                    ' otherwise blue
                Else
                    DBmapperListObj.TableStyle = fetchSetting("DBMapperStandardStyle", "TableStyleLight9")
                End If
            End If
            ' allow CUDFlags only on DBMappers with underlying List-objects that were created with a query
            If CUDFlags And (DBmapperListObj Is Nothing OrElse DBmapperListObj.SourceType <> Excel.XlListObjectSourceType.xlSrcQuery) Then
                CUDFlags = False
                ' remove and reset CUDFlags setting in the definitions
                definitionXML.SelectSingleNode("ns0:CUDFlags").Delete()
                definitionXML.AppendChildNode("CUDFlags", NamespaceURI:="DBModifDef", NodeValue:="False")
                Throw New Exception("CUDFlags only supported for DBMappers on ListObjects (created with DBSetQueryListObject)!")
            End If
        Catch ex As Exception
            UserMsg("Error when creating DBMapper '" + dbmodifName + "': " + ex.Message, "DBModifier Definitions Error")
        End Try
    End Sub

    ''' <summary>checks whether DBModifier needs saving, either because execOnSave is set or in case of CUD DBMappers if any i/u/d Flags are present</summary>
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
        Dim insertFlag As Boolean

        If Not CUDFlags Then Exit Sub
        Dim changedRangeRows As Integer = changedRange.Rows.Count
        Dim changedRangeColumns As Integer = changedRange.Columns.Count
        Dim targetRangeRows As Integer = TargetRange.Rows.Count
        Dim targetRangeColumns As Integer = TargetRange.Columns.Count
        Dim sheetColumns As Integer = changedRange.Parent.Columns.Count

        ' sanity check for single cell DB Mappers..
        If targetRangeColumns = 1 And targetRangeRows = 1 Then
            Dim retval As MsgBoxResult = QuestionMsg("DB Mapper Range with CUD Flags is only one cell, really set CUD Flags ?",, "Set CUD Flags for DB Mapper")
            If retval = vbCancel Then Exit Sub
        End If
        ' sanity check for whole column change (this happens when Ctrl-minus is pressed in the right area of the list-object)..
        If changedRangeColumns = 1 And targetRangeRows = changedRangeRows Then
            UserMsg("Whole column deleted, it is recommended to immediately close the DBSheet Workbook to avoid destroying the DBSheet!", "Set CUD Flags for DB Mapper")
            Exit Sub
        End If
        ' sanity check for whole range change (this happens when the table is auto-filled down by dragging while being INSIDE the table)..
        ' in this case excel extends the change to the whole table and additionally the dragged area...
        If targetRangeColumns = changedRangeColumns And targetRangeRows <= changedRangeRows Then
            Dim retval As MsgBoxResult = QuestionMsg("Change affects whole DB Mapper Range, this might lead to erroneous behaviour, really set CUD Flags ?",, "Set CUD Flags for DB Mapper")
            If retval = vbCancel Then Exit Sub
        End If
        If changedRangeRows > fetchSettingInt("maxRowCountCUD", "10000") Then
            If Not QuestionMsg("A large range was changed (" + changedRangeRows.ToString() + " > maxRowCountCUD:" + fetchSetting("maxRowCountCUD", "10000") + "), this will probably lead to CUD flag setting taking very long. Continue?",, "Set CUD Flags for DB Mapper") = MsgBoxResult.Ok Then Exit Sub
        End If
        If changedRange.Parent.ProtectContents Then
            UserMsg("Worksheet " + changedRange.Parent.Name + " is content protected, can't set CUD Flags !", "Set CUD Flags for DB Mapper")
            Exit Sub
        End If

        preventChangeWhileFetching = True
        ' DBMapper ranges relative to start of TargetRange and respecting a header row, so CUDMarkRow = changedRange.Row - TargetRange.Row + 1 ...
        Try
            If deleteFlag Then
                Dim countRow As Integer = 1
                ExcelDnaUtil.Application.AutoCorrect.AutoExpandListRange = False ' to prevent automatic creation of new column
                For Each changedRow As Excel.Range In changedRange.Rows
                    Dim CUDMarkRow As Integer = changedRow.Row - TargetRange.Row + 1
                    If TargetRange.Cells(CUDMarkRow, targetRangeColumns + 1).Value IsNot Nothing Then Continue For
                    TargetRange.Cells(CUDMarkRow, targetRangeColumns + 1).Value = "d"
                    TargetRange.Rows(CUDMarkRow).Font.Strikethrough = True
                    ExcelDnaUtil.Application.Statusbar = "Delete mark for row " + countRow.ToString() + "/" + changedRangeRows.ToString()
                    countRow += 1
                Next
                ExcelDnaUtil.Application.AutoCorrect.AutoExpandListRange = True
            Else
                ' empty DBMapper: data was inserted in an empty or first row, check if other cells (not inserted) are empty set insert-flag
                If changedRangeRows = 1 Then
                    insertFlag = True
                    For Each containedCell As Excel.Range In TargetRange.Rows(changedRange.Row - TargetRange.Row + 1).Cells
                        ' check without newly inserted/updated cells (copy paste) 
                        Dim possibleIntersection As Excel.Range = ExcelDnaUtil.Application.Intersect(containedCell, changedRange)
                        ' check if whole row is empty (except for the changedRange), formulas do not count as filled (automatically filled for lookups or other things)..
                        If containedCell.Value IsNot Nothing AndAlso possibleIntersection Is Nothing AndAlso Left(containedCell.Formula, 1) <> "=" Then
                            insertFlag = False
                            Exit For
                        End If
                    Next
                End If

                ' inside a list-object Ctrl & + and Ctrl & - add and remove a whole list-object range row, outside with selected row they add/remove a whole sheet row
                If (changedRangeColumns = targetRangeColumns Or changedRangeColumns = sheetColumns) And changedRangeRows = 1 Then
                    Dim CUDMarkRow As Integer = changedRange.Row - TargetRange.Row + 1
                    ' if all cells (especially first) are empty (=inserting a row with Ctrl & +) add insert flag
                    If IsNothing(TargetRange.Cells(CUDMarkRow, 1).Value) Then
                        insertFlag = True
                        ' additionally shift the CUD Markers down, except a whole sheet row was inserted (which did the shift already)
                        If Not changedRangeColumns = sheetColumns Then
                            TargetRange.Cells(CUDMarkRow, TargetRange.Columns.Count + 1).Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftDown)
                        End If
                    Else
                        ' probably deleted with Ctrl & -, warn user...
                        Dim retval As MsgBoxResult = QuestionMsg("A whole row was modified, in case you deleted a row with Ctrl & -, the row will not deleted in the database (use Ctrl-Shift-D instead, click cancel and refresh the DB-Sheet with Ctrl-Shift-R now)." + vbCrLf + "If you inserted a row, confirm the insertion by continuing now.",, "Set CUD Flags for DB Mapper", MsgBoxStyle.Exclamation)
                        If retval = MsgBoxResult.Cancel Then GoTo exitSub
                    End If
                End If

                Dim countRow As Integer = 1
                For Each changedRow As Excel.Range In changedRange.Rows
                    Dim CUDMarkRow As Integer = changedRow.Row - TargetRange.Row + 1
                    ' change only if not already set. Do this here as it is faster then...
                    If TargetRange.Cells(CUDMarkRow, targetRangeColumns + 1).Value IsNot Nothing Then Continue For
                    ' repair automatic copying down of patterns (non null columns) from header row in List-objects done by Excel
                    If CUDMarkRow = 2 Then
                        With TargetRange.Rows(2).Interior
                            .Pattern = Excel.XlPattern.xlPatternNone
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                    End If
                    ' check if row was added at the bottom add insert flag
                    If CUDMarkRow > TargetRange.Cells(targetRangeRows, targetRangeColumns).Row Then insertFlag = True
                    ' insert only if explicitly required by insertFlag (after Row insertion, first an update flag will be written by the change event which is overwritten by an explicit call to insertCUDMarks)
                    ExcelDnaUtil.Application.AutoCorrect.AutoExpandListRange = False ' to prevent automatic creation of new column
                    If Not insertFlag Then
                        TargetRange.Cells(CUDMarkRow, targetRangeColumns + 1).Value = "u"
                        TargetRange.Rows(CUDMarkRow).Font.Italic = True
                    Else
                        TargetRange.Cells(CUDMarkRow, targetRangeColumns + 1).Value = "i"
                    End If
                    ExcelDnaUtil.Application.AutoCorrect.AutoExpandListRange = True
                    ExcelDnaUtil.Application.Statusbar = "Create/Update mark for row " + countRow.ToString() + "/" + changedRangeRows.ToString()
                    countRow += 1
                Next
            End If
        Catch ex As Exception
            LogWarn("Exception in insertCUDMarks: " + ex.Message)
        End Try
exitSub:
        preventChangeWhileFetching = False
        ExcelDnaUtil.Application.Statusbar = False
    End Sub

    ''' <summary>extend DataRange to "whole" DBMApper area (first row (header/field names) to the right and first column (first primary key) down)</summary>
    Public Sub extendDataRange()
        Dim actWbNames As Excel.Names
        Try : actWbNames = ExcelDnaUtil.Application.ActiveWorkbook.Names : Catch ex As Exception
            UserMsg("Exception when trying to get the active workbook names for extending DBMapper Range: " + ex.Message + ", this might be either due to errors in the VBA-IDE (missing references) or due to opening this workbook from an MS-Office hyperlink, starting up Excel (timing issue). Switch to another workbook and back to fix.")
            Exit Sub
        End Try
        ' only extend like this if no CUD Flags or AutoIncFlag present (may have non existing first (primary) columns -> auto identity columns !)
        If Not (CUDFlags Or AutoIncFlag) Then
            If TargetRange.Cells(2, 1).Value Is Nothing Then Exit Sub ' only extend if there are multiple rows...
            Dim rowCount As Integer = TargetRange.Cells(1, 1).End(Excel.XlDirection.xlDown).Row - TargetRange.Cells(1, 1).Row + 1
            ' unfortunately the above method to find the column extent doesn't work with hidden columns, so count the filled cells directly...
            Dim colCount As Integer = 1
            While Not (TargetRange.Cells(1, colCount + 1).Value Is Nothing OrElse TargetRange.Cells(1, colCount + 1).Value.ToString() = "")
                colCount += 1
            End While
            ' if we don't want columns to automatically extend then reset to original column count
            If preventColResize Then colCount = TargetRange.Columns.Count
            preventChangeWhileFetching = True
            Try
                ' only if the referral is to a real range (not an offset formula !)
                If InStr(1, actWbNames.Item(paramTargetName).RefersTo, "=OFFSET(") = 0 Then
                    actWbNames.Item(paramTargetName).RefersTo = actWbNames.Item(paramTargetName).RefersToRange.Resize(rowCount, colCount)
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
                UserMsg("Worksheet " + TargetRange.Parent.Name + " is content protected, can't reset CUD Flags !")
                Exit Sub
            End If
            ExcelDnaUtil.Application.AutoCorrect.AutoExpandListRange = False ' to prevent automatic creation of new column
            If Not IsNothing(TargetRange.ListObject) AndAlso Not IsNothing(TargetRange.ListObject.DataBodyRange) Then
                TargetRange.ListObject.DataBodyRange.Columns(TargetRange.Columns.Count + 1).ClearContents
            End If
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
        definitionXML.AppendChildNode("preventColResize", NamespaceURI:="DBModifDef", NodeValue:=preventColResize.ToString())
        definitionXML.AppendChildNode("deleteBeforeMapperInsert", NamespaceURI:="DBModifDef", NodeValue:=deleteBeforeMapperInsert.ToString())
        definitionXML.AppendChildNode("onlyRefreshTheseDBSheetLookups", NamespaceURI:="DBModifDef", NodeValue:=onlyRefreshTheseDBSheetLookups)
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
            Dim maxMassChanges As Integer = fetchSettingInt("maxNumberMassChange", "30")
            Dim curWs As Excel.Worksheet = TargetRange.Parent ' this is necessary because using TargetRange directly in CountIf deletes the content of the CUD area !!
            Dim changesToBeDone As Integer = ExcelDnaUtil.Application.WorksheetFunction.CountIf(curWs.Range(TargetRange.Columns(TargetRange.Columns.Count + 1).Address), "<>")
            If changesToBeDone > maxMassChanges Then
                Dim retval As MsgBoxResult = QuestionMsg(theMessage:="Modifying more rows (" + changesToBeDone.ToString() + ") than defined warning limit (" + maxMassChanges.ToString() + "), continue?", questionTitle:="Execute DB Mapper")
                If retval = vbCancel Then Exit Sub
            End If
        End If
        'now create/get a connection (dbcnn) for environment in case it was not already created by a step in the sequence before (transactions!)
        If Not TransactionOpen Then
            ExcelDnaUtil.Application.StatusBar = "opening database connection for " + database
            If Not openDatabase() Then Exit Sub
        End If

        If deleteBeforeMapperInsert Then
            Try
                Dim DmlCmd As IDbCommand = idbcnn.CreateCommand()
                If Not TransactionOpen Then DBModifs.trans = idbcnn.BeginTransaction()
                With DmlCmd
                    .Transaction = DBModifs.trans
                    .CommandText = "delete from " + tableName
                    .CommandTimeout = CmdTimeout
                    .CommandType = CommandType.Text
                    .ExecuteNonQuery()
                End With
            Catch ex As Exception
                notifyUserOfDataError("Error when deleting all data in " + tableName + ": " + ex.Message, 1)
                GoTo cleanup
            End Try
        End If

        ' set up data adapter and data set for checking DBMapper columns
        Dim da As DbDataAdapter = Nothing
        Dim dscheck As New DataSet()
        ExcelDnaUtil.Application.StatusBar = "initialising the Data Adapter"
        Try
            If TypeName(idbcnn) = "SqlConnection" Then
                ' decent behaviour for SQL Server
                Using comm As New SqlCommand("SET ARITHABORT ON", idbcnn, DBModifs.trans)
                    comm.ExecuteNonQuery()
                End Using
                da = New SqlDataAdapter(New SqlCommand("select * from " + tableName, idbcnn, DBModifs.trans)) With {
                    .UpdateBatchSize = 20
                }
            ElseIf TypeName(idbcnn) = "OleDbConnection" Then
                da = New OleDbDataAdapter(New OleDbCommand("select * from " + tableName, idbcnn, DBModifs.trans)) With {
                    .UpdateBatchSize = 20
                }
            Else
                da = New OdbcDataAdapter(New OdbcCommand("select * from " + tableName, idbcnn, DBModifs.trans))
            End If
            da.SelectCommand.CommandTimeout = CmdTimeout
        Catch ex As Exception
            UserMsg("Error in initializing Data Adapter for " + tableName + ": " + ex.Message, "DBMapper Error")
        End Try

        ExcelDnaUtil.Application.StatusBar = "retrieving the schema for " + tableName
        Try
            da.FillSchema(dscheck, SchemaType.Source, tableName)
        Catch ex As Exception
            UserMsg("Error in retrieving Schema for " + tableName + ": " + ex.Message, "DBMapper Error")
        End Try

        ' get the DataTypeName from the database if it exists, so we have a more accurate type information for the parameterized commands (select/insert/update/delete)
        Dim schemaReader As DbDataReader = Nothing
        Dim schemaDataTypeCollection As New Collection
        Try
            schemaReader = da.SelectCommand.ExecuteReader()
            For Each schemaRow As DataRow In schemaReader.GetSchemaTable().Rows
                Try : schemaDataTypeCollection.Add(schemaRow("DataTypeName"), schemaRow("ColumnName")) : Catch ex As Exception : End Try
            Next
            ' cancel command to finish data-reader (otherwise close takes very long until timeout)
            da.SelectCommand.Cancel()
            schemaReader.Close()
        Catch ex As Exception
            If Not IsNothing(schemaReader) Then
                da.SelectCommand.Cancel()
                schemaReader.Close()
            End If
        End Try

        ' first check if all column names (except ignored) of DBMapper Range exist in table and collect field-names
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
                    UserMsg("Field '" + fieldname + "' does not exist in Table '" + tableName + "' and is not in ignoreColumns, Error in sheet " + TargetRange.Parent.Name, "DBMapper Error")
                    GoTo cleanup
                Else
                    ' only add if field was not already added by the LU correction below!
                    If Not allColumnsStr.Contains(fieldname) Then
                        allColumnsStr(fieldColNum) = fieldname
                        fieldColNum += 1
                    End If
                End If
            ElseIf Strings.Right(fieldname.ToUpper(), 2) = "LU" And dscheck.Tables(0).Columns.Contains(Left(fieldname, Len(fieldname) - 2)) Then
                ' if ignored and the column is a lookup then add the column without LU to preserve the order of columns !!!
                fieldname = Left(fieldname, Len(fieldname) - 2) ' correct the LU to real field-name
                allColumnsStr(fieldColNum) = fieldname
                fieldColNum += 1
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
        Dim ds As New DataSet()
        ' openingQuote and closingQuote are needed to quote columns containing blanks
        Dim openingQuote As String = fetchSetting("openingQuote" + env.ToString(), "")
        Dim closingQuote As String = fetchSetting("closingQuote" + env.ToString(), "")
        ' replace the select command to avoid columns that are default filled but non null-able (will produce errors if not assigned to new row)
        da.SelectCommand.CommandText = "SELECT " + openingQuote + String.Join(closingQuote + "," + openingQuote, allColumnsStr) + closingQuote + " FROM " + tableName
        ' fill schema again to reflect the changed columns
        Try
            da.FillSchema(ds, SchemaType.Source, tableName)
        Catch ex As Exception
            notifyUserOfDataError("Error in getting schema information for " + tableName + " (" + da.SelectCommand.CommandText + "): " + ex.Message, 1)
            GoTo cleanup
        End Try

        ' for avoidFill mode (no uploading of whole table) replace select with parameterized query (parameters are added below)
        If avoidFill Then da.SelectCommand.CommandText = "SELECT " + openingQuote + String.Join(closingQuote + "," + openingQuote, allColumnsStr) + closingQuote + " FROM " + tableName + primKeyCompound

        Dim allColumns(UBound(allColumnsStr)) As DataColumn
        For i As Integer = 0 To UBound(allColumnsStr)
            allColumns(i) = ds.Tables(0).Columns(allColumnsStr(i))
        Next

        ' assign primary key columns from DBMapper table, as there might be none defined on the table
        Dim primKeyColumns(primKeysCount - 1) As DataColumn
        For i As Integer = 0 To UBound(primKeyColumnsStr)
            primKeyColumns(i) = ds.Tables(0).Columns(primKeyColumnsStr(i))

            ' for avoidFill mode (no uploading of whole table) set up parameters for parameterized query from primary keys
            If avoidFill Then
                Dim param As DbParameter
                If TypeName(idbcnn) = "SqlConnection" Then
                    param = DirectCast(da.SelectCommand, SqlCommand).CreateParameter()
                ElseIf TypeName(idbcnn) = "OleDbConnection" Then
                    param = DirectCast(da.SelectCommand, OleDbCommand).CreateParameter()
                Else
                    param = DirectCast(da.SelectCommand, OdbcCommand).CreateParameter()
                End If
                With param
                    .ParameterName = "@" + Replace(primKeyColumnsStr(i), " ", "SPACE")
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
                custCmdBuilder = New CustomSqlCommandBuilder(ds.Tables(0), idbcnn, allColumns, schemaDataTypeCollection, openingQuote, closingQuote)
            ElseIf TypeName(idbcnn) = "OleDbConnection" Then
                custCmdBuilder = New CustomOleDbCommandBuilder(ds.Tables(0), idbcnn, allColumns, schemaDataTypeCollection, openingQuote, closingQuote)
            Else
                custCmdBuilder = New CustomOdbcCommandBuilder(ds.Tables(0), idbcnn, allColumns, schemaDataTypeCollection, openingQuote, closingQuote)
            End If
            da.UpdateCommand = custCmdBuilder.UpdateCommand()
            da.UpdateCommand.CommandTimeout = CmdTimeout
            da.DeleteCommand = custCmdBuilder.DeleteCommand()
            da.DeleteCommand.CommandTimeout = CmdTimeout
            da.InsertCommand = custCmdBuilder.InsertCommand()
            da.InsertCommand.CommandTimeout = CmdTimeout
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
        ' walk through all rows in DBMapper Target-range to store in data set
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
                    ' special treatment for date(time) fields, try to convert from double (OLE Automation standard: Julian datetime values) if not properly formatted
                    If Left(ds.Tables(0).Columns(primKey).DataType.Name, 4) = "Date" Then
                        If TypeName(primKeyValue) = "Double" Then primKeyValue = Date.FromOADate(primKeyValue)
                    End If
                    ' empty primary keys are valid if primary key has auto-increment property defined, so pass DBNull Value here to avoid exception in finding record below (record is not found of course)...
                    primKeyValues(i - 1) = IIf(IsNothing(primKeyValue) OrElse primKeyValue.ToString() = "", DBNull.Value, primKeyValue)
                    If avoidFill Then
                        da.SelectCommand.Parameters.Item("@" + Replace(primKey, " ", "SPACE")).Value = primKeyValues(i - 1)
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

                ' get the record for updating, however avoid finding record with empty primary key value if auto-increment is given
                ' also avoid finding if no data should be in the table (deleteBeforeMapperInsert)
                Dim foundRow As DataRow = Nothing
                If Not AutoIncrement And Not deleteBeforeMapperInsert Then
                    Try
                        foundRow = ds.Tables(0).Rows.Find(primKeyValues)
                    Catch ex As Exception
                        If Not notifyUserOfDataError("Problem getting record, Error: " + ex.Message + " in sheet " + TargetRange.Parent.Name + ", row " + rowNum.ToString(), rowNum) Then GoTo cleanup
                        GoTo nextRow
                    End Try
                End If

                Dim insertRecord As Boolean = False
                ' If we have an auto-incrementing primary key (empty primary key value !) or didn't find a record with the given primary key (rst.EOF) ...
                If AutoIncrement OrElse IsNothing(foundRow) Then
                    If insertIfMissing Or rowCUDFlag = "i" Or deleteBeforeMapperInsert Then
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
                            ' special treatment for date(time) fields, try to convert from double (OLE Automation standard: Julian datetime values) if not properly formatted
                            If Left(ds.Tables(0).Columns(primKey).DataType.Name, 4) = "Date" Then
                                If TypeName(primKeyValue) = "Double" Then primKeyValue = Date.FromOADate(primKeyValue)
                            End If
                            Try
                                ' skip empty primary field values for auto-incrementing identity fields ..
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
                    ExcelDnaUtil.Application.StatusBar = Left("Inserting " + IIf(AutoIncrement, "new auto-incremented key", primKeyValueStr) + " into " + tableName, 255)
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
                                        ' special treatment for date(time) fields, try to convert from double (OLE Automation standard: Julian datetime values) if not properly formatted
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
                    storedProcCmd = New SqlCommand(executeAdditionalProc, idbcnn, trans)
                ElseIf TypeName(idbcnn) = "OleDbConnection" Then
                    storedProcCmd = New OleDbCommand(executeAdditionalProc, idbcnn, trans)
                Else
                    storedProcCmd = New OdbcCommand(executeAdditionalProc, idbcnn, trans)
                End If
                storedProcCmd.CommandText = executeAdditionalProc
                result = storedProcCmd.ExecuteNonQuery()
            Catch ex As Exception
                hadError = True
                UserMsg("Error in executing additional stored procedure: " + ex.Message, "DBMapper Error")
                GoTo cleanup
            End Try
            LogInfo("executed " + executeAdditionalProc + ", affected rows: " + result.ToString())
        End If
cleanup:
        ExcelDnaUtil.Application.StatusBar = False
        If deleteBeforeMapperInsert And Not TransactionOpen Then
            If hadError Then
                DBModifs.trans.Rollback()
            Else
                DBModifs.trans.Commit()
            End If
        End If
        ' close connection to return it to the pool (automatically closes recordset objects, so no need for checkrst.Close() or rst.Close())...
        If calledByDBSeq = "" Then idbcnn.Close()
        ' ask for refresh/clear CUD marks (only DBSheet) after DB Modification was done
        If changesDone Then
            Dim DBFunctionSrcExtent = getUnderlyingDBNameFromRange(TargetRange)
            If DBFunctionSrcExtent <> "" Then
                If CUDFlags Then
                    ' also resetCUDFlags for CUDFlags DBMapper (DBSheet) that do not ask before execute and were called by a DBSequence
                    Try
                        ' reset CUDFlags before refresh to avoid problems with reduced TargetRange due to deletions!
                        Me.resetCUDFlags()
                    Catch ex As Exception
                        UserMsg("Error in resetting CUD Flags: " + ex.Message, "DBMapper Error")
                    End Try
                    If calledByDBSeq = "" Then
                        Dim retval As MsgBoxResult
                        ' only ask when DBModifier not done on Workbook save, in this case refresh automatically...
                        If Not WbIsSaving Then retval = QuestionMsg(theMessage:="Refresh Data Range of DB Mapper '" + dbmodifName + "' and the DBSheetLookups?", questionTitle:="Refresh DB Mapper and Lookups")
                        If WbIsSaving Or retval = vbOK Or retval = vbNo Then
                            ' refresh underlying DB Function 
                            doDBRefresh(Replace(DBFunctionSrcExtent, "DBFtarget", "DBFsource"))
                            ' also refresh lookups
                            Dim lookupDefs As String() = {}
                            Try
                                For Each nm As Excel.Name In ExcelDnaUtil.Application.ActiveWorkbook.Names
                                    Dim rng As Excel.Range = Nothing
                                    Try : rng = nm.RefersToRange : Catch ex As Exception : End Try
                                    If rng IsNot Nothing Then
                                        If InStr(nm.Name, "DBFsource") > 0 Then
                                            Dim WbkSepPos As Integer = InStr(nm.Name, "!")
                                            ' only for DB source names on DBSheetLookups lookup sheet and onlyRefreshTheseDBSheetLookups matching their address / onlyRefreshTheseDBSheetLookups being empty (all lookups are refreshed)
                                            If InStr(nm.RefersTo, "DBSheetLookups!") > 0 And (onlyRefreshTheseDBSheetLookups = "" Or InStr(onlyRefreshTheseDBSheetLookups, "," + Replace(rng.Address, "$", "")) > 0) Then
                                                ReDim Preserve lookupDefs(lookupDefs.Length)
                                                If WbkSepPos > 1 Then
                                                    lookupDefs(lookupDefs.Length - 1) = Mid(nm.Name, WbkSepPos + 1)
                                                Else
                                                    lookupDefs(lookupDefs.Length - 1) = nm.Name
                                                End If
                                            End If
                                        End If
                                    End If
                                Next
                            Catch ex As Exception
                                UserMsg("Exception: " + ex.Message, "get underlying DBFName from Range")
                                lookupDefs = {}
                            End Try
                            If lookupDefs.Length() > 0 Then
                                For Each DBFsourceName As String In lookupDefs
                                    doDBRefresh(DBFsourceName)
                                Next
                            End If
                        End If
                    Else
                        LogWarn("no refresh took place for DBMapper " + dbmodifName + " because called by DBSequence " + calledByDBSeq)
                    End If
                End If
            End If
        End If
    End Sub

    ''' <summary>notification of error for user including selection of error cell</summary>
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
        Dim retval As MsgBoxResult = QuestionMsg(message + ", continue with storing or cancel?", MsgBoxStyle.OkCancel, "DBMapper Error", MsgBoxStyle.Critical)
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

    ''' <summary>is DBAction parametrized (placeholders to be filled with values from named ranges), default to false</summary>
    Private ReadOnly parametrized As Boolean
    ''' <summary>only for parametrized DBAction: enclosing character around parameter placeholders, defaults to !</summary>
    Private ReadOnly paramEnclosing As String
    ''' <summary>only for parametrized DBAction: comma separated locations of numerical parameters that should always be converted as strings</summary>
    Private ReadOnly convertAsString As String = ""
    ''' <summary>only for parametrized DBAction: comma separated locations of numerical parameters that should always be converted as date values (using the default DBDate formating)</summary>
    Private ReadOnly convertAsDate As String = ""
    ''' <summary>only for parametrized DBAction: if all values in the given Ranges are empty for one row, continue concatenation with all values being NULL, else finish at this row (excluding it), defaults to false</summary>
    Private ReadOnly continueIfRowEmpty As Boolean
    ''' <summary>only for parametrized DBAction: string of named ranges to be used as parameters that are replaced into the template string, where the order of the parameter range determines which placeholder is being replaced</summary>
    Private ReadOnly paramRangesStr As String = ""

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
            parametrized = Convert.ToBoolean(getParamFromXML(definitionXML, "parametrized", "Boolean"))
            paramEnclosing = getParamFromXML(definitionXML, "paramEnclosing")
            convertAsString = getParamFromXML(definitionXML, "convertAsString")
            convertAsDate = getParamFromXML(definitionXML, "convertAsDate")
            continueIfRowEmpty = Convert.ToBoolean(getParamFromXML(definitionXML, "continueIfRowEmpty", "Boolean"))
            paramRangesStr = getParamFromXML(definitionXML, "paramRangesStr")
            ' AFTER setting TargetRange and all the rest check for defined action to have a decent getTargetRangeAddress for undefined actions
            ' this has to come last as it throws an error, thus skipping the rest of the parameters...
            If paramTarget.Cells(1, 1).Text = "" Then Throw New Exception("No Action defined in " + paramTargetName + "(" + getTargetRangeAddress() + ")")
        Catch ex As Exception
            UserMsg("Error when creating DB Action '" + dbmodifName + "': " + ex.Message, "DBModifier Definitions Error")
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
                executeText += targetCell.Value2 + " "
            Next
            If parametrized Then
                result = executeTemplateSQL(executeText)
            Else
                result = executeSQL(executeText)
            End If
        Catch ex As Exception
            hadError = True
            UserMsg("Error in executing DB Action " + paramTargetName + ":" + vbCrLf + ex.Message, "DBAction Error")
            ExcelDnaUtil.Application.StatusBar = False
            Exit Sub
        End Try
        Dim message As String = "DBAction " + paramTargetName + " executed, affected records: " + result.ToString() + IIf(result = 0, ". As no records were affected, you might check whether you need to set the <continue if row empty> flag in case there should be some data affected", "")
        If Not WbIsSaving And calledByDBSeq = "" Then
            UserMsg(message, "DBAction executed", MsgBoxStyle.Information)
        Else
            LogInfo(message)
        End If
        ExcelDnaUtil.Application.StatusBar = False
        ' close connection to return it to the pool...
        If calledByDBSeq = "" Then idbcnn.Close()
    End Sub

    Private Function executeSQL(executeText As String) As Integer
        Dim DmlCmd As IDbCommand
        If TypeName(idbcnn) = "SqlConnection" Then
            DmlCmd = New SqlCommand(executeText, idbcnn, trans)
        ElseIf TypeName(idbcnn) = "OleDbConnection" Then
            DmlCmd = New OleDbCommand(executeText, idbcnn, trans)
        Else
            DmlCmd = New OdbcCommand(executeText, idbcnn, trans)
        End If
        DmlCmd.CommandTimeout = CmdTimeout
        DmlCmd.CommandType = CommandType.Text
        executeSQL = DmlCmd.ExecuteNonQuery()
    End Function

    ''' <summary>concatenate parameters into SQL template string through replacing the placeholders by the values given in paramRanges and execute the final SQL, empty values and excel errors get a NULL value.</summary>
    ''' <param name="paramString">SQL template string, the parameter placeholders are identified with !1!, !2!, ... (assuming ! is the default paramEnclosing)</param>
    Private Function executeTemplateSQL(paramString As String) As Integer
        executeTemplateSQL = 0

        ' first get parameter ranges and assign them into the paramRanges list
        Dim paramRanges As New List(Of Excel.Range)
        If paramRangesStr = "" Then Throw New Exception("No parameter range(s) given")
        For Each paramRange As String In Split(paramRangesStr, ",")
            paramRanges.Add(checkAndReturnRange(paramRange))
        Next

        ' then walk through the parameter ranges and convert the values, putting each set of parameters from one range into parameters which are accumulated in parameterList
        ' parameterAddress/parameterAddressList is used for better error handling messages later on.
        Dim refCount As Integer = 0
        Dim convAsString() As String = Split(convertAsString, ",")
        Dim convAsDate() As String = Split(convertAsDate, ",")
        Dim parameterList As New List(Of List(Of String))
        Dim parameterAddressList As New List(Of List(Of String))
        Dim maxCellCount As Integer = 0
        For Each paramRange As Excel.Range In paramRanges
            Dim parameters As New List(Of String)
            Dim parameterAddress As New List(Of String)
            refCount += 1
            Dim stringConversion As Boolean = convAsString.Contains(refCount.ToString())
            Dim dateConversion As Boolean = convAsDate.Contains(refCount.ToString())
            If stringConversion And dateConversion Then Throw New Exception("Both string and date conversion set for parameter range " + refCount.ToString() + ": convertAsString: " + convertAsString + ",convertAsDate: " + convertAsDate)
            Dim cellCount = 0
            ' walk through all cells of the parameter range, converting cell values and storing for later transposing
            For Each paramCell As Excel.Range In paramRange
                Try
                    If IsNothing(paramCell.Value) Or IsXLCVErr(paramCell.Value) Or Len(paramCell.Value) = 0 Then
                        parameters.Add("NULL")
                    ElseIf IsNumeric(paramCell.Value) Then
                        If stringConversion Then
                            parameters.Add("'" + Convert.ToString(paramCell.Value, System.Globalization.CultureInfo.InvariantCulture) + "'")
                        ElseIf dateConversion Then
                            parameters.Add(formatDBDate(paramCell.Value, DefaultDBDateFormatting))
                        Else
                            parameters.Add(Convert.ToString(paramCell.Value, System.Globalization.CultureInfo.InvariantCulture))
                        End If
                    ElseIf IsDate(paramCell.Value) Then
                        parameters.Add(formatDBDate(paramCell.Value2, DefaultDBDateFormatting))
                    Else
                        parameters.Add("'" + paramCell.Value + "'")
                    End If
                Catch ex As Exception
                    Throw New Exception("Converting parameters in row " + cellCount.ToString() + " of parameter " + refCount.ToString() + " exception: " + ex.Message)
                End Try
                parameterAddress.Add(IIf(paramCell.Parent.Name <> ExcelDnaUtil.Application.ActiveSheet.Name, paramCell.Parent.Name + "!", "") + paramCell.Address)
                cellCount += 1
            Next
            If maxCellCount = 0 Then
                maxCellCount = cellCount
            ElseIf maxCellCount <> cellCount Then
                Throw New Exception("Length (" + cellCount.ToString() + ") of parameter range " + refCount.ToString() + " not the same as the previous parameter range (" + maxCellCount.ToString() + "), all ranges need to be the same size")
            End If
            parameterList.Add(parameters)
            parameterAddressList.Add(parameterAddress)
        Next

        ' finally transpose the parameters and replace the placeholders with their values into the final rowSQL before executing each rowSQL ...
        For cellCount As Integer = 1 To maxCellCount
            Dim rowSQL As String = paramString
            Dim parameterAddresses As String = ""
            Dim rowFilled As Boolean = False
            For paramCount As Integer = 1 To refCount
                If parameterList(paramCount - 1).Item(cellCount - 1) <> "NULL" Then rowFilled = True
                rowSQL = rowSQL.Replace(IIf(paramEnclosing = "", "!", paramEnclosing) + paramCount.ToString() + IIf(paramEnclosing = "", "!", paramEnclosing), parameterList(paramCount - 1).Item(cellCount - 1))
                parameterAddresses += parameterAddressList(paramCount - 1).Item(cellCount - 1) + ","
            Next
            ' leave early if needed
            If Not continueIfRowEmpty And Not rowFilled Then Exit Function
            ' only execute if any field in row is filled
            If rowFilled Then
                Try
                    executeTemplateSQL += executeSQL(rowSQL)
                Catch ex As Exception
                    Throw New Exception(ex.Message + " occurred in " + parameterAddresses + vbCrLf + rowSQL)
                End Try
            End If
        Next
    End Function

    ''' <summary>set the fields in the DB Modifier Create Dialog with attributes of object</summary>
    ''' <param name="theDBModifCreateDlg"></param>
    Public Overrides Sub setDBModifCreateFields(ByRef theDBModifCreateDlg As DBModifCreate)
        With theDBModifCreateDlg
            .envSel.SelectedIndex = getEnv() - 1
            .TargetRangeAddress.Text = getTargetRangeAddress()
            .Database.Text = database
            .execOnSave.Checked = execOnSave
            .AskForExecute.Checked = askBeforeExecute
            .parametrized.Checked = parametrized
            .continueIfRowEmpty.Checked = continueIfRowEmpty
            .paramEnclosing.Text = paramEnclosing
            .paramRangesStr.Text = paramRangesStr
            .convertAsDate.Text = convertAsDate
            .convertAsString.Text = convertAsString
        End With
    End Sub
End Class

''' <summary>DBSequences are used to group DBMappers and DBActions and run them in sequence together with refreshing DBFunctions and executing them in transaction brackets</summary>
Public Class DBSeqnce : Inherits DBModif

    ''' <summary>sequence of DB Mappers, DB Actions and DB Refreshes being executed in this sequence</summary>
    Private ReadOnly sequenceParams() As String = {}

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
            UserMsg("Error when creating DB Sequence '" + dbmodifName + "': " + ex.Message, "DBModifier Definitions Error")
        End Try
    End Sub

    Public Overrides Sub doDBModif(Optional WbIsSaving As Boolean = False, Optional calledByDBSeq As String = "", Optional TransactionOpen As Boolean = False)
        Dim TransactionIsOpen As Boolean = False
        hadError = False
        ' warning against recursions (should not happen...)
        If calledByDBSeq <> "" Then
            UserMsg("DB Sequence '" + dbmodifName + "' is being called by another DB Sequence (" + calledByDBSeq + "), this should not occur as infinite recursions are possible !", "Execute DB Sequence")
            Exit Sub
        End If
        ' ask for saving only if a) is not done on WorkbookSave b) is set to ask and c) is in interactive mode
        If Not WbIsSaving And askBeforeExecute And Not nonInteractive Then
            If Not confirmExecution() = MsgBoxResult.Yes Then Exit Sub
        End If

        ' reset the db connection in any case to allow for new connections at DBBegin
        idbcnn = Nothing
        Dim executedDBMappers As New Dictionary(Of String, Boolean)
        Dim modifiedDBMappers As New Dictionary(Of String, Boolean)
        For i As Integer = 0 To UBound(sequenceParams)
            Dim definition() As String = Split(sequenceParams(i), ":")
            Dim DBModiftype As String = definition(0)
            Dim DBModifname As String = definition(1)
            Select Case DBModiftype
                Case "DBMapper", "DBAction"
                    If Not hadError Then
                        LogInfo(DBModifname + "... ")
                        DBModifDefColl(DBModiftype).Item(DBModifname).doDBModif(WbIsSaving, calledByDBSeq:=MyBase.dbmodifName, TransactionOpen:=TransactionIsOpen)
                        If DBModiftype = "DBMapper" Then
                            If DirectCast(DBModifDefColl("DBMapper").Item(DBModifname), DBMapper).CUDFlags Then executedDBMappers(DBModifname) = True
                            If DirectCast(DBModifDefColl("DBMapper").Item(DBModifname), DBMapper).hadChanges Then modifiedDBMappers(DBModifname) = True
                        End If
                    End If
                Case "DBBegin"
                    LogInfo("DBBeginTrans... ")
                    If idbcnn Is Nothing Then
                        ' take database connection properties from sequence step after DBBegin. (checked) requirement: all steps have the same connection in this case!
                        Dim nextdefinition() As String = Split(sequenceParams(i + 1), ":")
                        If Not DBModifDefColl(nextdefinition(0)).Item(nextdefinition(1)).openDatabase(env) Then Exit Sub
                    End If
                    DBModifs.trans = idbcnn.BeginTransaction()
                    TransactionIsOpen = True
                Case "DBCommitRollback"
                    If Not hadError Then
                        LogInfo("DBCommitTrans... ")
                        DBModifs.trans.Commit()
                    Else
                        LogInfo("DBRollbackTrans... ")
                        DBModifs.trans.Rollback()
                    End If
                    TransactionIsOpen = False
                Case Else
                    If Not hadError Then
                        If Left(DBModiftype, 8) = "Refresh " Then
                            doDBRefresh(srcExtent:=DBModifname, executedDBMappers:=executedDBMappers, modifiedDBMappers:=modifiedDBMappers, TransactionIsOpen:=TransactionIsOpen)
                        Else
                            UserMsg("Unknown type of sequence step '" + DBModiftype + "' being called in DB Sequence (" + calledByDBSeq + ") !", "Execute DB Sequence")
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

Public Class DBModifDummy : Inherits DBModif

    Public Sub New()
        MyBase.New(Nothing)
    End Sub

    Public Sub executeRefresh(srcExtent)
        doDBRefresh(srcExtent:=srcExtent)
    End Sub

End Class

''' <summary>global helper functions for DBModifiers</summary>
Public Module DBModifs
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
    ''' <returns></returns>
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

    ''' <summary>gets defined names for DBModifier (DBMapper/DBAction/DBSeqnce) invocation in the current workbook and updates Ribbon with it</summary>
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
            If RangeParts.Count = 2 And RangeParts(0) <> "" And RangeParts(1) <> "" Then
                Dim SheetName = Replace(RangeParts(0), "'", "") ' for sheet-names with blanks surrounding quotations are needed, remove them here
                Dim Address = RangeParts(1)
                Dim srcExtent As String = ""
                Try : srcExtent = getUnderlyingDBNameFromRange(ExcelDnaUtil.Application.Worksheets(SheetName).Range(Address)) : Catch ex As Exception : End Try
                If srcExtent = "" Then Return "No valid address found in " + DBModifName + " (Sheet-name: " + SheetName + ", Address: " + Address + ")"
                Dim aDBModifier As DBModifDummy = New DBModifDummy()
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
End Module


Public Class CustomCommandBuilder
    Protected ReadOnly dataTable As DataTable
    Protected ReadOnly allColumns As DataColumn()
    Protected ReadOnly openingQuote As String
    Protected ReadOnly closingQuote As String

    Public Sub New(dataTable As DataTable, allColumns As DataColumn(), openingQuote As String, closingQuote As String)
        Me.dataTable = dataTable
        Me.allColumns = allColumns
        Me.openingQuote = openingQuote
        Me.closingQuote = closingQuote
    End Sub

    ''' <summary>Creates Insert command with support for Auto-increment (Identity) fields</summary>
    ''' <returns>Command for inserting</returns>
    Public Overridable Function InsertCommand() As DbCommand
        Throw New NotImplementedException()
    End Function

    ''' <summary>Creates Delete command</summary>
    ''' <returns>Command for deleting</returns>
    Public Overridable Function DeleteCommand() As DbCommand
        Throw New NotImplementedException()
    End Function

    ''' <summary>Creates Update command</summary>
    ''' <returns>Command for updating</returns>
    Public Overridable Function UpdateCommand() As DbCommand
        Throw New NotImplementedException()
    End Function

    Protected Function TableName() As String
        ' each identifier needs to be quoted by itself
        Return Me.openingQuote + String.Join(Me.closingQuote + "." + Me.openingQuote, Strings.Split(dataTable.TableName, ".")) + Me.closingQuote
    End Function

End Class

''' <summary>Custom Command builder for SQLServer to avoid primary key problems with built-in ones
''' derived (transposed into VB.NET) from https://www.cogin.com/articles/CustomCommandBuilder.php
''' Copyright by Dejan_Grujic 2004. http://www.cogin.com
''' </summary>
Public Class CustomSqlCommandBuilder
    Inherits CustomCommandBuilder

    Private ReadOnly connection As SqlConnection
    Private ReadOnly schemaDataTypeCollection As Collection

    Public Sub New(dataTable As DataTable, connection As SqlConnection, allColumns As DataColumn(), schemaDataTypeCollection As Collection, openingQuote As String, closingQuote As String)
        MyBase.New(dataTable, allColumns, openingQuote, closingQuote)
        Me.connection = connection
        Me.schemaDataTypeCollection = schemaDataTypeCollection
    End Sub

    ''' <summary>Creates Insert command with support for Auto-increment (Identity) fields</summary>
    ''' <returns>SqlCommand for inserting</returns>
    Public Overrides Function InsertCommand() As DbCommand
        Dim command As SqlCommand = GetTextCommand("")
        Dim intoString As New StringBuilder()
        Dim valuesString As New StringBuilder()
        Dim autoincrementColumns As ArrayList = AutoincrementKeyColumns()
        For Each column As DataColumn In allColumns
            If Not autoincrementColumns.Contains(column) Then
                If (intoString.Length > 0) Then
                    intoString.Append(", ")
                    valuesString.Append(", ")
                End If
                intoString.Append(Me.openingQuote).Append(column.ColumnName).Append(Me.closingQuote)
                valuesString.Append("@").Append(Replace(column.ColumnName, " ", "SPACE"))
                command.Parameters.Add(CreateParam(column))
            End If
        Next
        Dim commandText As String = "INSERT INTO " + TableName() + "(" + intoString.ToString() + ") VALUES (" + valuesString.ToString() + "); "
        If autoincrementColumns.Count > 0 Then
            commandText += "SELECT SCOPE_IDENTITY() AS " + Me.openingQuote + DirectCast(autoincrementColumns(0), DataColumn).ColumnName() + Me.closingQuote
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
    Public Overrides Function DeleteCommand() As DbCommand
        Dim command As SqlCommand = GetTextCommand("")
        Dim whereString As New StringBuilder()
        For Each column As DataColumn In dataTable.PrimaryKey
            If (whereString.Length > 0) Then
                whereString.Append(" AND ")
            End If
            whereString.Append(Me.openingQuote).Append(column.ColumnName).Append(Me.closingQuote).Append(" = @").Append(Replace(column.ColumnName, " ", "SPACE"))
            command.Parameters.Add(CreateParam(column))
        Next
        Dim commandText As String = "DELETE FROM " + TableName() + " WHERE " + whereString.ToString()
        command.CommandText = commandText
        Return command
    End Function

    ''' <summary>Creates Update command</summary>
    ''' <returns>SqlCommand for updating</returns>
    Public Overrides Function UpdateCommand() As DbCommand
        Dim command As SqlCommand = GetTextCommand("")
        Dim setString As New StringBuilder()
        Dim whereString As New StringBuilder()

        Dim primaryKeyColumns() As DataColumn = dataTable.PrimaryKey
        For Each column As DataColumn In allColumns
            If (System.Array.IndexOf(primaryKeyColumns, column) <> -1) Then
                ' primary keys go into where part
                If (whereString.Length > 0) Then
                    whereString.Append(" AND ")
                End If
                whereString.Append(Me.openingQuote).Append(column.ColumnName).Append(Me.closingQuote).Append("= @old").Append(Replace(column.ColumnName, " ", "SPACE"))
            Else
                ' other fields go into set part
                If (setString.Length > 0) Then
                    setString.Append(", ")
                End If
                setString.Append(Me.openingQuote).Append(column.ColumnName).Append(Me.closingQuote).Append(" = @").Append(Replace(column.ColumnName, " ", "SPACE"))
            End If
            command.Parameters.Add(CreateParam(column))
            command.Parameters.Add(CreateOldParam(column))
        Next
        Dim commandText As String = "UPDATE " + TableName() + " SET " + setString.ToString() + " WHERE " + whereString.ToString()
        command.CommandText = commandText
        Return command
    End Function

    Private Function CreateOldParam(column As DataColumn) As SqlParameter
        Dim sqlParam As New SqlParameter()
        Dim columnName As String = column.ColumnName
        sqlParam.ParameterName = "@old" + Replace(columnName, " ", "SPACE")
        sqlParam.SourceColumn = columnName
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam.DbType = TypeToDbType(column.DataType(), columnName, schemaDataTypeCollection)
        Return sqlParam
    End Function

    Private Function CreateParam(column As DataColumn) As SqlParameter
        Dim sqlParam As New SqlParameter()
        Dim columnName As String = column.ColumnName
        sqlParam.ParameterName = "@" + Replace(columnName, " ", "SPACE")
        sqlParam.SourceColumn = columnName
        sqlParam.DbType = TypeToDbType(column.DataType(), columnName, schemaDataTypeCollection)
        Return sqlParam
    End Function

    Private Function GetTextCommand(text As String) As SqlCommand
        Dim command As New SqlCommand With {
            .CommandType = CommandType.Text,
            .CommandText = text,
            .Connection = connection
        }
        Return command
    End Function

End Class

''' <summary>Custom Command builder for ODBC to avoid primary key problems with built-in ones
''' derived (transposed into VB.NET) from https://www.cogin.com/articles/CustomCommandBuilder.php
''' Copyright by Dejan_Grujic 2004. http://www.cogin.com
''' </summary>
Public Class CustomOdbcCommandBuilder
    Inherits CustomCommandBuilder

    Private ReadOnly connection As OdbcConnection
    Private ReadOnly schemaDataTypeCollection As Collection

    Public Sub New(dataTable As DataTable, connection As OdbcConnection, allColumns As DataColumn(), schemaDataTypeCollection As Collection, openingQuote As String, closingQuote As String)
        MyBase.New(dataTable, allColumns, openingQuote, closingQuote)
        Me.connection = connection
        Me.schemaDataTypeCollection = schemaDataTypeCollection
    End Sub

    ''' <summary>Creates Insert command with support for Auto-increment (Identity) fields</summary>
    ''' <returns>OdbcCommand for inserting</returns>
    Public Overrides Function InsertCommand() As DbCommand
        Dim command As OdbcCommand = GetTextCommand("")
        Dim intoString As New StringBuilder()
        Dim valuesString As New StringBuilder()
        Dim autoincrementColumns As ArrayList = AutoincrementKeyColumns()
        For Each column As DataColumn In allColumns
            If Not autoincrementColumns.Contains(column) Then
                If (intoString.Length > 0) Then
                    intoString.Append(", ")
                    valuesString.Append(", ")
                End If
                intoString.Append(Me.openingQuote).Append(column.ColumnName).Append(Me.closingQuote)
                valuesString.Append("@").Append(Replace(column.ColumnName, " ", "SPACE"))
                command.Parameters.Add(CreateParam(column))
            End If
        Next
        Dim commandText As String = "INSERT INTO " + TableName() + "(" + intoString.ToString() + ") VALUES (" + valuesString.ToString() + "); "
        If autoincrementColumns.Count > 0 Then
            commandText += "SELECT SCOPE_IDENTITY() AS " + Me.openingQuote + DirectCast(autoincrementColumns(0), DataColumn).ColumnName() + Me.closingQuote
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
    Public Overrides Function DeleteCommand() As DbCommand
        Dim command As OdbcCommand = GetTextCommand("")
        Dim whereString As New StringBuilder()
        For Each column As DataColumn In dataTable.PrimaryKey
            If (whereString.Length > 0) Then
                whereString.Append(" AND ")
            End If
            whereString.Append(Me.openingQuote).Append(column.ColumnName).Append(Me.closingQuote).Append(" = @").Append(Replace(column.ColumnName, " ", "SPACE"))
            command.Parameters.Add(CreateParam(column))
        Next
        Dim commandText As String = "DELETE FROM " + TableName() + " WHERE " + whereString.ToString()
        command.CommandText = commandText
        Return command
    End Function

    ''' <summary>Creates Update command</summary>
    ''' <returns>OdbcCommand for updating</returns>
    Public Overrides Function UpdateCommand() As DbCommand
        Dim command As OdbcCommand = GetTextCommand("")
        Dim setString As New StringBuilder()
        Dim whereString As New StringBuilder()

        Dim primaryKeyColumns() As DataColumn = dataTable.PrimaryKey
        For Each column As DataColumn In allColumns
            If (System.Array.IndexOf(primaryKeyColumns, column) <> -1) Then
                ' primary keys into where part
                If (whereString.Length > 0) Then
                    whereString.Append(" AND ")
                End If
                whereString.Append(Me.openingQuote).Append(column.ColumnName).Append(Me.closingQuote).Append("= @old").Append(Replace(column.ColumnName, " ", "SPACE"))
            Else
                ' other fields go into set part
                If (setString.Length > 0) Then
                    setString.Append(", ")
                End If
                setString.Append(Me.openingQuote).Append(column.ColumnName).Append(Me.closingQuote).Append(" = @").Append(Replace(column.ColumnName, " ", "SPACE"))
            End If
            command.Parameters.Add(CreateParam(column))
            command.Parameters.Add(CreateOldParam(column))
        Next
        Dim commandText As String = "UPDATE " + TableName() + " SET " + setString.ToString() + " WHERE " + whereString.ToString()
        command.CommandText = commandText
        Return command
    End Function

    Private Function CreateOldParam(column As DataColumn) As OdbcParameter
        Dim sqlParam As New OdbcParameter()
        Dim columnName As String = column.ColumnName
        sqlParam.ParameterName = "@old" + Replace(columnName, " ", "SPACE")
        sqlParam.SourceColumn = columnName
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam.DbType = TypeToDbType(column.DataType(), columnName, schemaDataTypeCollection)
        Return sqlParam
    End Function

    Private Function CreateParam(column As DataColumn) As OdbcParameter
        Dim sqlParam As New OdbcParameter()
        Dim columnName As String = column.ColumnName
        sqlParam.ParameterName = "@" + Replace(columnName, " ", "SPACE")
        sqlParam.SourceColumn = columnName
        sqlParam.DbType = TypeToDbType(column.DataType(), columnName, schemaDataTypeCollection)
        Return sqlParam
    End Function

    Private Function GetTextCommand(text As String) As OdbcCommand
        Dim command As New OdbcCommand With {
            .CommandType = CommandType.Text,
            .CommandText = text,
            .Connection = connection
        }
        Return command
    End Function
End Class

''' <summary>Custom Command builder for OleDB to avoid primary key problems with built-in ones
''' derived (transposed into VB.NET) from https://www.cogin.com/articles/CustomCommandBuilder.php
''' Copyright by Dejan_Grujic 2004. http://www.cogin.com
''' </summary>
Public Class CustomOleDbCommandBuilder
    Inherits CustomCommandBuilder

    Private ReadOnly connection As OleDbConnection
    Private ReadOnly schemaDataTypeCollection As Collection

    Public Sub New(dataTable As DataTable, connection As OleDbConnection, allColumns As DataColumn(), schemaDataTypeCollection As Collection, openingQuote As String, closingQuote As String)
        MyBase.New(dataTable, allColumns, openingQuote, closingQuote)
        Me.connection = connection
        Me.schemaDataTypeCollection = schemaDataTypeCollection
    End Sub

    ''' <summary>Creates Insert command with support for Auto-increment (Identity) fields</summary>
    ''' <returns>OleDbCommand for inserting</returns>
    Public Overrides Function InsertCommand() As DbCommand
        Dim command As OleDbCommand = GetTextCommand("")
        Dim intoString As New StringBuilder()
        Dim valuesString As New StringBuilder()
        Dim autoincrementColumns As ArrayList = AutoincrementKeyColumns()
        For Each column As DataColumn In allColumns
            If Not autoincrementColumns.Contains(column) Then
                If (intoString.Length > 0) Then
                    intoString.Append(", ")
                    valuesString.Append(", ")
                End If
                intoString.Append(Me.openingQuote).Append(column.ColumnName).Append(Me.closingQuote)
                valuesString.Append("@").Append(Replace(column.ColumnName, " ", "SPACE"))
                command.Parameters.Add(CreateParam(column))
            End If
        Next
        Dim commandText As String = "INSERT INTO " + TableName() + "(" + intoString.ToString() + ") VALUES (" + valuesString.ToString() + "); "
        If autoincrementColumns.Count > 0 Then
            commandText += "SELECT SCOPE_IDENTITY() AS " + Me.openingQuote + DirectCast(autoincrementColumns(0), DataColumn).ColumnName() + Me.closingQuote
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
    Public Overrides Function DeleteCommand() As DbCommand
        Dim command As OleDbCommand = GetTextCommand("")
        Dim whereString As New StringBuilder()
        For Each column As DataColumn In dataTable.PrimaryKey
            If (whereString.Length > 0) Then
                whereString.Append(" AND ")
            End If
            whereString.Append(Me.openingQuote).Append(column.ColumnName).Append(Me.closingQuote).Append(" = @").Append(Replace(column.ColumnName, " ", "SPACE"))
            command.Parameters.Add(CreateParam(column))
        Next
        Dim commandText As String = "DELETE FROM " + TableName() + " WHERE " + whereString.ToString()
        command.CommandText = commandText
        Return command
    End Function

    ''' <summary>Creates Update command</summary>
    ''' <returns>OleDbCommand for updating</returns>
    Public Overrides Function UpdateCommand() As DbCommand
        Dim command As OleDbCommand = GetTextCommand("")
        Dim setString As New StringBuilder()
        Dim whereString As New StringBuilder()

        Dim primaryKeyColumns() As DataColumn = dataTable.PrimaryKey
        For Each column As DataColumn In allColumns
            If (System.Array.IndexOf(primaryKeyColumns, column) <> -1) Then
                ' primary key into where part
                If (whereString.Length > 0) Then
                    whereString.Append(" AND ")
                End If
                whereString.Append(Me.openingQuote).Append(column.ColumnName).Append(Me.closingQuote).Append("= @old").Append(Replace(column.ColumnName, " ", "SPACE"))
            Else
                ' other fields go into set part
                If (setString.Length > 0) Then
                    setString.Append(", ")
                End If
                setString.Append(Me.openingQuote).Append(column.ColumnName).Append(Me.closingQuote).Append(" = @").Append(Replace(column.ColumnName, " ", "SPACE"))
            End If
            command.Parameters.Add(CreateParam(column))
            command.Parameters.Add(CreateOldParam(column))
        Next
        Dim commandText As String = "UPDATE " + TableName() + " SET " + setString.ToString() + " WHERE " + whereString.ToString()
        command.CommandText = commandText
        Return command
    End Function

    Private Function CreateOldParam(column As DataColumn) As OleDbParameter
        Dim sqlParam As New OleDbParameter()
        Dim columnName As String = column.ColumnName
        sqlParam.ParameterName = "@old" + Replace(columnName, " ", "SPACE")
        sqlParam.SourceColumn = columnName
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam.DbType = TypeToDbType(column.DataType(), columnName, schemaDataTypeCollection)
        Return sqlParam
    End Function

    Private Function CreateParam(column As DataColumn) As OleDbParameter
        Dim sqlParam As New OleDbParameter()
        Dim columnName As String = column.ColumnName
        sqlParam.ParameterName = "@" + Replace(columnName, " ", "SPACE")
        sqlParam.SourceColumn = columnName
        sqlParam.DbType = TypeToDbType(column.DataType(), columnName, schemaDataTypeCollection)
        Return sqlParam
    End Function

    Private Function GetTextCommand(text As String) As OleDbCommand
        Dim command As New OleDbCommand With {
            .CommandType = CommandType.Text,
            .CommandText = text,
            .Connection = connection
        }
        Return command
    End Function

End Class
