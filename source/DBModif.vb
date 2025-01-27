Imports ExcelDna.Integration
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Collections
Imports System.Data
Imports System.Data.Common
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Linq
Imports System.Text


#Region "DBModifs"
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

    ''' <summary>constructor with definition XML</summary>
    ''' <param name="definitionXML"></param>
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

    ''' <summary>used to release com objects</summary>
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
    ''' <returns>the name of the DBModifier</returns>
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
    ''' <summary>For DBSheets the previous length of the DBMapper area is needed to tell additions with Ctrl+ from deletions with Ctrl-</summary>
    Public previousCUDLength As Integer

    ''' <summary>constructor with definition XML and the target range for the DBMapper</summary>
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
            Return Not isEmptyArray(testRange)
        Else
            Return MyBase.DBModifSaveNeeded()
        End If
    End Function

    ''' <summary>pass back whether changes were done by the DB Modif object (needed to prevent deadlocks/warn user due to refreshing the underlying area)</summary>
    ''' <returns>true if changes done</returns>
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
        ' All cells in DBMapper are relative to the start of TargetRange (incl a header row), so CUDMarkRow = changedRange.Row - TargetRange.Row + 1 ...
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
                    ' shows that row was deleted
                    If previousCUDLength > TargetRange.Rows.Count Then
                        Dim retval As MsgBoxResult = QuestionMsg("A whole row was deleted with Ctrl & -, the row will not be deleted in the database (use Ctrl-Shift-D instead, click cancel and refresh the DB-Sheet with Ctrl-Shift-R now)",, "Set CUD Flags for DB Mapper", MsgBoxStyle.Exclamation)
                        If retval = MsgBoxResult.Cancel Then GoTo exitSub
                        deleteFlag = True
                    Else
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
                End If

                Dim CUDMarkRow As Integer = changedRange.Row - TargetRange.Row + 1
                ' inside a list-object Ctrl & + and Ctrl & - add and remove a whole list-object range row, outside with selected row they add/remove a whole sheet row
                If (changedRangeColumns = targetRangeColumns Or changedRangeColumns = sheetColumns) And changedRangeRows = 1 And Not deleteFlag Then
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

                ExcelDnaUtil.Application.AutoCorrect.AutoExpandListRange = False ' to prevent automatic creation of new column
                If changedRangeRows = 1 And TargetRange.Cells(CUDMarkRow, targetRangeColumns + 1).Value = "" Then
                    If Not deleteFlag Then
                        ' check if row was added at the bottom set insert flag
                        If CUDMarkRow > TargetRange.Cells(targetRangeRows, targetRangeColumns).Row Then insertFlag = True
                        If insertFlag Then
                            TargetRange.Cells(CUDMarkRow, targetRangeColumns + 1).Value = "i"
                        Else
                            TargetRange.Cells(CUDMarkRow, targetRangeColumns + 1).Value = "u"
                            TargetRange.Rows(CUDMarkRow).Font.Italic = True
                        End If
                    End If
                Else
                    If changedRange.Row <= TargetRange.Row Then
						' copy/paste above the DBMapper is nonsense.
						Dim retval As MsgBoxResult = QuestionMsg("A data range was pasted above the data area of the DBSheet, this renders the DBSheet disfunctional. Immediately refresh the DBSheet to regain functionality",, "Set CUD Flags for DB Mapper", MsgBoxStyle.Exclamation)
						GoTo exitSub
					End If
					' copy/paste of large ranges needs quicker setting of u/i, only do if no CUD flags already set (all CUD cells are empty)
					' can't use ExcelDnaUtil.Application.WorksheetFunction.CountIfs(changedRange.Columns(targetRangeColumns + 1), "<>") = 0 here as it clears the flags as a side effect..
					If isEmptyArray(changedRange.Columns(targetRangeColumns + 1).Value) Then
						Dim nonintersecting As Excel.Range = getNonIntersectingRowsTarget(changedRange, TargetRange, TargetRange.Column + targetRangeColumns)
						Dim intersecting As Excel.Range = ExcelDnaUtil.Application.Intersect(changedRange, TargetRange)
						Dim intersectAndNonintersect As Excel.Range = Nothing
						' check if nonintersect CUD Marks and intersect Marks overlap, if yes avoid setting intersect marks...
						Try : intersectAndNonintersect = ExcelDnaUtil.Application.Intersect(nonintersecting, TargetRange.Range(ExcelDnaUtil.Application.Cells(intersecting.Row - TargetRange.Row + 1, targetRangeColumns + 1), ExcelDnaUtil.Application.Cells(intersecting.Row + intersecting.Rows.Count - TargetRange.Row, targetRangeColumns + 1))) : Catch ex As Exception : End Try
						If Not IsNothing(intersecting) And IsNothing(intersectAndNonintersect) Then
							TargetRange.Range(ExcelDnaUtil.Application.Cells(intersecting.Row - TargetRange.Row + 1, targetRangeColumns + 1), ExcelDnaUtil.Application.Cells(intersecting.Row + intersecting.Rows.Count - TargetRange.Row, targetRangeColumns + 1)).Value = "u"
							TargetRange.Range(ExcelDnaUtil.Application.Cells(intersecting.Row - TargetRange.Row + 1, 1), ExcelDnaUtil.Application.Cells(intersecting.Row + intersecting.Rows.Count - TargetRange.Row, targetRangeColumns)).Font.Italic = True
						End If
						If Not IsNothing(nonintersecting) Then nonintersecting.Value = "i"
					End If
                    With TargetRange.Rows(2).Interior
                        .Pattern = Excel.XlPattern.xlPatternNone
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    ExcelDnaUtil.Application.AutoCorrect.AutoExpandListRange = True
                End If
                previousCUDLength = TargetRange.Rows.Count
            End If
        Catch ex As Exception
            LogWarn("Exception in insertCUDMarks: " + ex.Message)
        End Try
exitSub:
        preventChangeWhileFetching = False
        ExcelDnaUtil.Application.Statusbar = False
    End Sub

    ''' <summary>checks whether passed array has only empty entries</summary>
    ''' <param name="theArray"></param>
    ''' <returns>true if empty, else false</returns>
    Function isEmptyArray(theArray As Object) As Boolean
        For Each theVal In theArray
            If Not IsNothing(theVal) Then Return False
        Next
        Return True
    End Function

    ''' <summary>gets the non intersecting rows of two ranges as the single column range given in tgtcolumn</summary>
    ''' <param name="range1">first range whose non-intersected rows should be returned</param>
    ''' <param name="range2">second range whose non-intersected rows should be returned</param>
    ''' <param name="tgtcolumn">column of the returned range</param>
    ''' <returns></returns>
    Function getNonIntersectingRowsTarget(range1 As Excel.Range, range2 As Excel.Range, tgtcolumn As Integer) As Excel.Range
        If IsNothing(range1) Or IsNothing(range2) Then Return Nothing
        With ExcelDnaUtil.Application
            Dim theUnion As Excel.Range = .Union(range1, range2)
            Dim theintersect As Excel.Range = .Intersect(range1, range2)
            If IsNothing(theUnion) Then Return Nothing
            If IsNothing(theintersect) Then
                ' if no intersection, then non intersection is simply the changed range
                Return .Range(.Cells(range1.Row, tgtcolumn), .Cells(range1.Row + range1.Rows.Count - 1, tgtcolumn))
            Else
                If (range1.Row + range1.Rows.Count < range2.Row + range2.Rows.Count) Then
                    ' if changed area ends above the existing DBSheet and DBSheet was extended, then non intersection is simply the changed range
                    If TargetRange.Rows.Count > previousCUDLength Then
                        Return .Range(.Cells(range1.Row, tgtcolumn), .Cells(range1.Row + range1.Rows.Count - 1, tgtcolumn))
                    Else
                        ' if changed area ends above the existing DBSheet and DBSheet was not extended, then non intersection is nothing
                        Return Nothing
                    End If
                Else
                    ' return from last intersection row (exclusive) to lowest common cell
                    Return .Range(.Cells(theintersect.Row + theintersect.Rows.Count, tgtcolumn), .Cells(theUnion.Row + theUnion.Rows.Count - 1, tgtcolumn))
                End If
            End If
        End With
    End Function

    ''' <summary>extend DataRange to "whole" DBMApper area (first row (header/field names) to the right and first column (first primary key) down)</summary>
    Public Sub extendDataRange()
        Dim actWbNames As Excel.Names
        Try : actWbNames = ExcelDnaUtil.Application.ActiveWorkbook.Names : Catch ex As Exception
            UserMsg("Exception when trying to get the active workbook names for extending DBMapper Range: " + ex.Message + ", this might be either due to errors in the VBA-IDE (missing references) or due to opening this workbook from an MS-Office hyperlink, starting up Excel (timing issue). Switch to another workbook and back to fix.")
            Exit Sub
        End Try
        Try
            ' only extend like this if no CUD Flags or AutoIncFlag present (may have non existing first (primary) columns -> auto identity columns !)
            If Not (CUDFlags Or AutoIncFlag) Then
                If TargetRange.Cells(2, 1).Value Is Nothing Then Exit Sub ' only extend if there are multiple rows...
                Dim rowCount As Integer = TargetRange.Cells(1, 1).End(Excel.XlDirection.xlDown).Row - TargetRange.Cells(1, 1).Row + 1
                ' unfortunately the above method to find the column extent doesn't work with hidden columns, so count the filled cells directly...
                Dim colCount As Integer = 1
                ' if we don't want columns to automatically extend then take original column count
                If preventColResize Then
                    colCount = TargetRange.Columns.Count
                Else
                    While Not (TargetRange.Cells(1, colCount + 1).Value Is Nothing OrElse TargetRange.Cells(1, colCount + 1).Value.ToString() = "")
                        colCount += 1
                    End While
                End If
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
        Catch ex As Exception
            UserMsg(ex.Message)
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
                TargetRange.ListObject.DataBodyRange.Columns(TargetRange.Columns.Count + 1).Clear
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
        If (CUDFlags And primKeysCount = 0) Then
            UserMsg("CUD Flags are incompatible with primKeysCount = 0, exiting DBModification", "DBMapper Error")
            Exit Sub
        End If
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

        ' openingQuote and closingQuote are needed to quote columns containing blanks and other not allowed characters
        openingQuote = fetchSetting("openingQuote" + env.ToString(), "")
        closingQuote = fetchSetting("closingQuote" + env.ToString(), "")
        closingQuoteReplacement = fetchSetting("closingQuoteReplacement" + env.ToString(), "")

        If deleteBeforeMapperInsert Then
            Try
                Dim DmlCmd As IDbCommand = idbcnn.CreateCommand()
                ' for deleteBeforeMapperInsert outside of DBSequence still begin a transaction and reset hadError because both the deleting of the table and the new insert should be atomic
                If Not TransactionOpen Then
                    hadError = False
                    DBModifHelper.trans = idbcnn.BeginTransaction()
                End If
                With DmlCmd
                    .Transaction = DBModifHelper.trans
                    .CommandText = "DELETE FROM " + openingQuote + String.Join(closingQuote + "." + openingQuote, Strings.Split(tableName, ".")) + closingQuote
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
            ' INSTANCE.DBNAME.dbo.tablename becomes [INSTANCE].[DBNAME].[dbo].[tablename]
            Dim SelectStmt As String = "SELECT * FROM " + openingQuote + String.Join(closingQuote + "." + openingQuote, Strings.Split(tableName, ".")) + closingQuote
            If TypeName(idbcnn) = "SqlConnection" Then
                ' decent behaviour for SQL Server
                Using comm As New SqlCommand("SET ARITHABORT ON", idbcnn, DBModifHelper.trans)
                    comm.ExecuteNonQuery()
                End Using
                da = New SqlDataAdapter(New SqlCommand(SelectStmt, idbcnn, DBModifHelper.trans)) With {
                    .UpdateBatchSize = 20,
                    .ContinueUpdateOnError = True ' this is needed to check the error rows after update and present them
                }
            ElseIf TypeName(idbcnn) = "OleDbConnection" Then
                da = New OleDbDataAdapter(New OleDbCommand(SelectStmt, idbcnn, DBModifHelper.trans)) With {
                    .UpdateBatchSize = 20,
                    .ContinueUpdateOnError = True
                }
            Else
                da = New OdbcDataAdapter(New OdbcCommand(SelectStmt, idbcnn, DBModifHelper.trans)) With {
                    .UpdateBatchSize = 20,
                    .ContinueUpdateOnError = True
                }
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

        FieldParamMap = New Dictionary(Of String, String)
        ' first check if all column names (except ignored) of DBMapper Range exist in table and collect field-names
        Dim allColumnsStr(TargetRange.Columns.Count - 1) As String
        Dim allColumnsStrUnQuoted(TargetRange.Columns.Count - 1) As String
        Dim colNum As Integer = 1
        Dim fieldColNum As Integer = 0
        Dim ignoreMe As Boolean
        Do
            ignoreMe = False
            Dim fieldname As String = Trim(TargetRange.Cells(1, colNum).Value)
            ' only if not ignored...
            If InStr(1, LCase(ignoreColumns) + ",", LCase(fieldname) + ",") = 0 Then
                If Not dscheck.Tables(0).Columns.Contains(fieldname) Then
                    hadError = True
                    TargetRange.Parent.Activate
                    TargetRange.Cells(1, colNum).Select
                    UserMsg("Field '" + fieldname + "' does not exist in Table '" + tableName + "' and is not in ignoreColumns, Error in sheet " + TargetRange.Parent.Name, "DBMapper Error")
                    GoTo cleanup
                End If
            ElseIf Strings.Right(fieldname.ToUpper(), 2) = "LU" And dscheck.Tables(0).Columns.Contains(Left(fieldname, Len(fieldname) - 2)) Then
                ' if ignored and the column is a lookup then add the column without LU to preserve the order of columns !!!
                fieldname = Left(fieldname, Len(fieldname) - 2) ' correct the LU to real field-name, real field is then skipped as order is important here!
            Else
                ignoreMe = True
            End If

            ' only add if the field was not already added by the LU correction!
            If Not allColumnsStr.Contains(fieldname) And Not ignoreMe Then
                allColumnsStr(fieldColNum) = CorrectQuotes(fieldname)
                allColumnsStrUnQuoted(fieldColNum) = fieldname
                FieldParamMap.Add(fieldname, "Param" + CStr(fieldColNum))
                fieldColNum += 1
            End If

            colNum += 1 ' excel column
        Loop Until colNum > TargetRange.Columns.Count

        ' keep only those that were filled...
        ReDim Preserve allColumnsStr(fieldColNum - 1)
        ReDim Preserve allColumnsStrUnQuoted(fieldColNum - 1)

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
            primKeyCompound += openingQuote + CorrectQuotes(primKey) + closingQuote + " = @" + FieldParamMap(primKey) + IIf(i = primKeysCount, "", " AND ")
        Next

        ' the actual dataset
        Dim ds As New DataSet()

        Dim BaseSelect As String = "SELECT " + openingQuote + String.Join(closingQuote + "," + openingQuote, allColumnsStr) + closingQuote + " FROM " + openingQuote + String.Join(closingQuote + "." + openingQuote, Strings.Split(tableName, ".")) + closingQuote
        ' replace the select command to avoid columns that are default filled but non null-able (will produce errors if not assigned to new row)
        da.SelectCommand.CommandText = BaseSelect
        ' fill schema again to reflect the changed columns
        Try
            da.FillSchema(ds, SchemaType.Source, tableName)
        Catch ex As Exception
            notifyUserOfDataError("Error in getting schema information for " + tableName + " (" + da.SelectCommand.CommandText + "): " + ex.Message, 1)
            GoTo cleanup
        End Try

        Dim allColumns(UBound(allColumnsStrUnQuoted)) As DataColumn
        For i As Integer = 0 To UBound(allColumnsStrUnQuoted)
            allColumns(i) = ds.Tables(0).Columns(allColumnsStrUnQuoted(i))
        Next

        ' for avoidFill mode (no uploading of whole table) amend select with parameter clause (parameters are added below)
        If avoidFill Then da.SelectCommand.CommandText = BaseSelect + primKeyCompound

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
                    .ParameterName = "@" + FieldParamMap(primKeyColumnsStr(i))
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
            da.UpdateCommand.Transaction = DBModifHelper.trans
            da.DeleteCommand.UpdatedRowSource = UpdateRowSource.None
            da.DeleteCommand.Transaction = DBModifHelper.trans
            da.InsertCommand.UpdatedRowSource = UpdateRowSource.None
            da.InsertCommand.Transaction = DBModifHelper.trans
        Catch ex As Exception
            notifyUserOfDataError("Error in setting Transaction for Insert/Update/Delete Commands for Data Adapter for " + tableName + ": " + ex.Message, 1)
            GoTo cleanup
        End Try

        Dim rowNum As Long = 2
        ' walk through all rows in DBMapper Target-range to store in data set
        Dim finishLoop As Boolean
        Dim totalColCount As Long = TargetRange.Columns.Count
        Dim totalRowCount As Long = TargetRange.Rows.Count
        Do
            ' if CUDFlags are set, only insert/update/delete if CUDFlags column (right to DBMapper range) is filled...
            Dim rowCUDFlag As String = TargetRange.Cells(rowNum, totalColCount + 1).Value
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
                        da.SelectCommand.Parameters.Item("@" + FieldParamMap(primKey)).Value = primKeyValues(i - 1)
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
                ' also avoid finding if no data should be in the table (deleteBeforeMapperInsert) and there is an explicit "i"nsert in a DBSheet
                Dim foundRow As DataRow = Nothing
                If Not AutoIncrement And Not deleteBeforeMapperInsert And primKeysCount <> 0 And rowCUDFlag <> "i" Then
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
                    If insertIfMissing Or rowCUDFlag = "i" Or deleteBeforeMapperInsert Or primKeysCount = 0 Then
                        insertRecord = True
                        ' ... add a new record if insertIfMissing flag is set Or CUD Flag insert or no primKey is given
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
                    ExcelDnaUtil.Application.StatusBar = Left("Inserting " + IIf(AutoIncrement, "new auto-incremented key", primKeyValueStr) + " into " + tableName + ", " + CStr(rowNum) + "/" + CStr(totalRowCount), 255)
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
                    Loop Until colNum > totalColCount

                    If insertRecord Then
                        Try
                            ds.Tables(0).Rows.Add(foundRow)
                        Catch ex As Exception
                            If Not notifyUserOfDataError("Error inserting row " + rowNum.ToString() + " in sheet " + TargetRange.Parent.Name + ": " + ex.Message, rowNum) Then GoTo cleanup
                        End Try
                    Else
                        ExcelDnaUtil.Application.StatusBar = Left("Updated fields for " + primKeyValueStr + " in " + tableName + ", " + CStr(rowNum) + "/" + CStr(totalRowCount), 255)
                    End If
                End If

                ' delete only with CUDFlags...
                If (CUDFlags And rowCUDFlag = "d") Then
                    Try
                        foundRow.Delete()
                    Catch ex As Exception
                        If Not notifyUserOfDataError("Error deleting row " + rowNum.ToString() + " in sheet " + TargetRange.Parent.Name + ": " + ex.Message, rowNum) Then GoTo cleanup
                    End Try
                    ExcelDnaUtil.Application.StatusBar = Left("Deleting " + primKeyValueStr + " in " + tableName + ", " + CStr(rowNum) + "/" + CStr(totalRowCount), 255)
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
        ExcelDnaUtil.Application.StatusBar = Left("Applying modifications (inserts/updates/deletes) in Database for " + tableName, 255)
        ' no capture of exception here as 1) the adapters were created with ContinueUpdateOnError=true and 2) the table is checked for errors after the update
        ' in this way the actual rows where the database error occured can be retrieved/presented to the user
        da.Update(ds.Tables(0))
        If ds.Tables(0).HasErrors Then
            Dim errMessage As String = ""
            Try
                For Each row As DataRow In ds.Tables(0).GetErrors()
                    errMessage += row.RowError + vbCrLf + vbCrLf
                    Dim infoStr As String = ""
                    For i As Integer = 0 To row.ItemArray.Count - 1
                        infoStr += ds.Tables(0).Columns(i).ToString + ": " + row.ItemArray(i).ToString + vbCrLf
                    Next
                    errMessage += infoStr + vbCrLf
                Next
            Catch ex As Exception : errMessage += ", " + ex.Message : End Try
            ' hadError = True ... don't need to set as it is set by notifyUserOfDataError
            If Not notifyUserOfDataError("Update Error in sheet " + TargetRange.Parent.Name + ": " + vbCrLf + errMessage) Then GoTo cleanup
        Else
            changesDone = True
        End If

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
                DBModifHelper.trans.Rollback()
            Else
                DBModifHelper.trans.Commit()
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
    ''' <param name="rowNum">error cell row (optional, if not given then everything is selected)</param>
    ''' <param name="colNum">error cell column (optional, if not given then whole row is selected)</param>
    ''' <returns></returns>
    Private Function notifyUserOfDataError(message As String, Optional rowNum As Long = -1, Optional colNum As Integer = -1) As Boolean
        hadError = True
        If Not nonInteractive Then
            TargetRange.Parent.Activate
            If rowNum = -1 Then
                TargetRange.Select()
            Else
                If colNum = -1 Then
                    TargetRange.Rows(rowNum).Select()
                Else
                    TargetRange.Cells(rowNum, colNum).Select()
                End If
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

    ''' <summary>constructor with definition XML and the target range for the cells containing the DBAction</summary>
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

    ''' <summary>do the modification defined in DBAction</summary>
    ''' <param name="WbIsSaving">for asking for confirmation and exiting</param>
    ''' <param name="calledByDBSeq">if inside a sequence (defined by the name in this parameter), no User message is displayed and the connection is not closed</param>
    ''' <param name="TransactionOpen">if a transaction is open, no database connection needs to be opened</param>
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

    ''' <summary>execute the SQL text given in executeText</summary>
    ''' <param name="executeText">the sql code to be executed</param>
    ''' <returns>the number of rows affected by the sql code</returns>
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

    ''' <summary>constructor with definition XML</summary>
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
                    DBModifHelper.trans = idbcnn.BeginTransaction()
                    TransactionIsOpen = True
                Case "DBCommitRollback"
                    If Not hadError Then
                        LogInfo("DBCommitTrans... ")
                        DBModifHelper.trans.Commit()
                    Else
                        LogInfo("DBRollbackTrans... ")
                        DBModifHelper.trans.Rollback()
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

''' <summary>Dummy DBModif Class for executeRefresh during externally callable executeDBModif procedure. No data, just executeRefresh(srcExtent) procedure to refresh DBMappers</summary>
Public Class DBModifDummy : Inherits DBModif
    Public Sub New()
        MyBase.New(Nothing)
    End Sub

    ''' <summary>to refresh DB Modifiers from VBA Macros (externally callable executeDBModif procedure), the DBMappers underlying DB function (DBlistfetch or DBSetQuery) can be called using its srcExtent</summary>
    ''' <param name="srcExtent">DBFsource....</param>
    Public Sub executeRefresh(srcExtent)
        doDBRefresh(srcExtent:=srcExtent)
    End Sub
End Class

#End Region


#Region "CustomCommandbuilders"

''' <summary>Custom Command builder base class for SQL Server, ODBC and OLE DB to avoid primary key problems with built-in ones
''' derived (transposed into VB.NET) from https://www.cogin.com/articles/CustomCommandBuilder.php
''' Copyright by Dejan_Grujic 2004. http://www.cogin.com
''' </summary>
Public Class CustomCommandBuilder
    ''' <summary>the data table to get all the schema information from</summary>
    Protected ReadOnly dataTable As DataTable
    ''' <summary>array of all column information (DataColumn) in the data table</summary>
    Protected ReadOnly allColumns As DataColumn()

    Public Sub New(dataTable As DataTable, allColumns As DataColumn())
        Me.dataTable = dataTable
        Me.allColumns = allColumns
    End Sub

    ''' <summary>get auto increment columns</summary>
    ''' <returns>array of auto increment DataColumn entries</returns>
    Protected Function AutoincrementKeyColumns() As ArrayList
        AutoincrementKeyColumns = New ArrayList
        For Each primaryKeyColumn As DataColumn In dataTable.PrimaryKey
            If primaryKeyColumn.AutoIncrement Then
                AutoincrementKeyColumns.Add(primaryKeyColumn)
            End If
        Next
    End Function

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

    ''' <summary>build correct quoted table name from the dataTables TableName</summary>
    ''' <returns>correct table name</returns>
    Protected Function TableName() As String
        ' each identifier needs to be quoted by itself
        Return openingQuote + String.Join(closingQuote + "." + openingQuote, Strings.Split(dataTable.TableName, ".")) + closingQuote
    End Function

End Class


''' <summary>Custom Command builder for SQLServer class</summary>
Public Class CustomSqlCommandBuilder
    Inherits CustomCommandBuilder

    ''' <summary>the driver specific connection needed to get the driver specific command in GetTextCommand</summary>
    Private ReadOnly connection As SqlConnection
    ''' <summary>collection for casting .NET data type to ADO.NET DbType in DBModifHelper.TypeToDbType</summary>
    Private ReadOnly schemaDataTypeCollection As Collection

    Public Sub New(dataTable As DataTable, connection As SqlConnection, allColumns As DataColumn(), schemaDataTypeCollection As Collection)
        MyBase.New(dataTable, allColumns)
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
                intoString.Append(openingQuote).Append(CorrectQuotes(column.ColumnName)).Append(closingQuote)
                valuesString.Append("@").Append(FieldParamMap(column.ColumnName))
                command.Parameters.Add(CreateParam(column))
            End If
        Next
        Dim commandText As String = "INSERT INTO " + TableName() + "(" + intoString.ToString() + ") VALUES (" + valuesString.ToString() + "); "
        If autoincrementColumns.Count > 0 Then
            commandText += "SELECT SCOPE_IDENTITY() AS " + openingQuote + CorrectQuotes(DirectCast(autoincrementColumns(0), DataColumn).ColumnName()) + closingQuote
        End If
        command.CommandText = commandText
        Return command
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
            whereString.Append(openingQuote).Append(CorrectQuotes(column.ColumnName)).Append(closingQuote).Append(" = @").Append(FieldParamMap(column.ColumnName))
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
                whereString.Append(openingQuote).Append(CorrectQuotes(column.ColumnName)).Append(closingQuote).Append("= @old").Append(FieldParamMap(column.ColumnName))
            Else
                ' other fields go into set part
                If (setString.Length > 0) Then
                    setString.Append(", ")
                End If
                setString.Append(openingQuote).Append(CorrectQuotes(column.ColumnName)).Append(closingQuote).Append(" = @").Append(FieldParamMap(column.ColumnName))
            End If
            command.Parameters.Add(CreateParam(column))
            command.Parameters.Add(CreateOldParam(column))
        Next
        Dim commandText As String = "UPDATE " + TableName() + " SET " + setString.ToString() + " WHERE " + whereString.ToString()
        command.CommandText = commandText
        Return command
    End Function

    ''' <summary>build "old" (for where clause) SqlParameter for DataColumn column</summary>
    ''' <param name="column">DataColumn to build parameter for</param>
    ''' <returns>parameter for DataColumn</returns>
    Private Function CreateOldParam(column As DataColumn) As SqlParameter
        Dim sqlParam As New SqlParameter()
        Dim columnName As String = column.ColumnName
        sqlParam.ParameterName = "@old" + FieldParamMap(columnName)
        sqlParam.SourceColumn = columnName
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam.DbType = TypeToDbType(column.DataType(), columnName, schemaDataTypeCollection)
        Return sqlParam
    End Function

    ''' <summary>build "new" (for set clause) SqlParameter for DataColumn column</summary>
    ''' <param name="column">DataColumn to build parameter for</param>
    ''' <returns>parameter for DataColumn</returns>
    Private Function CreateParam(column As DataColumn) As SqlParameter
        Dim sqlParam As New SqlParameter()
        Dim columnName As String = column.ColumnName
        sqlParam.ParameterName = "@" + FieldParamMap(columnName)
        sqlParam.SourceColumn = columnName
        sqlParam.DbType = TypeToDbType(column.DataType(), columnName, schemaDataTypeCollection)
        Return sqlParam
    End Function

    ''' <summary>build a specific Sql command to be used for the Insert, Update and Delete builders</summary>
    ''' <param name="text">command text, not used here</param>
    ''' <returns>specific command</returns>
    Private Function GetTextCommand(text As String) As SqlCommand
        Dim command As New SqlCommand With {
            .CommandType = CommandType.Text,
            .CommandText = text,
            .Connection = connection
        }
        Return command
    End Function

End Class


''' <summary>Custom Command builder for ODBC class</summary>
Public Class CustomOdbcCommandBuilder
    Inherits CustomCommandBuilder

    ''' <summary>the driver specific connection needed to get the driver specific command in GetTextCommand</summary>
    Private ReadOnly connection As OdbcConnection
    ''' <summary>collection for casting .NET data type to ADO.NET DbType in DBModifHelper.TypeToDbType</summary>
    Private ReadOnly schemaDataTypeCollection As Collection

    Public Sub New(dataTable As DataTable, connection As OdbcConnection, allColumns As DataColumn(), schemaDataTypeCollection As Collection)
        MyBase.New(dataTable, allColumns)
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
                intoString.Append(openingQuote).Append(CorrectQuotes(column.ColumnName)).Append(closingQuote)
                valuesString.Append("@").Append(FieldParamMap(column.ColumnName))
                command.Parameters.Add(CreateParam(column))
            End If
        Next
        Dim commandText As String = "INSERT INTO " + TableName() + "(" + intoString.ToString() + ") VALUES (" + valuesString.ToString() + "); "
        If autoincrementColumns.Count > 0 Then
            commandText += "SELECT SCOPE_IDENTITY() AS " + openingQuote + CorrectQuotes(DirectCast(autoincrementColumns(0), DataColumn).ColumnName()) + closingQuote
        End If
        command.CommandText = commandText
        Return command
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
            whereString.Append(openingQuote).Append(CorrectQuotes(column.ColumnName)).Append(closingQuote).Append(" = @").Append(FieldParamMap(column.ColumnName))
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
                whereString.Append(openingQuote).Append(CorrectQuotes(column.ColumnName)).Append(closingQuote).Append("= @old").Append(FieldParamMap(column.ColumnName))
            Else
                ' other fields go into set part
                If (setString.Length > 0) Then
                    setString.Append(", ")
                End If
                setString.Append(openingQuote).Append(CorrectQuotes(column.ColumnName)).Append(closingQuote).Append(" = @").Append(FieldParamMap(column.ColumnName))
            End If
            command.Parameters.Add(CreateParam(column))
            command.Parameters.Add(CreateOldParam(column))
        Next
        Dim commandText As String = "UPDATE " + TableName() + " SET " + setString.ToString() + " WHERE " + whereString.ToString()
        command.CommandText = commandText
        Return command
    End Function

    ''' <summary>build "old" (for where clause) OdbcParameter for DataColumn column</summary>
    ''' <param name="column">DataColumn to build parameter for</param>
    ''' <returns>parameter for DataColumn</returns>
    Private Function CreateOldParam(column As DataColumn) As OdbcParameter
        Dim sqlParam As New OdbcParameter()
        Dim columnName As String = column.ColumnName
        sqlParam.ParameterName = "@old" + FieldParamMap(columnName)
        sqlParam.SourceColumn = columnName
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam.DbType = TypeToDbType(column.DataType(), columnName, schemaDataTypeCollection)
        Return sqlParam
    End Function

    ''' <summary>build "new" (for set clause) OdbcParameter for DataColumn column</summary>
    ''' <param name="column">DataColumn to build parameter for</param>
    ''' <returns>parameter for DataColumn</returns>
    Private Function CreateParam(column As DataColumn) As OdbcParameter
        Dim sqlParam As New OdbcParameter()
        Dim columnName As String = column.ColumnName
        sqlParam.ParameterName = "@" + FieldParamMap(columnName)
        sqlParam.SourceColumn = columnName
        sqlParam.DbType = TypeToDbType(column.DataType(), columnName, schemaDataTypeCollection)
        Return sqlParam
    End Function

    ''' <summary>build a specific Odbc command to be used for the Insert, Update and Delete builders</summary>
    ''' <param name="text">command text, not used here</param>
    ''' <returns>specific command</returns>
    Private Function GetTextCommand(text As String) As OdbcCommand
        Dim command As New OdbcCommand With {
            .CommandType = CommandType.Text,
            .CommandText = text,
            .Connection = connection
        }
        Return command
    End Function
End Class


''' <summary>Custom Command builder for OleDB class</summary>
Public Class CustomOleDbCommandBuilder
    Inherits CustomCommandBuilder

    ''' <summary>the driver specific connection needed to get the driver specific command in GetTextCommand</summary>
    Private ReadOnly connection As OleDbConnection
    ''' <summary>collection for casting .NET data type to ADO.NET DbType in DBModifHelper.TypeToDbType</summary>
    Private ReadOnly schemaDataTypeCollection As Collection

    Public Sub New(dataTable As DataTable, connection As OleDbConnection, allColumns As DataColumn(), schemaDataTypeCollection As Collection)
        MyBase.New(dataTable, allColumns)
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
                intoString.Append(openingQuote).Append(CorrectQuotes(column.ColumnName)).Append(closingQuote)
                valuesString.Append("@").Append(FieldParamMap(column.ColumnName))
                command.Parameters.Add(CreateParam(column))
            End If
        Next
        Dim commandText As String = "INSERT INTO " + TableName() + "(" + intoString.ToString() + ") VALUES (" + valuesString.ToString() + "); "
        If autoincrementColumns.Count > 0 Then
            commandText += "SELECT SCOPE_IDENTITY() AS " + openingQuote + CorrectQuotes(DirectCast(autoincrementColumns(0), DataColumn).ColumnName()) + closingQuote
        End If
        command.CommandText = commandText
        Return command
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
            whereString.Append(openingQuote).Append(CorrectQuotes(column.ColumnName)).Append(closingQuote).Append(" = @").Append(FieldParamMap(column.ColumnName))
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
                whereString.Append(openingQuote).Append(CorrectQuotes(column.ColumnName)).Append(closingQuote).Append("= @old").Append(FieldParamMap(column.ColumnName))
            Else
                ' other fields go into set part
                If (setString.Length > 0) Then
                    setString.Append(", ")
                End If
                setString.Append(openingQuote).Append(CorrectQuotes(column.ColumnName)).Append(closingQuote).Append(" = @").Append(FieldParamMap(column.ColumnName))
            End If
            command.Parameters.Add(CreateParam(column))
            command.Parameters.Add(CreateOldParam(column))
        Next
        Dim commandText As String = "UPDATE " + TableName() + " SET " + setString.ToString() + " WHERE " + whereString.ToString()
        command.CommandText = commandText
        Return command
    End Function

    ''' <summary>build "old" (for where clause) OdbcParameter for DataColumn column</summary>
    ''' <param name="column">DataColumn to build parameter for</param>
    ''' <returns>parameter for DataColumn</returns>
    Private Function CreateOldParam(column As DataColumn) As OleDbParameter
        Dim sqlParam As New OleDbParameter()
        Dim columnName As String = column.ColumnName
        sqlParam.ParameterName = "@old" + FieldParamMap(columnName)
        sqlParam.SourceColumn = columnName
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam.DbType = TypeToDbType(column.DataType(), columnName, schemaDataTypeCollection)
        Return sqlParam
    End Function

    ''' <summary>build "new" (for set clause) OdbcParameter for DataColumn column</summary>
    ''' <param name="column">DataColumn to build parameter for</param>
    ''' <returns>parameter for DataColumn</returns>
    Private Function CreateParam(column As DataColumn) As OleDbParameter
        Dim sqlParam As New OleDbParameter()
        Dim columnName As String = column.ColumnName
        sqlParam.ParameterName = "@" + FieldParamMap(columnName)
        sqlParam.SourceColumn = columnName
        sqlParam.DbType = TypeToDbType(column.DataType(), columnName, schemaDataTypeCollection)
        Return sqlParam
    End Function

    ''' <summary>build a specific OleDb command to be used for the Insert, Update and Delete builders</summary>
    ''' <param name="text">command text, not used here</param>
    ''' <returns>specific command</returns>
    Private Function GetTextCommand(text As String) As OleDbCommand
        Dim command As New OleDbCommand With {
            .CommandType = CommandType.Text,
            .CommandText = text,
            .Connection = connection
        }
        Return command
    End Function

End Class

#End Region