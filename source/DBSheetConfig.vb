Imports System.IO
Imports System.Windows.Forms
Imports ExcelDna.Integration
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop


'''<summary>Helper class for DBSheetHandler and DBSheetConnection for easier manipulation of DBSheet definition / Connection configuration data</summary> 
Public Module DBSheetConfig
    ''' <summary>the current cell where the DBSheet Definition is inserted at</summary>
    Dim curCell As Excel.Range
    ''' <summary>the list object of the main query for the db mapper</summary>
    Dim createdListObject As Excel.ListObject
    ''' <summary>the lookups list of the DBSheet definition (xml element with query, name, etc.)</summary>
    Dim lookupsList() As String
    ''' <summary>the complete dbsheet configuration (XML)</summary>
    Dim curConfig As String
    ''' <summary>the added and hidden worksheet with lookups inside</summary>
    Dim lookupWS As Excel.Worksheet
    ''' <summary>the Database table name of the DBSheet</summary>
    Dim tableName As String
    ''' <summary>counter to know how many cells we filled for the dbmapper query 
    ''' (at least 2: dbsetquery function and query string, if additional where clause exists, 
    ''' add one for where clause, then one for each parameter)
    ''' </summary>
    Dim addedCells As Integer

    Public Sub createDBSheet()
        'MsgBox("not yet implemented..")
        'Exit Sub
        Dim openFileDialog1 = New OpenFileDialog With {
            .InitialDirectory = fetchSetting("DBSheetDefinitions", ""),
            .Filter = "XML files (*.xml)|*.xml",
            .RestoreDirectory = True
        }
        Dim result As DialogResult = openFileDialog1.ShowDialog()
        If result = Windows.Forms.DialogResult.OK Then
            ' store currently selected cell, where DBSetQuery for DBMapper will be placed.
            curCell = ExcelDnaUtil.Application.ActiveCell
            ' Get the DBSheet Definition file name and read into curConfig
            Dim dsdPath As String = openFileDialog1.FileName
            curConfig = File.ReadAllText(dsdPath, Text.Encoding.Default)
            ' get the table name of the DBSheet for setting the DBMapper name
            tableName = getEntry("table", curConfig)
            ' get query
            Dim queryStr As String = getEntry("query", curConfig)
            If queryStr = "" Then
                MsgBox("No query found in DBSheetConfig.", vbCritical, "DBSheet Creation Error")
                Exit Sub
            End If
            Dim whereClause As String = getEntry("whereClause", curConfig)
            ' queryStr inserted below DBSetQuery
            addedCells = 1
            Dim changedWhereClause As String = ""
            If whereClause <> "" Then
                ' check for where clauses and modify for parameter setting in formula
                changedWhereClause = "="
                Dim whereParts As String() = Split(whereClause, "?")
                For i = 0 To UBound(whereParts)
                    If whereParts(i) <> "" Then
                        ' each parameter adds a cell below DBSetQuery
                        addedCells += 1
                        ' create concatenation formula for parameter setting, each ? is replaced by a separate row reference below the where clause
                        changedWhereClause += IIf(i = 0, """WHERE ", "&""") & whereParts(i) & """&R[" & i + 1 & "]C"
                    End If
                Next
                queryStr = Replace(queryStr, "WHERE " & whereClause, "")
                ' whereClause inserted below queryStr
                addedCells += 1
            End If
            ' get lookup fields in complete columns definitions
            lookupsList = getEntryList("columns", "field", "lookup", curConfig, True)
            If Not IsNothing(lookupsList) Then
                lookupWS = ExcelDnaUtil.Application.ActiveWorkbook.Worksheets.Add()
                Try
                    lookupWS.Name = Left("LS" & tableName, 31)
                Catch ex As Exception
                    MsgBox("Error setting lookup Worksheet Name to '" & Left("LS" & tableName, 31) & "':" + ex.Message, vbCritical, "DBSheet Creation Error")
                    lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
                    Exit Sub
                End Try
                ' add lookup Queries in separate sheet
                Dim lookupCol As Integer = 1
                For Each LookupDef As String In lookupsList
                    ' fetch Lookupquery and get rid of template table def
                    Dim LookupQuery As String = Replace(getEntry("lookup", LookupDef, 1), "!T!", "T1")
                    Dim lookupName As String = Replace(getEntry("name", LookupDef, 1), "*", "")
                    ' replace looked up ID names with ID name + "LU" in query string
                    Dim foundDelim As Integer
                    For Each delimStr As String In {",", " ", vbCrLf}
                        foundDelim = InStr(queryStr, " " & lookupName & delimStr)
                        If foundDelim > 0 Then
                            queryStr = Replace(queryStr, " " & lookupName & delimStr, " " & lookupName & "LU" & delimStr)
                            Exit For
                        End If
                        foundDelim = InStr(queryStr, "." & lookupName & delimStr)
                        If foundDelim > 0 Then
                            queryStr = Replace(queryStr, "." & lookupName & delimStr, "." & lookupName & " " & lookupName & "LU" & delimStr)
                            Exit For
                        End If
                    Next
                    If foundDelim = 0 Then
                        MsgBox("Error in changing lookupName '" & lookupName & "' to '" & lookupName & "LU' in select statement of DBSheet query, it has to begin with blank and end with ','blank or CrLf !", vbCritical, "DBSheet Creation Error")
                        lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
                        Exit Sub
                    End If
                    If getEntry("fkey", LookupDef, 1) <> "" Then 'database lookup
                        lookupWS.Cells(1, lookupCol + 1).Value = LookupQuery
                        lookupWS.Cells(1, lookupCol + 1).WrapText = False
                        lookupWS.Cells(2, lookupCol).Name = lookupName & "Lookup"
                        ' then create the DBListFetch with the lookup query 
                        ConfigFiles.createFunctionsInCells(lookupWS.Cells(1, lookupCol), {"RC", "=DBListFetch(RC[1],""""," & lookupName & "Lookup" & ")"})
                        ' database lookups have two columns
                        lookupCol += 2
                    Else                                         'fixed value lookup
                        Dim lookupValues As String() = Split(LookupQuery, "||")
                        Dim lrow As Integer
                        For lrow = 0 To UBound(lookupValues)
                            lookupWS.Cells(2 + lrow, lookupCol).value = lookupValues(lrow)
                        Next
                        lookupWS.Range(lookupWS.Cells(2, lookupCol), lookupWS.Cells(2 + lrow - 1, lookupCol)).Name = lookupName & "Lookup"
                        ' fixed value lookups have only one column
                        lookupCol += 1
                    End If
                Next
                lookupWS.Visible = Excel.XlSheetVisibility.xlSheetHidden
                curCell.Parent.Select()
            End If
            ' then create the DBSetQuery assigning the (yet to be filled) query to the above listobject
            ' add DBSetQuery with queryStr as Basis for the final DBMapper
            ' first create a ListObject
            createdListObject = ConfigFiles.createListObject(curCell)
            If IsNothing(createdListObject) Then Exit Sub
            With curCell
                ' add the query as text
                Try
                    .Offset(1, 0).Value = queryStr
                    .Offset(1, 0).WrapText = False
                Catch ex As Exception
                    MsgBox("Error in adding query (" & queryStr & ")", vbCritical, "DBSheet Creation Error")
                    lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
                    Exit Sub
                End Try
                ' add an additional where clause as a concatenation string
                If changedWhereClause <> "" Then
                    Try
                        .Offset(2, 0).Value = changedWhereClause
                        .Offset(2, 0).WrapText = False
                    Catch ex As Exception
                        MsgBox("Error in adding where clause (" & changedWhereClause & ")", vbCritical, "DBSheet Creation Error")
                        lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
                        Exit Sub
                    End Try
                End If
            End With
            ' finally add the DBSetQuery for the main DB Mapper, only taking the query without the where clause (because we can't prefill the where parameters, 
            ' the user has to do that before extending the query definition to the where clause as well)
            ConfigFiles.createFunctionsInCells(curCell, {"RC", "=DBSetQuery(R[1]C,"""",RC[1])"})
            ' finish creation in async called function (need to have the results from the above calculations)
            ExcelAsyncUtil.QueueAsMacro(Sub()
                                            finishDBMapperCreation()
                                        End Sub)
        End If
    End Sub

    Private Sub finishDBMapperCreation()
        ' store lookup columns (<>LU) to be ignored in DBMapper
        Dim ignoreColumns As String = ""
        Try
            If Not IsNothing(lookupsList) Then
                ' replace fieldname of Lookups in DBMapper with fieldname + "LU"
                For Each LookupDef As String In lookupsList
                    If getEntry("fkey", LookupDef, 1) <> "" Then 'only necessary for database lookups
                        Dim lookupName As String = Replace(getEntry("name", LookupDef, 1), "*", "")
                        If IsNothing(ExcelDnaUtil.Application.Range(lookupName & "Lookup").Cells(1, 1).Value) Then
                            Dim answr As MsgBoxResult = MsgBox("lookup area '" & lookupName & "Lookup' probably contains no values (maybe an error), continue?", vbCritical + vbOKCancel, "DBSheet Creation Error")
                            If answr = vbCancel Then
                                lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
                                Exit Sub
                            End If
                        End If
                        ' create dropdown (validation) for lookup column
                        curCell.Offset(2 + addedCells, 0).Formula = "=OFFSET(" & lookupName & "Lookup,0,0,,1)" ' this is necessary as Formula1 in Validation.Add doesn't accept english formulas
                        Dim lookupColumn As Excel.ListColumn
                        Try
                            lookupColumn = createdListObject.ListColumns(lookupName & "LU")
                        Catch ex As Exception
                            MsgBox("lookup column '" & lookupName & "LU' not found in ListRange", vbCritical, "DBSheet Creation Error")
                            lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
                            Exit Sub
                        End Try
                        ' this is necessary as Excel>=2016 introduces the @operator automatically in formulas, referring to just that value in the same row. which is undesired here..
                        Dim localOffsetFormula As String = Replace(curCell.Offset(2 + addedCells, 0).FormulaLocal, "@", "")
                        Try
                            lookupColumn.DataBodyRange.Validation.Add(
                            Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop, Operator:=Excel.XlFormatConditionOperator.xlBetween,
                            Formula1:=localOffsetFormula)
                        Catch ex As Exception
                            MsgBox("Error in adding validation formula " & localOffsetFormula & " to column '" & lookupName & "LU': " & ex.Message, vbCritical, "DBSheet Creation Error")
                            lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
                            Exit Sub
                        End Try
                        ' as the listobject was automatically extended by the setting of the new column (looked up value) above, the resolution formulas go into the last column now.
                        Dim newCol As Excel.ListColumn = createdListObject.ListColumns.Add()
                        ' add vlookup function field for resolution of lookups to ID in main Query at the end of the DBMapper table
                        newCol.Name = lookupName
                        Dim lookupFormula As String = "=IF([@[" & lookupName & "LU]]<>"""",VLOOKUP([@[" & lookupName & "LU]]" & "," & lookupName & "Lookup" & ",2,False),"""")"
                        Try
                            newCol.DataBodyRange.Formula = lookupFormula
                        Catch ex As Exception
                            lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
                            MsgBox("Error in adding lookup formula " & lookupFormula & " to new column " & lookupName & ": " + ex.Message, vbCritical, "DBSheet Creation Error")
                            Exit Sub
                        End Try
                        ' hide the resolution column
                        newCol.Range.EntireColumn.Hidden = True
                        ' add lookup column to ignored columns (only resolution column will be stored in DB)
                        ignoreColumns += lookupName + "LU,"
                        curCell.Offset(2 + addedCells, 0).Formula = ""
                    End If
                Next
                ignoreColumns = Left(ignoreColumns, ignoreColumns.Length - 1)
            End If
        Catch ex As Exception
            MsgBox("Error in DBSheet Creation: " + ex.Message, vbCritical, "DBSheet Creation Error")
            lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
            Exit Sub
        End Try

        ' set DBMapper Rangename
        Dim NamesList As Excel.Names = ExcelDnaUtil.Application.ActiveWorkbook.Names
        Dim alreadyExists As Boolean = False
        Try
            Dim testExist As String = NamesList.Item("DBMapper" & tableName).ToString
        Catch ex As Exception
            alreadyExists = True
        End Try
        If Not alreadyExists Then
            MsgBox("Error adding DBModifier 'DBMapper" & tableName & "', Name already exists in Workbook!", vbCritical, "DBSheet Creation Error")
            Exit Sub
        End If
        Try
            NamesList.Add(Name:="DBMapper" & tableName, RefersTo:=curCell.Offset(0, 1))
        Catch ex As Exception
            MsgBox("Error when assigning name 'DBMapper" & tableName & "' to DBMapper starting cell (one cell to the right of active cell): " & ex.Message, vbCritical, "DBSheet Creation Error")
            Exit Sub
        End Try
        ' get the database name
        Dim databaseName As String = Replace(getEntry("connID", curConfig), fetchSetting("connIDPrefixDBtype", "MSSQL"), "")
        ' primary columns count (first <primCols> columns are primary columns)s
        Dim primCols = getEntry("primcols", curConfig)
        ' create DBMapper for DBSheet
        Dim CustomXmlParts As Object = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
        If CustomXmlParts.Count = 0 Then ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.Add("<root xmlns=""DBModifDef""></root>")
        CustomXmlParts = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
        ' NamespaceURI:="DBModifDef" is required to avoid adding a xmlns attribute to each element.
        CustomXmlParts(1).SelectSingleNode("/ns0:root").AppendChildNode("DBMapper" & tableName, NamespaceURI:="DBModifDef")
        Dim dbModifNode As CustomXMLNode = CustomXmlParts(1).SelectSingleNode("/ns0:root/ns0:DBMapper" & tableName)
        dbModifNode.AppendChildNode("execOnSave", NamespaceURI:="DBModifDef", NodeValue:="True")
        dbModifNode.AppendChildNode("askBeforeExecute", NamespaceURI:="DBModifDef", NodeValue:="True")
        dbModifNode.AppendChildNode("env", NamespaceURI:="DBModifDef", NodeValue:="")
        dbModifNode.AppendChildNode("database", NamespaceURI:="DBModifDef", NodeValue:=databaseName)
        dbModifNode.AppendChildNode("tableName", NamespaceURI:="DBModifDef", NodeValue:=tableName)
        dbModifNode.AppendChildNode("primKeysStr", NamespaceURI:="DBModifDef", NodeValue:=primCols)
        dbModifNode.AppendChildNode("insertIfMissing", NamespaceURI:="DBModifDef", NodeValue:="True")
        dbModifNode.AppendChildNode("executeAdditionalProc", NamespaceURI:="DBModifDef", NodeValue:="")
        dbModifNode.AppendChildNode("ignoreColumns", NamespaceURI:="DBModifDef", NodeValue:=ignoreColumns)
        dbModifNode.AppendChildNode("CUDFlags", NamespaceURI:="DBModifDef", NodeValue:="True")
        dbModifNode.AppendChildNode("IgnoreDataErrors", NamespaceURI:="DBModifDef", NodeValue:="False")
        'get new definitions into ribbon right now...
        getDBModifDefinitions()
        ' extend Datarange for new DBMappers immediately after definition...
        DirectCast(Globals.DBModifDefColl("DBMapper").Item("DBMapper" & tableName), DBMapper).extendDataRange(ignoreCUDFlag:=True)
    End Sub

    ''' <summary>fetches value in entryMarkup within XMLString, search starts optionally at position startSearch (default 1)</summary>
    ''' <param name="entryMarkup"></param>
    ''' <param name="XMLString"></param>
    ''' <param name="startSearch">start position for search, position of end of entryMarkup is returned here, allowing iterative fetching of multiple entryMarkup elements (see getEntryList)</param>
    ''' <returns>the fetched value</returns>
    Public Function getEntry(entryMarkup As String, XMLString As String, Optional ByRef startSearch As Integer = 1) As String
        Dim markStart As String, markEnd As String
        Dim fetchBeg, fetchEnd As Integer

        On Error GoTo getEntry_Err
        If Len(XMLString) = 0 Then
            getEntry = vbNullString
            Exit Function
        End If

        markStart = "<" & entryMarkup & ">"
        markEnd = "</" & entryMarkup & ">"

        fetchEnd = startSearch
        fetchBeg = InStr(fetchEnd, XMLString, markStart)
        If fetchBeg = 0 Then
            getEntry = vbNullString
            Exit Function
        End If
        fetchEnd = InStr(fetchBeg, XMLString, markEnd)
        startSearch = fetchEnd
        getEntry = Mid$(XMLString, fetchBeg + Len(markStart), fetchEnd - (fetchBeg + Len(markStart)))
        Exit Function

getEntry_Err:
        LogError("Error: " & Err.Description & ", line " & Erl() & " in DBSheetConfig.getEntry")
    End Function

    ''' <summary>fetches entryMarkup parts contained within lists demarked by listMarkup within parentMarkup inside XMLString</summary>
    ''' <param name="parentMarkup"></param>
    ''' <param name="listMarkup"></param>
    ''' <param name="entryMarkup">element inside listMarkup that should be fetched, if empty take whole listMarkup instead</param>
    ''' <param name="XMLString"></param>
    ''' <param name="fetchListMarkup">if true, take listMarkup elements where entryMarkup was found, else take entryMarkup element</param>
    ''' <returns>list containing parts, if entryMarkup = vbNullString then list contains parts demarked by listMarkup</returns>
    Public Function getEntryList(parentMarkup As String, listMarkup As String, entryMarkup As String, XMLString As String, Optional fetchListMarkup As Boolean = False) As Object
        Dim list() As String = Nothing
        Dim i As Long, posEnd As Long
        Dim isFilled As Boolean
        Dim parentEntry As String, ListItem As String, part As String

        On Error GoTo getEntryList_Err
        If Len(XMLString) = 0 Then
            getEntryList = Nothing
            Exit Function
        End If

        i = 0 : posEnd = 1 : isFilled = False
        parentEntry = getEntry(parentMarkup, XMLString)
        Do
            ' first get outer element demarked by listMarkup
            ListItem = getEntry(listMarkup, XMLString, posEnd)
            If Len(entryMarkup) = 0 Then
                part = ListItem
            Else
                ' take inner element for check and returning (if fetchListMarkup not set)
                part = getEntry(entryMarkup, ListItem)
            End If
            If Len(part) > 0 Then
                ' take outer element where entryMarkup was found
                If fetchListMarkup Then part = ListItem
                isFilled = True
                ReDim Preserve list(i)
                list(i) = part
                i += 1
            End If
        Loop Until ListItem = ""
        If isFilled Then
            getEntryList = list
        Else
            getEntryList = Nothing
        End If
        Exit Function

getEntryList_Err:
        LogError("Error: " & Err.Description & ", line " & Erl() & " in DBSheetConfig.getEntryList")
    End Function


End Module

