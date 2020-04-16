Imports System.IO
Imports System.Windows.Forms
Imports ExcelDna.Integration
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop


'''<summary>Helper module  for easier manipulation of DBSheet definition / Connection configuration data</summary> 
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
    ''' <summary>the database name</summary>
    Dim databaseName As String
    ''' <summary>the Database table name of the DBSheet</summary>
    Dim tableName As String
    ''' <summary>counter to know how many cells we filled for the dbmapper query 
    ''' (at least 2: dbsetquery function and query string, if additional where clause exists, 
    ''' add one for where clause, then one for each parameter)
    ''' </summary>
    Dim addedCells As Integer
    Dim tblPlaceHolder As String
    Dim specialNonNullableChar As String
    ''' <summary>for DBSheetCreateForm, store the password once so we don't have to enter it again...</summary>
    Public existingPwd As String
    ''' <summary>public clipboard row for DBSheet definition rows (foreign lookup info)</summary>
    Public clipboardDataRow As DBSheetDefRow


    Public Sub createDBSheet()
        Dim openFileDialog1 As OpenFileDialog = New OpenFileDialog With {
            .InitialDirectory = fetchSetting("DBSheetDefinitions" + Globals.env, ""),
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
            tblPlaceHolder = fetchSetting("tblPlaceHolder" + env.ToString, "!T!")
            specialNonNullableChar = fetchSetting("specialNonNullableChar" + env.ToString, "*")
            databaseName = Replace(getEntry("connID", curConfig), fetchSetting("connIDPrefixDBtype", "MSSQL"), "")
            ' get the table name of the DBSheet for setting the DBMapper name
            tableName = getEntry("table", curConfig)
            ' if database is contained in table name, only get rightmost identifier as table name..
            If InStr(tableName.ToLower, databaseName.ToLower + ".") > 0 Then tableName = Strings.Mid(tableName, InStrRev(tableName, ".") + 1)
            ' get query
            Dim queryStr As String = getEntry("query", curConfig)
            If queryStr = "" Then
                ErrorMsg("No query found in DBSheetConfig.", "DBSheet Creation Error")
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
                        changedWhereClause += If(i = 0, """WHERE ", "&""") + whereParts(i) + """&R[" + (i + 1).ToString + "]C"
                    End If
                Next
                queryStr = Replace(queryStr, "WHERE " + whereClause, "")
                ' whereClause inserted below queryStr
                addedCells += 1
            End If
            ' get lookup fields in complete columns definitions
            lookupsList = getEntryList("columns", "field", "lookup", curConfig, True)
            Dim selectPart As String = Left(queryStr, InStr(queryStr, "FROM ") - 1)
            Dim selectPartModif As String = selectPart ' select part with appending LU to lookups
            If Not IsNothing(lookupsList) Then
                lookupWS = ExcelDnaUtil.Application.ActiveWorkbook.Worksheets.Add()
                Dim lookupWSname As String = Left(Replace(Guid.NewGuid().ToString, "-", ""), 31)
                Try
                    lookupWS.Name = lookupWSname
                Catch ex As Exception
                    ErrorMsg("Error setting lookup Worksheet Name to '" + lookupWSname + "': " + ex.Message, "DBSheet Creation Error")
                    lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
                    Exit Sub
                End Try
                ' add lookup Queries in separate sheet
                'TODO:check if the same lookup already exists and skip creation to avoid duplicates that influence each other...
                Dim lookupCol As Integer = 1
                For Each LookupDef As String In lookupsList
                    ' fetch Lookupquery and get rid of template table def
                    Dim LookupQuery As String = Replace(getEntry("lookup", LookupDef, 1), tblPlaceHolder, "LT")
                    Dim lookupName As String = Replace(getEntry("name", LookupDef, 1), specialNonNullableChar, "")
                    ' replace fieldname of Lookups in query with fieldname + "LU" only for database lookups
                    If getEntry("fkey", LookupDef, 1) <> "" Then
                        ' replace looked up ID names with ID name + "LU" in query string
                        Dim foundDelim As Integer
                        For Each delimStr As String In {",", vbCrLf}
                            foundDelim = InStr(selectPartModif, " " + lookupName + delimStr)
                            If foundDelim > 0 Then
                                selectPartModif = Replace(selectPartModif, " " + lookupName + delimStr, " " + lookupName + "LU" + delimStr)
                                Exit For
                            End If
                            foundDelim = InStr(selectPartModif, "." + lookupName + delimStr)
                            If foundDelim > 0 Then
                                selectPartModif = Replace(selectPartModif, "." + lookupName + delimStr, "." + lookupName + " " + lookupName + "LU" + delimStr)
                                Exit For
                            End If
                        Next
                        If foundDelim = 0 Then
                            ErrorMsg("Error in changing lookupName '" + lookupName + "' to '" + lookupName + "LU' in select statement of DBSheet query, it has to begin with blank and end with ','blank or CrLf !", "DBSheet Creation Error")
                            lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
                            Exit Sub
                        End If
                        lookupWS.Cells(1, lookupCol + 1).Value = LookupQuery
                        lookupWS.Cells(1, lookupCol + 1).WrapText = False
                        lookupWS.Cells(2, lookupCol).Name = lookupName + "Lookup"
                        ' then create the DBListFetch with the lookup query 
                        ConfigFiles.createFunctionsInCells(lookupWS.Cells(1, lookupCol), {"RC", "=DBListFetch(RC[1],""""," + lookupName + "Lookup" + ")"})
                        ' database lookups have two columns
                        lookupCol += 2

                    Else
                        'simple value lookup (one column), no need to resolve to an ID
                        If InStr(LookupQuery, "||") > 0 Then ' fixed values separated by ||
                            Dim lrow As Integer
                            Dim lookupValues As String() = Split(LookupQuery, "||")
                            For lrow = 0 To UBound(lookupValues)
                                lookupWS.Cells(2 + lrow, lookupCol).value = lookupValues(lrow)
                            Next
                            lookupWS.Range(lookupWS.Cells(2, lookupCol), lookupWS.Cells(2 + lrow - 1, lookupCol)).Name = lookupName + "Lookup"
                            ' fixed value lookups have only one column
                            lookupCol += 1
                        Else ' single column DB lookup
                            lookupWS.Cells(1, lookupCol + 1).Value = LookupQuery
                            lookupWS.Cells(1, lookupCol + 1).WrapText = False
                            lookupWS.Cells(2, lookupCol).Name = lookupName + "Lookup"
                            ConfigFiles.createFunctionsInCells(lookupWS.Cells(1, lookupCol), {"RC", "=DBListFetch(RC[1],""""," + lookupName + "Lookup" + ")"})
                            ' single column DB lookups have two columns because of dbfunction and query definition in two cells..
                            lookupCol += 2
                        End If
                    End If
                Next
                lookupWS.Visible = Excel.XlSheetVisibility.xlSheetHidden
                curCell.Parent.Select()
            End If
            ' exchange the select part with the LU modified select part
            queryStr = Replace(queryStr, selectPart, selectPartModif)
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
                    ErrorMsg("Error in adding query (" + queryStr + ")", "DBSheet Creation Error")
                    lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
                    Exit Sub
                End Try
                ' add an additional where clause as a concatenation string
                If changedWhereClause <> "" Then
                    Try
                        .Offset(2, 0).Value = changedWhereClause
                        .Offset(2, 0).WrapText = False
                    Catch ex As Exception
                        ErrorMsg("Error in adding where clause (" + changedWhereClause + ")", "DBSheet Creation Error")
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
        Dim queryErrorPos As Integer = InStr(curCell.Value.ToString, "Error")
        If queryErrorPos > 0 Then
            ErrorMsg("DBSheet Query had an error:" + vbCrLf + Mid(curCell.Value.ToString, queryErrorPos + Len("Error in query table refresh: ")), "DBSheet Creation Error")
            If Not IsNothing(lookupWS) Then lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
            Exit Sub
        End If
        ' name the worksheet to tableName, if defined in the settings
        If CBool(fetchSetting("DBsheetAutoName", "False")) Then
            Try
                curCell.Parent.Name = Left(tableName, 31)
            Catch ex As Exception
                ErrorMsg("DBSheet setting worksheet name to '" + Left(tableName, 31) + "', error:" + ex.Message, "DBSheet Creation Error")
            End Try
        End If
        ' some visual aid for DBSheets
        If curCell.Column = 1 And curCell.Row = 1 Then curCell.EntireColumn.ColumnWidth = 0.4
        Dim ignoreColumns As String = ""
        Try
            If Not IsNothing(lookupsList) Then
                For Each LookupDef As String In lookupsList
                    Dim lookupName As String = Replace(getEntry("name", LookupDef, 1), specialNonNullableChar, "")
                    If IsNothing(ExcelDnaUtil.Application.Range(lookupName + "Lookup").Cells(1, 1).Value) Then
                        Dim answr As MsgBoxResult = QuestionMsg("lookup area '" + lookupName + "Lookup' probably contains no values (maybe an error), continue?", MsgBoxStyle.OkCancel, "DBSheet Creation Error")
                        If answr = vbCancel Then
                            lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
                            Exit Sub
                        End If
                    End If

                    ' ..... create dropdown (validation) for lookup column
                    ' a workaround with getting the local formula is necessary as Formula1 in Validation.Add doesn't accept english formulas
                    curCell.Offset(2 + addedCells, 0).Formula = "=OFFSET(" + lookupName + "Lookup,0,0,,1)"
                    ' necessary as Excel>=2016 introduces the @operator automatically in formulas referring to list objects, referring to just that value in the same row. which is undesired here..
                    Dim localOffsetFormula As String = Replace(curCell.Offset(2 + addedCells, 0).FormulaLocal.ToString, "@", "")
                    ' get lookupColumn (lookupName + "LU" for 2-column database lookups, lookupName only for 1-column lookups)
                    Dim lookupColumn As Excel.ListColumn
                    Try
                        ' only for 2-column database lookups add LU
                        If getEntry("fkey", LookupDef, 1) <> "" Then
                            lookupColumn = createdListObject.ListColumns(lookupName + "LU")
                        Else
                            lookupColumn = createdListObject.ListColumns(lookupName)
                        End If
                    Catch ex As Exception
                        ErrorMsg("lookup column '" + lookupName + "LU' not found in ListRange", "DBSheet Creation Error")
                        lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
                        Exit Sub
                    End Try
                    ' add validation to look columns
                    Try
                        ' if nothing was fetched, there is no DataBodyRange, so add validation to the second row of the column range...
                        If IsNothing(lookupColumn.DataBodyRange) Then
                            lookupColumn.Range.Cells(2, 1).Validation.Delete ' remove existing validations, just in case it exists, otherwise add would throw exception... 
                            lookupColumn.Range.Cells(2, 1).Validation.Add(
                                Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop, Operator:=Excel.XlFormatConditionOperator.xlBetween,
                                Formula1:=localOffsetFormula)
                        Else
                            lookupColumn.DataBodyRange.Validation.Delete()  ' remove existing validations, just in case it exists, otherwise add would throw exception... 
                            lookupColumn.DataBodyRange.Validation.Add(
                                Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop, Operator:=Excel.XlFormatConditionOperator.xlBetween,
                                Formula1:=localOffsetFormula)
                        End If
                    Catch ex As Exception
                        ErrorMsg("Error in adding validation formula " + localOffsetFormula + " to column '" + lookupName + "LU': " + ex.Message, "DBSheet Creation Error")
                        lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
                        Exit Sub
                    End Try
                    ' adding resolution formulas is only necessary for 2-column database lookups
                    If getEntry("fkey", LookupDef, 1) <> "" Then
                        ' add vlookup function field for resolution of lookups to ID in main Query at the end of the DBMapper table
                        Dim lookupFormula As String = "=IF([@[" + lookupName + "LU]]<>"""",VLOOKUP([@[" + lookupName + "LU]]" + "," + lookupName + "Lookup" + ",2,False),"""")"
                        ' if no data was fetched, add a row...
                        If IsNothing(createdListObject.DataBodyRange) Then createdListObject.ListRows.AddEx()
                        ' now add the resolution formula column
                        Dim newCol As Excel.ListColumn = createdListObject.ListColumns.Add()
                        newCol.Name = lookupName
                        Try
                            newCol.DataBodyRange.Formula = lookupFormula
                        Catch ex As Exception
                            lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
                            ErrorMsg("Error in adding lookup formula " + lookupFormula + " to new column " + lookupName + ": " + ex.Message, "DBSheet Creation Error")
                            Exit Sub
                        End Try
                        ' hide the resolution formula column
                        newCol.Range.EntireColumn.Hidden = True
                        ' add lookup column to ignored columns (only resolution column will be stored in DB)
                        ignoreColumns += lookupName + "LU,"
                    End If
                    ' clean up our workaround target...
                    curCell.Offset(2 + addedCells, 0).Formula = ""
                Next
                If ignoreColumns.Length > 0 Then ignoreColumns = Left(ignoreColumns, ignoreColumns.Length - 1)
            End If
        Catch ex As Exception
            ErrorMsg("Error in DBSheet Creation: " + ex.Message, "DBSheet Creation Error")
            If Not IsNothing(lookupWS) Then lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
            Exit Sub
        End Try
        ' remove autofilter...
        createdListObject.ShowAutoFilter = False
        ' set DBMapper Rangename
        Dim NamesList As Excel.Names = ExcelDnaUtil.Application.ActiveWorkbook.Names
        Dim alreadyExists As Boolean = False
        Try
            Dim testExist As String = NamesList.Item("DBMapper" + tableName).ToString
        Catch ex As Exception
            alreadyExists = True
        End Try
        If Not alreadyExists Then
            ErrorMsg("Error adding DBModifier 'DBMapper" + tableName + "', Name already exists in Workbook!", "DBSheet Creation Error")
            If Not IsNothing(lookupWS) Then lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
            Exit Sub
        End If
        Try
            NamesList.Add(Name:="DBMapper" + tableName, RefersTo:=curCell.Offset(0, 1))
        Catch ex As Exception
            ErrorMsg("Error when assigning name 'DBMapper" + tableName + "' to DBMapper starting cell (one cell to the right of active cell): " + ex.Message, "DBSheet Creation Error")
            If Not IsNothing(lookupWS) Then lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
            Exit Sub
        End Try
        ' primary columns count (first <primCols> columns are primary columns)s
        Dim primCols As String = getEntry("primcols", curConfig)
        Try
            ' some visual aid for DBSHeets
            If curCell.Column = 1 And curCell.Row = 1 Then
                ' freeze top row and primary column(s) if more than one column...
                curCell.Offset(1, If(createdListObject.ListColumns.Count > 1, 1 + CInt(primCols), 0)).Select()
                ExcelDnaUtil.Application.ActiveWindow.FreezePanes = True
            End If
        Catch ex As Exception
            ErrorMsg("Exception: " + ex.Message, "DBSheet Creation Error")
            If Not IsNothing(lookupWS) Then lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
            Exit Sub
        End Try

        ' create DBMapper Configuration for DBSheet
        Dim CustomXmlParts As Object = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
        If CustomXmlParts.Count = 0 Then ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.Add("<root xmlns=""DBModifDef""></root>")
        CustomXmlParts = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
        ' NamespaceURI:="DBModifDef" is required to avoid adding a xmlns attribute to each element.
        CustomXmlParts(1).SelectSingleNode("/ns0:root").AppendChildNode("DBMapper" + tableName, NamespaceURI:="DBModifDef")
        Dim dbModifNode As CustomXMLNode = CustomXmlParts(1).SelectSingleNode("/ns0:root/ns0:DBMapper" + tableName)
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
        DBModifs.getDBModifDefinitions()
        ' extend Datarange for new DBMappers immediately after definition...
        DirectCast(Globals.DBModifDefColl("DBMapper").Item("DBMapper" + tableName), DBMapper).extendDataRange()
        ' switch back to DBAddin tab for easier handling...
        Globals.theRibbon.ActivateTab("DBaddinTab")
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
            getEntry = ""
            Exit Function
        End If

        markStart = "<" + entryMarkup + ">"
        markEnd = "</" + entryMarkup + ">"

        fetchEnd = startSearch
        fetchBeg = InStr(fetchEnd, XMLString, markStart)
        If fetchBeg = 0 Then
            getEntry = ""
            Exit Function
        End If
        fetchEnd = InStr(fetchBeg, XMLString, markEnd)
        startSearch = fetchEnd
        getEntry = Strings.Mid(XMLString, fetchBeg + Len(markStart), fetchEnd - (fetchBeg + Len(markStart)))
        Exit Function

getEntry_Err:
        ErrorMsg("Error: " + Err.Description + " in DBSheetConfig.getEntry")
    End Function

    ''' <summary>creates markup with setting value content in entryMarkup, used in DBSheetCreateForm.xmlDbsheetConfig</summary>
    ''' <param name="entryMarkup"></param>
    ''' <param name="content"></param>
    ''' <returns>the markup</returns>
    Public Function setEntry(ByVal entryMarkup As String, ByVal content As String) As String
        setEntry = "<" + entryMarkup + ">" + content + "</" + entryMarkup + ">"
    End Function

    ''' <summary>fetches entryMarkup parts contained within lists demarked by listMarkup within parentMarkup inside XMLString</summary>
    ''' <param name="parentMarkup"></param>
    ''' <param name="listMarkup"></param>
    ''' <param name="entryMarkup">element inside listMarkup that should be fetched, if empty take whole listMarkup instead</param>
    ''' <param name="XMLString"></param>
    ''' <param name="fetchListMarkup">if true, take listMarkup elements where entryMarkup was found, else take entryMarkup element</param>
    ''' <returns>list containing parts, if entryMarkup = "" then list contains parts demarked by listMarkup</returns>
    Public Function getEntryList(parentMarkup As String, listMarkup As String, entryMarkup As String, XMLString As String, Optional fetchListMarkup As Boolean = False) As Object
        Dim list() As String = Nothing
        Dim i As Long, posEnd As Long
        Dim parentEntry As String, ListItem As String, part As String

        If Len(XMLString) = 0 Then
            getEntryList = Nothing
            Exit Function
        End If

        i = 0 : posEnd = 1
        Try
            parentEntry = getEntry(parentMarkup, XMLString)
            Do
                ' first get outer element demarked by listMarkup
                ListItem = getEntry(listMarkup, XMLString, posEnd)
                If entryMarkup = "" Then
                    part = ListItem
                Else
                    ' take inner element for check and returning (if fetchListMarkup not set)
                    part = getEntry(entryMarkup, ListItem)
                End If
                If Len(part) > 0 Then
                    ' take outer element where entryMarkup was found
                    If fetchListMarkup Then part = ListItem
                    ReDim Preserve list(i)
                    list(i) = part
                    i += 1
                End If
            Loop Until ListItem = ""
        Catch ex As Exception
            ErrorMsg("Exception in getEntryList: " + ex.Message)
        End Try
        getEntryList = list
    End Function

End Module


