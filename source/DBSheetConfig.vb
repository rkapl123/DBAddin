Imports ExcelDna.Integration
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Windows.Forms

'''<summary>Helper module for creation of DBSheets and DBSheet definition manipulation helpers</summary> 
Public Module DBSheetConfig
    ''' <summary>the current cell where the DBSheet Definition is inserted at</summary>
    Public curCell As Excel.Range
    ''' <summary>the list object of the main query for the db mapper</summary>
    Public createdListObject As Excel.ListObject
    ''' <summary>the lookups list of the DBSheet definition (xml element with query, name, etc.)</summary>
    Dim lookupsList() As String
    ''' <summary>the complete db-sheet configuration (XML)</summary>
    Dim curConfig As String
    ''' <summary>the added and hidden worksheet with lookups inside</summary>
    Public lookupWS As Excel.Worksheet
    ''' <summary>the database name</summary>
    Dim databaseName As String
    ''' <summary>the Database table name of the DBSheet</summary>
    Dim tableName As String
    ''' <summary>counter to know how many cells we filled for the db-mapper query 
    ''' (at least 2: dbsetquery function and query string, if an additional where clause exists, 
    ''' add one for this where clause and then one for each parameter)
    ''' </summary>
    Dim addedCells As Integer
    ''' <summary>these three need to be global, so that finishDBMapperCreation also knows about them</summary>
    Dim whereClauseStart, tblPlaceHolder, specialNonNullableChar As String
    ''' <summary>for DBSheetCreateForm, store the password once so we don't have to enter it again...</summary>
    Public existingPwd As String
    ''' <summary>public clipboard row for DBSheet definition rows (foreign lookup info)</summary>
    Public clipboardDataRow As DBSheetDefRow
    ''' <summary>if an existing DBSheet is overwritten, this is set to the existing DBModifier Name</summary>
    Public existingName As String


    ''' <summary>create a DBSheet by creating lookups (with dblistfetch) and a dbsetquery that acts as a list-object for a CUD DBMapper. Called by clickAssignDBSheet (Ribbon) and assignDBSheet_Click (DBSheetCreateForm)</summary>
    Public Sub createDBSheet(Optional dbsheetDefPath As String = "")
        If ExcelDnaUtil.Application.ActiveWorkbook.Windows(1).WindowState = Excel.XlWindowState.xlMinimized Then
            UserMsg("No assignment possible when active workbook is minimized!", "DBSheet Creation Error")
            Exit Sub
        End If
        ' store currently selected cell, where DBSetQuery for DBMapper will be placed.
        curCell = ExcelDnaUtil.Application.ActiveCell
        existingName = getDBModifNameFromRange(curCell)
        If InStr(1, existingName, "DBMapper") > 0 OrElse (UCase(Left(curCell.Formula, 11)) = "=DBSETQUERY" And InStr(1, getDBModifNameFromRange(curCell.Offset(0, 1)), "DBMapper") > 0) Then
            Dim answer As MsgBoxResult = QuestionMsg("Existing DBSheet detected in selected area, shall this be overwritten?", MsgBoxStyle.OkCancel)
            If answer = MsgBoxResult.Cancel Then Exit Sub
            If UCase(Left(curCell.Formula, 11)) <> "=DBSETQUERY" Then
                ' either dbsetquery (needed for curCell) is to the cell to the left
                curCell = ExcelDnaUtil.Application.Range(existingName).Cells(1, 1).Offset(0, -1)
            Else
                ' or the db-mapper area (needed for existingName for removing definitions and db-mapper name) is to the right
                existingName = getDBModifNameFromRange(curCell.Offset(0, 1))
            End If
        End If
        Dim openFileDialog1 As OpenFileDialog = Nothing
        Dim result As DialogResult
        If dbsheetDefPath = "" Then
            ' ask for DBsheet definitions stored in xml file
            openFileDialog1 = New OpenFileDialog With {
                .InitialDirectory = fetchSetting("lastDBsheetAssignPath", fetchSetting("DBSheetDefinitions" + env(), "")),
                .Filter = "XML files (*.xml)|*.xml",
                .RestoreDirectory = True
            }
            result = openFileDialog1.ShowDialog()
            If result = Windows.Forms.DialogResult.OK Then dbsheetDefPath = openFileDialog1.FileName
        End If

        If dbsheetDefPath <> "" Then
            setUserSetting("lastDBsheetAssignPath", Strings.Left(dbsheetDefPath, InStrRev(dbsheetDefPath, "\") - 1))
            ' Get the DBSheet Definition file name and read into curConfig
            curConfig = File.ReadAllText(dbsheetDefPath, Text.Encoding.Default)
            tblPlaceHolder = fetchSetting("tblPlaceHolder" + env.ToString(), "!T!")
            specialNonNullableChar = fetchSetting("specialNonNullableChar" + env.ToString, "*")
            databaseName = Replace(getEntry("connID", curConfig), fetchSetting("connIDPrefixDBtype", "MSSQL"), "")
            ' get the table name of the DBSheet for setting the DBMapper name
            tableName = getEntry("table", curConfig)
            ' if database is contained in table name, only get rightmost identifier as table name..
            If InStr(tableName.ToLower, databaseName.ToLower + ".") > 0 Then tableName = Strings.Mid(tableName, InStrRev(tableName, ".") + 1)
            ' get query
            Dim queryStr As String = getEntry("query", curConfig)
            If queryStr = "" Then
                UserMsg("No query found in DBSheetConfig !", "DBSheet Creation Error")
                Exit Sub
            End If
            If QuestionMsg("Should TOP 100 be put into query, in case of very large underlying tables this helps in creating the DBSheet (you can restrict the query later on) ?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                queryStr = "SELECT TOP 100 " + queryStr.Substring(7) ' skip "SELECT " in queryStr
            End If
            whereClauseStart = queryStr.IndexOf("WHERE", StringComparison.OrdinalIgnoreCase)
            ' queryStr inserted below DBSetQuery
            addedCells = 1
            Dim changedWhereClause As String = ""
            If whereClauseStart >= 0 Then
                Dim whereClause As String = queryStr.Substring(whereClauseStart)
                ' check for where clauses and modify for parameter setting in formula
                Dim lastCharParam As Boolean = (Strings.Right(whereClause, 1) = "?")
                Dim whereParts As String() = Split(whereClause, "?")
                For i = 0 To UBound(whereParts)
                    If whereParts(i) <> "" Then
                        ' each parameter adds a cell below DBSetQuery
                        addedCells += 1
                        ' create concatenation formula for parameter setting, each ? is replaced by a separate row reference below the where clause
                        changedWhereClause += If(i = 0, "=""", "&""") + whereParts(i) + If(i < UBound(whereParts) Or lastCharParam, """&R[" + (i + 1).ToString() + "]C", """")
                        ' before where Part: formula op (at begin) then concat operator ... afterwards correct R[refRow]C ref, at end only closing quote unless whereStr ended with "?"
                    End If
                Next
                ' remove in queryStr as where clause sits in a separate cell now (enhancement of DBSetquery query param to this cell to be done later by user)
                queryStr = Replace(queryStr, whereClause, "")
                ' whereClause inserted below queryStr
                addedCells += 1
            End If

            ' when doing any changes to existing db-sheets setting calc to manual is needed to avoid triggering unwanted recalculations
            Dim calcMode As Long = ExcelDnaUtil.Application.Calculation
            Try
                ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual
            Catch ex As Exception
                UserMsg("The Calculation mode can't be set, maybe you are in the formula/cell editor?", "Create Function In Cell")
                Exit Sub
            End Try

            ' get lookup fields in complete columns definitions
            lookupsList = getEntryList("columns", "field", "lookup", curConfig, True)
            Dim selectPart As String = Left(queryStr, InStr(queryStr, "FROM ") - 1)
            Dim selectPartModif As String = selectPart ' select part with appending LU to lookups
            If lookupsList IsNot Nothing Then
                ' get existing sheet DBSheetLookups, if it doesn't exist create it anew
                If Not existsSheet("DBSheetLookups", ExcelDnaUtil.Application.ActiveWorkbook) Then
                    lookupWS = ExcelDnaUtil.Application.ActiveWorkbook.Worksheets.Add()
                    lookupWS.Name = "DBSheetLookups"
                Else
                    Dim answer As MsgBoxResult = QuestionMsg("Existing DBSheetLookups sheet detected, should all lookup definitions be removed (if definitions with existing names but different meanings are added, this might lead to errors)?", MsgBoxStyle.YesNoCancel)
                    If answer = MsgBoxResult.Cancel Then Exit Sub
                    If answer = MsgBoxResult.Yes Then
                        ExcelDnaUtil.Application.Worksheets("DBSheetLookups").Cells.Clear
                        For Each LookupDef As String In lookupsList
                            Dim lookupRangeName As String = tableName + Replace(getEntry("name", LookupDef, 1), specialNonNullableChar, "") + "Lookup"
                            If existsName(lookupRangeName) Then
                                Try : ExcelDnaUtil.Application.Names(lookupRangeName).Delete : Catch ex As Exception : End Try
                            End If
                        Next
                    End If
                    lookupWS = ExcelDnaUtil.Application.Worksheets("DBSheetLookups")
                End If
                ' add lookup Queries in separate sheet
                Dim lookupCol As Integer = 1
                For Each LookupDef As String In lookupsList
                    ' fetch Lookup query and get rid of template table def
                    Dim LookupQuery As String = Replace(getEntry("lookup", LookupDef, 1), tblPlaceHolder, "LT")
                    Dim lookupName As String = Replace(getEntry("name", LookupDef, 1), specialNonNullableChar, "")
                    Dim lookupRangeName As String = tableName + lookupName + "Lookup"
                    If existsName(lookupRangeName) Then
                        ' overwrite existing lookup with warning...
                        lookupCol = lookupWS.Range(lookupRangeName).Column
                    Else
                        ' step to the right
                        If Not IsNothing(lookupWS.Cells(1, lookupCol).Value) Then
                            If Not IsNothing(lookupWS.Cells(1, lookupCol + 1).Value) Then
                                lookupCol = lookupWS.Cells(1, lookupCol).End(Excel.XlDirection.xlToRight).Column + 1
                            Else
                                lookupCol += 1
                            End If
                        End If
                    End If
                    ' replace field-name of Lookups in query with field-name + "LU" only for database lookups
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
                            UserMsg("Error in preparing lookupName '" + lookupName + "' (conversion to '" + lookupName + "LU') in select statement of DBSheet query:" + vbCrLf + selectPart + vbCrLf + "The fieldname part always has to begin with blank and end with ',' or a newline (CrLf)!", "DBSheet Creation Error")
                            lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
                            Exit Sub
                        End If
                        lookupWS.Cells(1, lookupCol + 1).Value = LookupQuery
                        lookupWS.Cells(1, lookupCol + 1).WrapText = False
                        ' only create name and dblistfetch if lookup doesn't already exist!
                        If Not existsName(lookupRangeName) Then
                            lookupWS.Cells(2, lookupCol).Name = lookupRangeName
                            ' then create the DBListFetch with the lookup query
                            createFunctionsInCells(lookupWS.Cells(1, lookupCol), {"RC", "=DBListFetch(RC[1],""""," + lookupRangeName + ")"})
                        Else
                            LogWarn("DB Sheet Lookup " + lookupRangeName + " already exists in " + lookupWS.Range(lookupRangeName).Address + ", check if this is really the correct one !")
                        End If
                    Else
                        'simple value lookup (one column), no need to resolve to an ID
                        If InStr(LookupQuery, "||") > 0 Then ' fixed values separated by ||
                            Dim lrow As Integer
                            Dim lookupValues As String() = Split(LookupQuery, "||")
                            For lrow = 0 To UBound(lookupValues)
                                lookupWS.Cells(2 + lrow, lookupCol).value = lookupValues(lrow)
                            Next
                            ' add the name, so there is something in the top row (for moving to right...
                            lookupWS.Cells(1, lookupCol).Value = lookupRangeName
                            ' only create name and dblistfetch if lookup doesn't already exist!
                            If Not existsName(lookupRangeName) Then
                                ' fixed value lookups have only one column
                                lookupWS.Range(lookupWS.Cells(2, lookupCol), lookupWS.Cells(2 + lrow - 1, lookupCol)).Name = lookupRangeName
                            Else
                                LogWarn("DB Sheet Lookup " + lookupRangeName + " already exists in " + lookupWS.Range(lookupRangeName).Address + ", check if this is really the correct one !")
                            End If
                        Else ' single column DB lookup
                            lookupWS.Cells(1, lookupCol + 1).Value = LookupQuery
                            lookupWS.Cells(1, lookupCol + 1).WrapText = False
                            ' only create name and dblistfetch if lookup doesn't already exist!
                            If Not existsName(lookupRangeName) Then
                                lookupWS.Cells(2, lookupCol).Name = lookupRangeName
                                createFunctionsInCells(lookupWS.Cells(1, lookupCol), {"RC", "=DBListFetch(RC[1],""""," + lookupRangeName + ")"})
                            Else
                                LogWarn("DB Sheet Lookup " + lookupRangeName + " already exists in " + lookupWS.Range(lookupRangeName).Address + ", check if this is really the correct one !")
                            End If
                        End If
                    End If
                Next
                lookupWS.Visible = Excel.XlSheetVisibility.xlSheetHidden
                curCell.Parent.Select()
            End If
            ' exchange the select part with the LU modified select part
            queryStr = Replace(queryStr, selectPart, selectPartModif)
            ' then create the DBSetQuery assigning the (yet to be filled) query to the above list-object
            ' add DBSetQuery with queryStr as Basis for the final DBMapper
            ' if DBMapper already exists remove everything to allow recreating DBSheets
            If existingName <> "" Then
                Try
                    If curCell.Column = 1 And curCell.Row = 1 Then curCell.EntireColumn.ColumnWidth = 10 ' reset minimized column, otherwise list object looks stupid.
                    curCell.Offset(0, 1).ListObject.Delete()
                    ExcelDnaUtil.Application.ActiveWindow.FreezePanes = False ' remove the freeze-pane, it will be applied later again.
                    ExcelDnaUtil.Application.ActiveWorkbook.Names(existingName).Delete
                    Dim theCustomXmlParts As Object = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
                    ' remove old node of DBMapper in definitions
                    If Not IsNothing(theCustomXmlParts(1).SelectSingleNode("/ns0:root/ns0:DBMapper[@Name='" + Replace(existingName, "DBMapper", "") + "']")) Then theCustomXmlParts(1).SelectSingleNode("/ns0:root/ns0:DBMapper[@Name='" + Replace(existingName, "DBMapper", "") + "']").Delete

                    ' just in case the same definition was added again, remove query cache to force re-fetching 
                    ' get the callID of the underlying name of the target (key of the queryCache and StatusCollection)
                    Dim callID As String = "[" + curCell.Parent.Parent.Name + "]" + curCell.Parent.Name + "!" + curCell.Address
                    If queryCache.ContainsKey(callID) Then queryCache.Remove(callID)
                Catch ex As Exception
                    UserMsg("Error deleting existing list-object for DBSheet for table " + tableName + ": " + ex.Message, "DBSheet Creation Error")
                    Exit Sub
                End Try
            End If
            ' create a ListObject
            createdListObject = createListObject(curCell)
            If createdListObject Is Nothing Then Exit Sub

            With curCell
                ' add the query as text
                Try
                    .Offset(1, 0).Value = queryStr
                    .Offset(1, 0).WrapText = False ' avoid row height increasing
                Catch ex As Exception
                    UserMsg("Error in adding query (" + queryStr + "): " + ex.Message, "DBSheet Creation Error")
                    lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
                    Exit Sub
                End Try
                ' add the additional where clause as a concatenation string
                If changedWhereClause <> "" Then
                    Try
                        .Offset(2, 0).Value = changedWhereClause
                        .Offset(2, 0).WrapText = False
                    Catch ex As Exception
                        UserMsg("Error in adding where clause (" + changedWhereClause + "): " + ex.Message, "DBSheet Creation Error")
                        lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
                        Exit Sub
                    End Try
                End If
            End With
            ' finally add the DBSetQuery for the main DB Mapper, only taking the query without the where clause (because we can't prefill the where parameters, 
            ' the user has to do that before extending the query definition to the where clause as well)
            ' also set calc to automatic back to trigger recalculation after leaving with QueueAsMacro

            createFunctionsInCells(curCell, {"RC", "=DBSetQuery(R[1]C,"""",RC[1])"})
            ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
            ' finish creation in async called sub (need to have the results from the above createFunctionsInCells/invocations)
            ExcelAsyncUtil.QueueAsMacro(Sub()
                                            finishDBMapperCreation()
                                        End Sub)
        End If
    End Sub

    ''' <summary>after creating lookups and setting the dbsetquery finish the list-object area with reverse lookups and drop-downs</summary>
    Private Sub finishDBMapperCreation()

        ' store lookup columns (<>LU) to be ignored in DBMapper
        Dim queryErrorPos As Integer = InStr(curCell.Value.ToString(), "Error")
        If queryErrorPos > 0 Then
            UserMsg("DBSheet Query had an error:" + vbCrLf + Mid(curCell.Value.ToString(), queryErrorPos + Len("Error in query table refresh: ")), "DBSheet Creation Error")
            If lookupWS IsNot Nothing Then lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
            Exit Sub
        End If
        ' name the worksheet to tableName, if defined in the settings
        If fetchSettingBool("DBSheetAutoName", "False") Then
            Try
                curCell.Parent.Name = Left(tableName, 31) ' prevent errors due to long names
            Catch ex As Exception
                UserMsg("DBSheet setting worksheet name to '" + Left(tableName, 31) + "', error:" + ex.Message, "DBSheet Creation Error")
            End Try
        End If
        ' for "full" DBSheets, minimize first column as much as possible
        If curCell.Column = 1 And curCell.Row = 1 Then curCell.EntireColumn.ColumnWidth = 0.4
        Dim ignoreColumns As String = ""
        Try
            If lookupsList IsNot Nothing Then
                For Each LookupDef As String In lookupsList
                    Dim lookupName As String = Replace(getEntry("name", LookupDef, 1), specialNonNullableChar, "")
                    Dim lookupRangeName As String = tableName + lookupName + "Lookup"
                    ' check if both columns of the lookup are empty (lookup key can be empty but the value may not) to check for empty lookup query results
                    If IsNothing(ExcelDnaUtil.Application.Range(lookupRangeName).Cells(1, 1).Value) And IsNothing(ExcelDnaUtil.Application.Range(lookupRangeName).Cells(1, 2).Value) Then
                        Dim answr As MsgBoxResult = QuestionMsg("lookup area '" + lookupRangeName + "' contains no values (maybe an error), continue?", MsgBoxStyle.OkCancel, "DBSheet Creation Error")
                        If answr = vbCancel Then
                            lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
                            Exit Sub
                        End If
                    End If

                    ' ..... create dropdown (validation) for lookup column
                    ' workaround by setting formula in a temporary cell to get the local language formula. This is necessary as Formula1 in Validation.Add doesn't accept English formulas
                    curCell.Offset(2 + addedCells, 0).Formula = "=OFFSET(" + lookupRangeName + ",0,0,,1)"
                    ' necessary as Excel>=2016 introduces the @operator automatically in formulas referring to list objects, referring to just that value in the same row. which is undesired here..
                    Dim localOffsetFormula As String = Replace(curCell.Offset(2 + addedCells, 0).FormulaLocal.ToString(), "@", "")
                    ' get lookupColumn (lookupName + "LU" for 2-column database lookups, lookupName only for 1-column lookups)
                    Dim lookupColumn As Excel.ListColumn
                    Dim finalLookupname = ""
                    Try
                        ' only for 2-column database lookups add LU
                        finalLookupname = If(getEntry("fkey", LookupDef, 1) <> "", lookupName + "LU", lookupName)
                        lookupColumn = createdListObject.ListColumns(finalLookupname)
                    Catch ex As Exception
                        UserMsg("lookup column " + finalLookupname + " not found in ListRange", "DBSheet Creation Error")
                        lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
                        Exit Sub
                    End Try
                    ' add validation to look columns
                    Try
                        ' if nothing was fetched, there is no DataBodyRange, so add validation to the second row of the column range...
                        If IsNothing(lookupColumn.DataBodyRange) Then
                            lookupColumn.Range.Cells(2, 1).Validation.Delete() ' remove existing validations, just in case it exists, otherwise add would throw exception... 
                            lookupColumn.Range.Cells(2, 1).Validation.Add(
                                Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop, Operator:=Excel.XlFormatConditionOperator.xlEqual, Formula1:=localOffsetFormula)
                        Else
                            lookupColumn.DataBodyRange.Validation.Delete()   ' remove existing validations, just in case it exists, otherwise add would throw exception... 
                            lookupColumn.DataBodyRange.Validation.Add(
                                Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop, Operator:=Excel.XlFormatConditionOperator.xlEqual, Formula1:=localOffsetFormula)
                        End If
                    Catch ex As Exception
                        UserMsg("Error in adding validation formula " + localOffsetFormula + " to column " + lookupColumn.Name + ": " + ex.Message, "DBSheet Creation Error")
                        lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
                        Exit Sub
                    End Try
                    ' adding resolution formulas is only necessary for 2-column database lookups
                    If getEntry("fkey", LookupDef, 1) <> "" Then
                        ' add vlookup function field for resolution of lookups to ID in main Query at the end of the DBMapper table
                        Dim lookupFormula As String = "=IF([@[" + lookupName + "LU]]<>"""",IF(ISERROR(VLOOKUP([@[" + lookupName + "LU]]," + lookupRangeName + ",2,False)),[@[" + lookupName + "LU]],VLOOKUP([@[" + lookupName + "LU]]," + lookupRangeName + ",2,False)),"""")"
                        ' if no data was fetched, add a row...
                        If IsNothing(createdListObject.DataBodyRange) Then createdListObject.ListRows.AddEx()
                        ' now add the resolution formula column
                        Dim newCol As Excel.ListColumn = createdListObject.ListColumns.Add()
                        newCol.Name = lookupName
                        Try
                            newCol.DataBodyRange.Formula = lookupFormula
                        Catch ex As Exception
                            lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
                            UserMsg("Error in adding lookup formula " + lookupFormula + " to new column " + lookupName + ": " + ex.Message, "DBSheet Creation Error")
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
            UserMsg("Error in DBSheet Creation: " + ex.Message, "DBSheet Creation Error")
            If lookupWS IsNot Nothing Then lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
            Exit Sub
        End Try
        ' remove auto-filter...
        createdListObject.ShowAutoFilter = False
        ' set DBMapper Range-name
        Dim NamesList As Excel.Names
        Try : NamesList = ExcelDnaUtil.Application.ActiveWorkbook.Names : Catch ex As Exception
            LogWarn("Exception when trying to get the active workbook names for checking if new DBMapper Range name exists: " + ex.Message + ", this might be either due to errors in the VBA-IDE (missing references) or due to opening this workbook from an MS-Office hyperlink, starting up Excel (timing issue). Switch to another workbook and back to fix.")
            Exit Sub
        End Try
        Dim alreadyExists As Boolean = True
        Try
            Dim testExist As String = NamesList.Item("DBMapper" + tableName).ToString()
        Catch ex As Exception
            ' exception only triggered if name not already exists !
            alreadyExists = False
        End Try
        If alreadyExists Then
            UserMsg("Error adding DBModifier 'DBMapper" + tableName + "', Name already exists in Workbook!", "DBSheet Creation Error")
            If lookupWS IsNot Nothing Then lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
            Exit Sub
        End If
        Try
            NamesList.Add(Name:="DBMapper" + tableName, RefersTo:=createdListObject.Range) ' curCell.Offset(0, 1) DBMapper starting cell (one cell to the right of active cell)
        Catch ex As Exception
            UserMsg("Error when assigning name 'DBMapper" + tableName + "' to DBSetQuery Target: " + ex.Message, "DBSheet Creation Error")
            If lookupWS IsNot Nothing Then lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
            Exit Sub
        End Try

        ' primary columns count (first <primCols> columns are primary columns)
        Dim primCols As String = getEntry("primcols", curConfig)
        Try
            ' some visual aid for DBSheets
            If curCell.Column = 1 And curCell.Row = 1 Then
                ' freeze top row and primary column(s) if more than one column...
                curCell.Offset(1, If(createdListObject.ListColumns.Count > 1, 1 + CInt(primCols), 0)).Select()
                ExcelDnaUtil.Application.ActiveWindow.FreezePanes = True
            End If
        Catch ex As Exception
            UserMsg("Exception: " + ex.Message, "DBSheet Creation Error")
            If lookupWS IsNot Nothing Then lookupWS.Visible = Excel.XlSheetVisibility.xlSheetVisible
            Exit Sub
        End Try

        ' create DBMapper Configuration for DBSheet
        Dim CustomXmlParts As Object = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
        If CustomXmlParts.Count = 0 Then ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.Add("<root xmlns=""DBModifDef""></root>")
        CustomXmlParts = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
        ' NamespaceURI:="DBModifDef" is required to avoid adding a xmlns attribute to each element.
        CustomXmlParts(1).SelectSingleNode("/ns0:root").AppendChildNode("DBMapper", NamespaceURI:="DBModifDef")
        ' new appended elements are last, get it to append further child elements
        Dim dbModifNode As CustomXMLNode = CustomXmlParts(1).SelectSingleNode("/ns0:root").LastChild
        ' append the detailed settings to the definition element
        dbModifNode.AppendChildNode("Name", NodeType:=MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, NodeValue:=tableName)
        dbModifNode.AppendChildNode("execOnSave", NamespaceURI:="DBModifDef", NodeValue:="True")
        dbModifNode.AppendChildNode("askBeforeExecute", NamespaceURI:="DBModifDef", NodeValue:="True")
        dbModifNode.AppendChildNode("env", NamespaceURI:="DBModifDef", NodeValue:="0")
        dbModifNode.AppendChildNode("database", NamespaceURI:="DBModifDef", NodeValue:=databaseName)
        dbModifNode.AppendChildNode("tableName", NamespaceURI:="DBModifDef", NodeValue:=tableName)
        dbModifNode.AppendChildNode("primKeysStr", NamespaceURI:="DBModifDef", NodeValue:=primCols)
        dbModifNode.AppendChildNode("insertIfMissing", NamespaceURI:="DBModifDef", NodeValue:="True")
        dbModifNode.AppendChildNode("executeAdditionalProc", NamespaceURI:="DBModifDef", NodeValue:="")
        dbModifNode.AppendChildNode("ignoreColumns", NamespaceURI:="DBModifDef", NodeValue:=ignoreColumns)
        dbModifNode.AppendChildNode("CUDFlags", NamespaceURI:="DBModifDef", NodeValue:="True")
        dbModifNode.AppendChildNode("IgnoreDataErrors", NamespaceURI:="DBModifDef", NodeValue:="False")

        'get new definitions into ribbon right now...
        getDBModifDefinitions(ExcelDnaUtil.Application.ActiveWorkbook)

        ' format non null-able fields specially, this needs to be after DB Mapper has been initialized (theme colors!)
        Dim existingHeaderColour As Integer = createdListObject.TableStyle.TableStyleElements(Excel.XlTableStyleElementType.xlHeaderRow).Interior.Color
        ' walk through all fields of DBSheet
        Dim fieldList() As String = getEntryList("columns", "field", "", curConfig, True)
        For Each fieldDef As String In fieldList
            Dim fieldName As String = getEntry("name", fieldDef)
            ' for non null-able fields ...
            If Left(fieldName, 1) = specialNonNullableChar Then
                fieldName = Replace(fieldName, specialNonNullableChar, "")
                ' with 2 column lookups the LU column is visible (actual resolved field column is hidden)
                If getEntry("fkey", fieldDef) <> "" Then fieldName += "LU"
                ' ... fill non-null field headers with darker pattern
                With createdListObject.ListColumns(fieldName).Range(1, 1).Interior
                    .Pattern = Excel.XlPattern.xlPatternGray25
                    .Color = existingHeaderColour
                End With
            End If
        Next
        ' avoid spill over from definition cells (query, where clause, etc.) into DBSheet area in case a row is inserted
        createdListObject.ListColumns.Item(1).Range.Offset(0, -1).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

        ' extend Data range for new DBMappers immediately after definition...
        DirectCast(DBModifDefColl("DBMapper").Item("DBMapper" + tableName), DBMapper).extendDataRange()
        ' switch back to DBAddin tab for easier handling...
        theRibbon.ActivateTab("DBaddinTab")
        curCell.Select()
        If whereClauseStart >= 0 Then
            UserMsg("Attention: A where clause was defined for this DBSheet, you need to extend the DBSetQuery function's Query argument in cell " + curCell.Address + "!", "DBSheet Creation")
        End If
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
        UserMsg("Error: " + Err.Description + " in DBSheetConfig.getEntry")
    End Function

    ''' <summary>creates markup with setting value content in entryMarkup, used in DBSheetCreateForm.xmlDbsheetConfig</summary>
    ''' <param name="entryMarkup"></param>
    ''' <param name="content"></param>
    ''' <returns>the markup</returns>
    Public Function setEntry(ByVal entryMarkup As String, ByVal content As String) As String
        setEntry = "<" + entryMarkup + ">" + content + "</" + entryMarkup + ">"
    End Function

    ''' <summary>fetches entryMarkup parts contained within lists denoted by listMarkup within parentMarkup inside XMLString</summary>
    ''' <param name="parentMarkup"></param>
    ''' <param name="listMarkup"></param>
    ''' <param name="entryMarkup">element inside listMarkup that should be fetched, if empty take whole listMarkup instead</param>
    ''' <param name="XMLString"></param>
    ''' <param name="fetchListMarkup">if true, take listMarkup elements where entryMarkup was found, else take entryMarkup element</param>
    ''' <returns>list containing parts, if entryMarkup = "" then list contains parts denoted by listMarkup</returns>
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
                ' first get outer element denoted by listMarkup
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
            UserMsg("Exception in getEntryList: " + ex.Message)
        End Try
        getEntryList = list
    End Function
End Module