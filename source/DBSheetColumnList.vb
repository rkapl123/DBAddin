Imports System.Windows.Forms

''' <summary>Helper class for easier manipulation of DBsheetCols Listview</summary>
Friend Class DBSheetColumnList
	' is wrapped by DBSheetColumnList to simplify access to it
	Private theCols As DataGridView

	Public Sub New(theDBsheetCols As DataGridView)
		theCols = theDBsheetCols
	End Sub

	Public Sub Clear()
		theCols.Rows.Clear()
	End Sub

	ReadOnly Property RowCount() As Integer
		Get
			Return theCols.Rows.Count
		End Get
	End Property

	ReadOnly Property ColumnCount() As Integer
		Get
			Return theCols.Columns.Count
		End Get
	End Property

	ReadOnly Property Headers() As String()
		Get
			Return {"name", "ftable", "fkey", "flookup", "outer", "primkey", "type", "sort", "lookup"}
		End Get
	End Property


	Property Selection() As Integer
		Get
			If theCols.CurrentCell Is Nothing Then
				Return -1
			Else
				Return theCols.CurrentCell.RowIndex + 1
			End If
		End Get
		Set(ByVal Value As Integer)
			If Value >= 0 And Value <= RowCount Then
				theCols.Rows(Value).Selected = True
			Else
				theCols.ClearSelection()
			End If
		End Set
	End Property

	''
	' Property "Value" gets column definition values in row/col with val
	' @param row
	' @param col
	' @return the value
	''
	' Property "Value" updates column definition values in row/col with val, (0/1) boolean values are transformed to "Y"(True)/""(False)
	' @param row
	' @param col
	' @param val
	Property Value(ByVal row As Integer, ByVal col As Integer) As Object
		Get
			Dim result As Object
			result = theCols.Item(col, row).Value
			Select Case col
				Case 4 To 6 : result = IIf(CStr(result) = "Y", 1, 0)
			End Select
			Return result
		End Get
		Set(ByVal Value As Object)
			If RowCount > 0 And row >= 0 Then
				Select Case col
					Case 4 To 6 : Value = IIf(CInt(Value) = 1 Or CStr(Value) = "Y", "Y", "")
				End Select
				theCols.Item(col, row).Value = CStr(Value)
			End If
		End Set
	End Property


	Public Function hasRows() As Boolean
		Return theCols.Rows.Count > 0
	End Function

	Public Function firstEntrySelected() As Boolean
		Return theCols.CurrentCell.RowIndex = 1
	End Function

	Public Function lastEntrySelected() As Boolean
		Return theCols.CurrentCell.RowIndex = theCols.Rows.Count
	End Function

	Public Function selectionMade() As Boolean
		Return Not (theCols.CurrentCell Is Nothing)
	End Function

	Public Function newRow() As Integer
		Dim currRow As DataGridViewRow = New DataGridViewRow()
		Return theCols.Rows.Add(currRow)
	End Function

	Public Sub removeRow(ByRef index As Integer)
		If hasRows() Then theCols.Rows.RemoveAt(index)
	End Sub

	''' <summary>checks in existing columns (column 1 in DataGridView theCols) whether theColumnVal exists already in DBsheetCols And returns the found row of DBsheetCols</summary>
	''' <param name="theColumnVal"></param>
	''' <returns>found row in DataGridView</returns>
	Public Function checkForValue(theColumnVal As String) As Integer
		theCols.SelectionMode = DataGridViewSelectionMode.FullRowSelect
		Try
			For Each row As DataGridViewRow In theCols.Rows
				If (row.Cells(2).Value.ToString().Equals(theColumnVal)) Then Return row.Index
			Next
		Catch ex As Exception
			ErrorMsg(ex.Message)
		End Try
		Return -1
	End Function

	''
	' "shifts" entries in dbsheet column definitions up(dir=1) or down (dir=-1) by
	'           exchanging neighboring entries.
	' @param dir
	Public Sub shiftEntry(ByRef dir As Integer)
		Dim temp As String

		Dim curSel As Integer = Selection
		If curSel >= 0 Then
			For j As Integer = 0 To ColumnCount - 1
				temp = Value(curSel + dir, j)
				Value(curSel + dir, j) = Value(curSel, j)
				Value(curSel, j) = temp
			Next
			curSel += dir
			Selection = curSel
		End If
	End Sub
End Class