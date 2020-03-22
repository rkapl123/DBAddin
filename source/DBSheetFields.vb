Imports System.Collections.Specialized
Friend Class DBSheetFields

	'  Helper class for DBSheetHandler to make DBSheet field (columns) handling easier

	Public lookUpsFilled As Boolean
	Private mFieldsByCol As OrderedDictionary
	Private mColumnsByFld As OrderedDictionary
	Private mIsForeignLookup As OrderedDictionary

	Public Sub New()
		mFieldsByCol = New OrderedDictionary()
		mColumnsByFld = New OrderedDictionary()
		mIsForeignLookup = New OrderedDictionary()
	End Sub

	ReadOnly Property Count() As Integer
		Get
			Return mColumnsByFld.Count
		End Get
	End Property

	ReadOnly Property Column(ByVal aField As String) As Integer
		Get
			Return CInt(mColumnsByFld(aField))
		End Get
	End Property

	ReadOnly Property Columns() As Object
		Get
			Return mColumnsByFld
		End Get
	End Property

	ReadOnly Property Field(ByVal aColumn As Integer) As Object
		Get
			Return CStr(mFieldsByCol(aColumn - 1))
		End Get
	End Property

	ReadOnly Property Fields() As Object
		Get
			Return mFieldsByCol
		End Get
	End Property

	ReadOnly Property IsForeignLookup(ByVal aField As String) As Boolean
		Get
			Return CBool(mIsForeignLookup(aField))
		End Get
	End Property

End Class