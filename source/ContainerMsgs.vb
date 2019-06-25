Imports Microsoft.Office.Interop

''' <summary>Provides a data structure for transporting information from the calling function to calculation event procedure in DBFuncEventHandler the key for the calc msg container</summary>
Public Class ContainerCalcMsgs
    ' used in all DB functions:
    ''' <summary>the query used for the calling function</summary>
    Public Query As String
    ''' <summary>the calling function's cell</summary>
    Public caller As Excel.Range
    ''' <summary>the calling function's worksheet</summary>
    Public callsheet As Excel.Worksheet
    ''' <summary>calling ID</summary>
    Public callID As String
    ''' <summary>the connection string used for the calling function</summary>
    Public ConnString As String
    ''' <summary>to check whether we're currently working on that container</summary>
    Public working As Boolean
    ''' <summary>to check whether an error occurred</summary>
    Public errOccured As Boolean
    ''' <summary>should headers be added on top of list/records?</summary>
    Public HeaderInfo As Boolean

    ' only used in DBCellFetch:
    ''' <summary>DBCellFetch: the usual column separator</summary>
    Public colSep As String
    ''' <summary>DBCellFetch: the usual row separator</summary>
    Public rowSep As String
    ''' <summary>DBCellFetch: the last column separator</summary>
    Public lastColSep As String
    ''' <summary>DBCellFetch: the last row separator</summary>
    Public lastRowSep As String
    ''' <summary>DBCellFetch: should header infos be interleaved with values (fieldname1 fieldvalue1, fieldname2 fieldvalue2,...)</summary>
    Public InterleaveHeader As Boolean

    ' only used in DBListFetch:
    ''' <summary>DBListFetch: Range to put the data into</summary>
    Public targetRange As Excel.Range
    ''' <summary>DBListFetch: Range to copy formulas down from</summary>
    Public formulaRange As Excel.Range
    ''' <summary>DBListFetch: how to deal with extending List Area</summary>
    Public extendArea As Integer
    ''' <summary>DBListFetch: should 1st row formats be autofilled down?</summary>
    Public autoformat As Boolean
    ''' <summary>DBListFetch: should row numbers be displayed in 1st column?</summary>
    Public ShowRowNumbers As Boolean
    ''' <summary>DBListFetch: the name of the target area as a string</summary>
    Public targetRangeName As String
    ''' <summary>DBListFetch: the name of the target formula area as a string</summary>
    Public formulaRangeName As String
    ''' <summary>DBListFetch: should columns/control be autofitted ?</summary>
    Public AutoFit As Boolean

    ' only used in DBRowFetch
    ''' <summary>DBRowFetch: Range to put the data into</summary>
    Public targetArray As Object
End Class

''' <summary>Provides a data structure for transporting information back from the calculation event procedure in DBFuncEventHandler to the calling function</summary>
Public Class ContainerStatusMsgs
    ''' <summary>any status msg used for displaying in the result of function</summary>
    Public statusMsg As String
End Class
