Imports System.Data.Common
Imports System.Data.Odbc
Imports System.Data.SqlClient
Imports System.Data.OleDb

''' <summary>Helper Class for DBSheetCreateForm and AdHocSQL</summary>
Public Class DBConnHelper
    ''' <summary>the connection string for dbsheet definitions, different from the normal one (extended rights for schema viewing required)</summary>
    Public dbsheetConnString As String
    ''' <summary>identifier needed to fetch database from connection string (e.g Database=)</summary>
    Public dbidentifier As String
    ''' <summary>statement/procedure to get all databases in a DB instance</summary>
    Public dbGetAllStr As String
    ''' <summary>field-name where databases are returned by dbGetAllStr</summary>
    Public DBGetAllFieldName As String
    ''' <summary>the DB connection for the dbsheet definition activities</summary>
    Public dbshcnn As DbConnection
    ''' <summary>identifier needed to put password into connection string (e.g PWD=)</summary>
    Public dbPwdSpec As String
    ''' <summary>the environment of the DBConnHelper</summary>
    Public DBenv As String

    Public Sub New(passedEnv As String)
        DBenv = passedEnv
        setConnectionString()
        dbGetAllStr = fetchSetting("dbGetAll" + DBenv, "NONEXISTENT")
        If dbGetAllStr = "NONEXISTENT" Then
            UserMsg("No dbGetAllStr given for environment: " + DBenv + ", please correct and rerun.", "DBSheet Definition Error")
            Exit Sub
        End If
        DBGetAllFieldName = fetchSetting("dbGetAllFieldName" + DBenv, "NONEXISTENT")
        If DBGetAllFieldName = "NONEXISTENT" Then
            UserMsg("No DBGetAllFieldName given for environment: " + DBenv + ", please correct and rerun.", "DBSheet Definition Error")
            Exit Sub
        End If
        dbidentifier = fetchSetting("DBidentifierCCS" + DBenv, "NONEXISTENT")
        If dbidentifier = "NONEXISTENT" Then
            UserMsg("No DB identifier given for environment: " + DBenv + ", please correct and rerun.", "DBSheet Definition Error")
            Exit Sub
        End If
        dbPwdSpec = fetchSetting("dbPwdSpec" + DBenv, "")
    End Sub

    Public Function getCommand(commandText As String) As System.Data.Common.DbCommand
        If TypeName(dbshcnn) = "SqlConnection" Then
            getCommand = New SqlCommand(commandText, dbshcnn)
        ElseIf TypeName(dbshcnn) = "OleDbConnection" Then
            getCommand = New OleDbCommand(commandText, dbshcnn)
        Else
            getCommand = New OdbcCommand(commandText, dbshcnn)
        End If
    End Function

    ''' <summary>opens a database connection with active connection string</summary>
    Public Sub openConnection(Optional databaseName As String = "", Optional usedForDBSheetCreate As Boolean = False)
        ' connections are pooled by ADO depending on the connection string:
        If dbshcnn Is Nothing Or databaseName <> "" Then
            setConnectionString()
            ' add password to connection string
            If InStr(1, dbsheetConnString, dbPwdSpec) > 0 And dbPwdSpec <> "" Then
                If Strings.Len(existingPwd) > 0 Then
                    If InStr(1, dbsheetConnString, dbPwdSpec) > 0 Then
                        dbsheetConnString = Change(dbsheetConnString, dbPwdSpec, existingPwd, ";")
                    Else
                        dbsheetConnString = dbsheetConnString + ";" + dbPwdSpec + existingPwd
                    End If
                Else
                    Throw New Exception("Password is required by connection string: " + dbsheetConnString)
                End If
            End If
            ' add database name to connection string, needed for schema retrieval!!!
            If databaseName <> "" Then
                If InStr(1, dbsheetConnString.ToUpper, dbidentifier.ToUpper) > 0 Then
                    dbsheetConnString = Change(dbsheetConnString, dbidentifier, databaseName, ";")
                Else
                    dbsheetConnString = dbsheetConnString + ";" + dbidentifier + databaseName
                End If
            End If
            ' need to change/set the connection timeout in the connection string as the property is readonly then...
            If InStr(dbsheetConnString, "Connection Timeout=") > 0 Then
                dbsheetConnString = Change(dbsheetConnString, "Connection Timeout=", CnnTimeout.ToString(), ";")
            ElseIf InStr(dbsheetConnString, "Connect Timeout=") > 0 Then
                dbsheetConnString = Change(dbsheetConnString, "Connect Timeout=", CnnTimeout.ToString(), ";")
            Else
                dbsheetConnString += ";Connection Timeout=" + CnnTimeout.ToString()
            End If
        End If
        Dim correctConnString As String = ""
        Try
            If Not usedForDBSheetCreate And (InStr(dbsheetConnString.ToLower, "provider=sqloledb") Or InStr(dbsheetConnString.ToLower, "driver=sql server")) Then
                ' ADO.NET doesn't like provider= and driver= 
                If fetchSubstr(dbsheetConnString, "provider=", ";", True) <> "" Then
                    correctConnString = Replace(dbsheetConnString, fetchSubstr(dbsheetConnString, "provider=", ";", True) + ";", "")
                End If
                If fetchSubstr(correctConnString, "driver=", ";", True) <> "" Then
                    correctConnString = Replace(correctConnString, fetchSubstr(correctConnString, "driver=", ";", True) + ";", "")
                End If
                dbshcnn = New SqlConnection(correctConnString)
            ElseIf Not usedForDBSheetCreate And InStr(dbsheetConnString.ToLower, "oledb") Then
                dbshcnn = New OleDbConnection(dbsheetConnString)
            Else
                ' for DBSheetCreate dialog always use ODBC as performance is not important there but rather vendor compatibility...
                ' change to ODBC driver setting
                dbsheetConnString = Replace(dbsheetConnString, fetchSetting("ConnStringSearch" + env(), "provider=SQLOLEDB"), fetchSetting("ConnStringReplace" + env(), "driver=SQL SERVER"))
                dbshcnn = New OdbcConnection(dbsheetConnString)
            End If
            dbshcnn.Open()
        Catch ex As Exception
            dbsheetConnString = Replace(dbsheetConnString, dbPwdSpec + existingPwd, dbPwdSpec + "*******")
            dbshcnn = Nothing
            Throw New Exception("Error connecting to DB: " + ex.Message + ", connection string: " + dbsheetConnString + IIf(correctConnString <> "", " (corrected ConnString: " + correctConnString + ")", ""))
        End Try
    End Sub

    ''' <summary>set the dbSheet connection string, used in initialization and openConnection</summary>
    Private Sub setConnectionString()
        ' do we have a separate dbsheet connection string?
        dbsheetConnString = fetchSetting("DBSheetConnString" + DBenv, "NONEXISTENT")
        If dbsheetConnString = "NONEXISTENT" Then
            ' no, try normal connection string 
            dbsheetConnString = fetchSetting("ConstConnString" + DBenv, "NONEXISTENT")
            If dbsheetConnString = "NONEXISTENT" Then
                ' actually this cannot happen....
                UserMsg("No connection string given for environment: " + DBenv + ", please correct and rerun.", "DBSheet Definition Error")
                Exit Sub
            End If
        End If
    End Sub
End Class