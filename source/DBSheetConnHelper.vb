﻿Imports System.Data.Common
Imports System.Data.Odbc
Imports System.Data.SqlClient
Imports System.Data.OleDb

''' <summary>Helper Class for DBSheetCreateForm and AdHocSQL</summary>
Public Class DBSheetConnHelper
    ''' <summary>the connection string for dbsheet definitions, different from the normal one (extended rights for schema viewing required)</summary>
    Public dbsheetConnString As String
    ''' <summary>identifier needed to fetch database from connection string (eg Database=)</summary>
    Public dbidentifier As String
    ''' <summary>statement/procedure to get all databases in a DB instance</summary>
    Public dbGetAllStr As String
    ''' <summary>fieldname where databases are returned by dbGetAllStr</summary>
    Public DBGetAllFieldName As String
    ''' <summary>the DB connection for the dbsheet definition activities</summary>
    Public dbshcnn As DbConnection
    ''' <summary>identifier needed to put password into connection string (eg PWD=)</summary>
    Public dbPwdSpec As String

    Public Sub New()
        setConnectionString()
        dbGetAllStr = fetchSetting("dbGetAll" + Globals.env(), "NONEXISTENT")
        If dbGetAllStr = "NONEXISTENT" Then
            Globals.UserMsg("No dbGetAllStr given for environment: " + Globals.env() + ", please correct and rerun.", "DBSheet Definition Error")
            Exit Sub
        End If
        DBGetAllFieldName = fetchSetting("dbGetAllFieldName" + Globals.env(), "NONEXISTENT")
        If DBGetAllFieldName = "NONEXISTENT" Then
            Globals.UserMsg("No DBGetAllFieldName given for environment: " + Globals.env() + ", please correct and rerun.", "DBSheet Definition Error")
            Exit Sub
        End If
        dbidentifier = fetchSetting("DBidentifierCCS" + Globals.env(), "NONEXISTENT")
        If dbidentifier = "NONEXISTENT" Then
            Globals.UserMsg("No DB identifier given for environment: " + Globals.env() + ", please correct and rerun.", "DBSheet Definition Error")
            Exit Sub
        End If
        dbPwdSpec = fetchSetting("dbPwdSpec" + Globals.env(), "")
    End Sub

    ''' <summary>opens a database connection with active connstring</summary>
    Public Sub openConnection(Optional databaseName As String = "")
        ' connections are pooled by ADO depending on the connection string:
        If dbshcnn Is Nothing Or databaseName <> "" Then
            setConnectionString()
            ' add password to connection string
            If InStr(1, dbsheetConnString, dbPwdSpec) > 0 And dbPwdSpec <> "" Then
                If Strings.Len(existingPwd) > 0 Then
                    If InStr(1, dbsheetConnString, dbPwdSpec) > 0 Then
                        dbsheetConnString = Globals.Change(dbsheetConnString, dbPwdSpec, existingPwd, ";")
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
                    dbsheetConnString = Globals.Change(dbsheetConnString, dbidentifier, databaseName, ";")
                Else
                    dbsheetConnString = dbsheetConnString + ";" + dbidentifier + databaseName
                End If
            End If
            ' need to change/set the connection timeout in the connection string as the property is readonly then...
            If InStr(dbsheetConnString, "Connection Timeout=") > 0 Then
                dbsheetConnString = Globals.Change(dbsheetConnString, "Connection Timeout=", Globals.CnnTimeout.ToString(), ";")
            ElseIf InStr(dbsheetConnString, "Connect Timeout=") > 0 Then
                dbsheetConnString = Globals.Change(dbsheetConnString, "Connect Timeout=", Globals.CnnTimeout.ToString(), ";")
            Else
                dbsheetConnString += ";Connection Timeout=" + Globals.CnnTimeout.ToString()
            End If
        End If
        Dim correctConnString As String = ""
        Try
            If InStr(dbsheetConnString.ToLower, "provider=sqloledb") Or InStr(dbsheetConnString.ToLower, "driver=sql server") Then
                ' ADO.NET doesn't like provider= and driver= 
                If Globals.fetch(dbsheetConnString, "provider=", ";", True) <> "" Then
                    correctConnString = Replace(dbsheetConnString, Globals.fetch(dbsheetConnString, "provider=", ";", True) + ";", "")
                End If
                If Globals.fetch(correctConnString, "driver=", ";", True) <> "" Then
                    correctConnString = Replace(correctConnString, Globals.fetch(correctConnString, "driver=", ";", True) + ";", "")
                End If
                dbshcnn = New SqlConnection(correctConnString)
            ElseIf InStr(dbsheetConnString.ToLower, "oledb") Then
                dbshcnn = New OleDbConnection(dbsheetConnString)
            Else
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
        dbsheetConnString = fetchSetting("DBSheetConnString" + Globals.env(), "NONEXISTENT")
        If dbsheetConnString = "NONEXISTENT" Then
            ' no, try normal connection string 
            dbsheetConnString = fetchSetting("ConstConnString" + Globals.env(), "NONEXISTENT")
            If dbsheetConnString = "NONEXISTENT" Then
                ' actually this cannot happen....
                Globals.UserMsg("No Connectionstring given for environment: " + Globals.env() + ", please correct and rerun.", "DBSheet Definition Error")
                Exit Sub
            End If
        End If
    End Sub
End Class