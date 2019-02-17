' createTableViewConfigs.vbs
'
' creates standard xcl DBAddin config files for all tables and views in a database
' the name of the database is taken from the directory this script is executed from, so
' if called from within "_mydatabase" or ".mydatabase" will produce all configfiles for tables/views in mydatabase
'
' connect using connectionStr...
connectionStr = "driver=SQL Server;server=OEBFADB00;Trusted_Connection=Yes;database="

dim dbscnn
DBName = Left(Wscript.ScriptFullName, instrrev(Wscript.ScriptFullName,"\")-1)
Wscript.echo DBName 
Set objShell = CreateObject("Wscript.Shell")
objShell.CurrentDirectory = DBName 
DBName = MID(DBName, instrrev(DBName,"\")+2)

Set fso = CreateObject("Scripting.FileSystemObject")
on error resume next
fso.DeleteFile("*.xcl")
fso.DeleteFile("*.wrd")
on error goto 0

if openConn = False then Wscript.Quit
Set rstSchema = CreateObject("ADODB.RecordSet")
rstSchema.Open "SELECT * FROM INFORMATION_SCHEMA.TABLES",dbscnn, 3, 1, 1
Do Until rstSchema.EOF
	if rstSchema.Fields("TABLE_NAME") <> "dtproperties" Then
		Set tf = fso.CreateTextFile(rstSchema.Fields("TABLE_NAME") & ".xcl", True)
		tf.WriteLine("RC" & chr(9) & "=DBListFetch(""Select TOP 10000 * FROM " & DBName & ".." & rstSchema.Fields("TABLE_NAME") & ""","""",R[1]C,,,TRUE,TRUE,TRUE)") 
		tf.Close
		Set tf = fso.CreateTextFile(rstSchema.Fields("TABLE_NAME") & ".wrd", True)
		tf.WriteLine("DATABASE \c """ & connectionStr & DBName & """ \s ""Select TOP 10000 * FROM " & DBName & ".." & rstSchema.Fields("TABLE_NAME") & """ \h \* MERGEFORMAT")
		tf.Close
	end if
  rstSchema.MoveNext
Loop
rstSchema.Close

wscript.echo "XCL and WRD Configs created for all Tables/Views of " & DBName & " ..."

Function openConn()
  Set dbscnn = CreateObject("ADODB.Connection")
  dbscnn.ConnectionString = connectionStr & DBName
  dbscnn.ConnectionTimeout = 15
  On Error resume next
  dbscnn.Open
  	If err <> 0 then
  		retval = MsgBox(Error(Err.Number), vbOKOnly, "Error connecting to DB...")
  		openConn = False
  	Else
  		openConn = True
  	End If
End Function
