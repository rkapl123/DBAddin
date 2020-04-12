## DBAddin

DBAddin is a ExcelDNA based Add-in for Database interoperability.

First, DBaddin provides DB Functions (see [DBFuncs Userdoc](DBFuncs.md)), which are an alternative to the Excel built-in MSQuery (integrated statically into worksheets having severe limitations in terms of querying and constructing parameterized queries (MS-Query allows parameterized queries only in simple queries that can be displayed graphically)).  

Next, methods for working with database data ([DBModifications](DBModif.md): DBMapper, DBAction and DBSequences) are included. This also includes a row entry oriented way to modify data in so called DBSheets (see [DBSheets](DBSheets.md)).

DBAddin.NET is the successor to the VB6 based DBAddin (as found on [sourceforge](https://sourceforge.net/projects/dbaddin/)).

Testing for MS SQL Server and other databases (MySQL, Oracle, PostgreSQL, DB2, Sybase and Access) can be done using the Testing Workbook "DBFuncsTest.xls".
To use that Testing Workbook you'll need the pubs database, where I have scripts available for Oracle, Sybase, DB2, PostgreSQL and MySql [here](PUBS_database_scripts.zip) (the MS-SQLserver version can be downloaded [here](https://www.microsoft.com/en-us/download/details.aspx?id=23654)). I've also added a pubs.mdb Access database in the test folder.

Generated API documentation: [DBFuncs API-doc](api/index.html).
