## DBAddin

DBAddin is a ExcelDNA based Automation Add-in for Database interoperability, see [DBFuncs Userdoc](DBFuncs.md) and [DBFuncs API-doc](api/index.html).  

DB Functions are an alternative to the Excel built-in MSQuery, which is integrated statically into the worksheet and has some limitations in terms of querying possibilities and flexibility of constructing parameterized queries (MS-Query allows parameterized queries only in simple queries that can be displayed graphically).  

Other useful functions for easier creation of queries and a feature for working with database data (DBMapper) are included as well.  

This is the successor to the COM based DBAddin (as found on [sourceforge](https://sourceforge.net/projects/dbaddin/)).

Testing for MS SQL Server and other databases (MySQL, Oracle, PostgreSQL, DB2, Sybase and Access) can be done using the Testing Workbook "DBFuncsTest.xls".
To use that Testing Workbook you'll need the pubs database, where I have scripts available for Oracle, Sybase, DB2, PostgreSQL and MySql [here](PUBS_database_scripts.zip) (the MS-SQLserver version can be downloaded [here](https://www.microsoft.com/en-us/download/details.aspx?id=23654)).  
I've also added a pubs.mdb Access database in the test folder.

DBAddin is distributed under the [GNU Public License V3](http://www.gnu.org/copyleft/gpl.html).
