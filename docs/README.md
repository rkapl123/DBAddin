## DBAddin

DBAddin is a ExcelDNA based Automation Add-in for Database interoperability (currently only the DB Functions for database querying are provided, see [DBFuncs](DBFuncs.md)).
DB Functions are an alternative to the Excel built-in MSQuery, which is integrated statically into the worksheet and has some limitations in terms of querying possibilities and flexibility of constructing parameterized queries (MS-Query allows parameterized queries only in simple queries that can be displayed graphically).
DB Functions also include the possibility for filling "data bound" controls (ComboBoxes and Listboxes) with data from queries. 
Other useful functions for easier creation of queries and working with database data are included as well.

DBAddin has been tested extensively (actually it's in production) only with Excel XP/2010/2016 and MS-SQLserver, other databases (MySQL, Oracle, PostgreSQL, DB2 and Sybase SQLserver) have just been tested with the associated Testworkbook "DBFuncsTest.xls".

To use that Testworkbook you'll need the pubs database, where I have scripts available for Oracle, Sybase, DB2, PostgreSQL and MySql [here](PUBS_database_scripts.zip) (the MS-SQLserver version can be downloaded [here](https://www.microsoft.com/en-us/download/details.aspx?id=23654)).  

DBAddin is distributed under the [GNU Public License V3](http://www.gnu.org/copyleft/gpl.html).
