## DBSheets

DBSheets are an extension to DBMappers with CUD (Create/Update/Delete) flags to modify Database data directly in an Excel sheet, i.e. to insert, update and delete rows for a defined set of fields of a given table.
The modifications are done in DBMappers, which are filled by using a specified query. The DBMapper contains indirect lookup values for updating foreign key columns as well as resolution formulas to achieve the "normal" Table data (direct values).
The allowed IDs and the visible lookup values for those foreign key columns are stored in a hidden lookup sheet.

Work in DBSheets is done with context menus (right mouse) and shortcut keys to the context menu items:

![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/ContextMenuDBSheets.PNG)

*   New records are added by either adding data in an empty row below the displayed data or by selecting "insert Row (Ctl-Sh-I)" in the cell context menu and adding the data into the inserted empty line.
*   Existing records are updated by simply changing cells.
*   Existing records are deleted by selecting "delete Row (Ctl-Sh-D)" in the cell context menu.
*   When the Excel-Workbook is saved (Ctl-s) then the current DBSheet is stored as well.
*   Selecting "refresh Data (Ctl-Sh-R)" in the cell context menu will refresh data (undoing any changes made so far)

There is also a supporting tool "Create DBSheet definition" (Ribbon menu: DB Addin tools, DBsheet def) available for building and editing DBSheet definitions.

In the following sections, the major capabilities of DBSheets are presented, followed by a description of "Create DBSheet definition".

### Background

Prerequisites for understanding this documentation and using DBSheets is a basic proficiency with SQL and database design.
Good books on this topic are "The Practical SQL Handbook: Using Structured Query Language (3rd Edition)" and its successor "The Practical SQL Handbook: Using SQL Variants (4th Edition)", available free online courses are:  

*   [http://www.sql-und-xml.de/sql-tutorial/](http://www.sql-und-xml.de/sql-tutorial/) (German)
*   [http://www.w3schools.com/sql/default.asp](http://www.w3schools.com/sql/default.asp) (English)
*   [http://www.sql-tutorial.net/](http://www.sql-tutorial.net/) (English)
*   [http://sqlcourse.com](http://sqlcourse.com/) (English)
*   [http://sqlcourse2.com](http://sqlcourse2.com/) (English)

### Working with DBSheets (DBMappers with CUDFlags)

I use the enclosed the test workbook called "DBSheetsTest.xlsx" as an example to guide through the possibilities of DBSheets.

This Workbook uses the pubs database for MSSQL, (available for download/installation [from Microsoft](http://www.microsoft.com/downloads/details.aspx?FamilyId=06616212-0356-46A0-8DA2-EEBC53A68034&displaylang=en) or search "pubs download" on Microsoft.com, if this is not installed on your DB server already). For MySQL, Sybase, Oracle and DB2, I enclosed the pubs database myself in the test folder.

DBSheets consist of an Excel data table containing the table data, the header and - if any foreign lookup fields are present - the resolution formulas for those foreign lookups, and one (hidden) sheet containing the (foreign table) lookups.

![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/DBSheetInAction.PNG)

The header row and the primary key column(s), which are located leftmost, should not be modified, except for new rows where there is no primary key yet.

Lookup columns are restricted with excels cell validation, including a dropdown of allowed values.

Data is changed simply by editing existing data in cells, this marks the row(s) to be changed at the rightmost end of the data table.

Inserting is done by entering data into empty rows, either allowing the database to introduce a new primary key by leaving the primary column empty (only possible for singular primary keys) or by setting the primary key yourself. This marks the row(s) to be inserted.

Empty rows can be created within the existing data area by using context menu "insert row" or pressing Ctl-Sh-I in the row where data should be inserted. If multiple rows are selected then the same amount of empty rows is inserted.

Deleting is done using context menu "delete row" or pressing Ctl-Sh-D in the row to be deleted. This marks the row(s) to be deleted.

To make the editions permanent, save the workbook (save button or Ctl-s). After a potential warning message (as set in the DBMapper), the current DBSheet is stored to the database, producing warnings/errors as the database (or triggers) checks for validity.

### DBSheet definition file

The DBSheet definition file contains following information in XML format:

1.  the query for fetching the main table data to be edited
2.  the connection ID referring to the connection definition in the  global connection definition file. The connection definition contains the connection string, information on how parts of the connection string can be interpreted (database, password), how the collection of available databases can be retrieved and the windows users permitted to create/edit DBSheet definitions.
3.  The primary column count.
4.  all column definitions including the foreign key lookups. These consist of a lookup name, being the name of the column and either a select statement or a list of values. The select statement has to return exactly two columns returning the lookup values first and then the IDs to be looked up. (the main table's column value set should be contained in those, so every column value can be looked up).  
    Duplicates naturally should be strictly avoided in the return set of a query for a referential foreign key column as they would lead to ambiguities.

### Supporting Tool "Create DBSheet definition"

You start either entering a password for the connection string (if needed) or immediately selecting a database for your DBSheet. The connection informations are placed in the same central settings file (referenced by DBAddin.xll.config file attribute, in the example called "DBAddinCentral.config") as the other DBAddin connections. This settings file can be placed on a network drive to be available for all DBSheets users. You can also use the settings editor in the DBAddin ribbon to edit this file.

![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/DBSheetsDefBlank.PNG)

In case of successfully connecting to the database server, the dropdown "database" becomes available and you can proceed to selecting a database and afterwards a table.

![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/DBSheetsDefSelectDatabase.PNG)

"Load DBSheet def from File" is used to directly load a stored definition into the tool.

After having selected a database, select the main table for which a DBsheet should be created in the dropdown "Table", which fills the available fields of the table into the dropdowns of column "name". Once a field has been chosen, the password/database/table entry becomes unavailable. Only clearing ALL fields from the DBSheets definition will allow a change to the password/database/table entry again.

![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/DBSheetsDefSelectTableColumn.PNG)

After that you can start to populate the fields that should be edited by selecting them in the column "name".
A quick way to add all available columns is to click "add all Fields" (or pressing Alt-f)

If the field is required to be filled (non-nullable) then an special character (here: an asterisk) is put in front of it (shown also in the list of fields for choosing), the special character is removed however when generating/customizing queries and lookup queries). The first field is automatically set to be a primary key field, any subsequent fields that should serve as primary key can be marked as such by ticking the "primkey" column. Primary columns must always come first in a DBSheet, so any primary key column after a "non-primary key" column is prevented by DBsheet Creation.

If the field should be a lookup from a foreign table then select a foreign table in the ftable column:

![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/DBSheetsDefSelectForTable.PNG)

After selecting the foreign table, the key of the foreign table can be selected in the fkey column. This key is used to join the main table with the foreign table. In case it should be an outer join (allowing for missing entries in the foreign table), tick column "outer".

To finish foreign table lookup, select the Foreign Table Lookup Column in the flookup column, which serves as a meaningful description for the foreign key (usually some "Name", "Code" or "Description" field).

![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/DBSheetsDefSelectForLookupsGenerate.PNG)

After selecting the flookup column, the tool asks if the flookup statement, which is used to query the database for the key/lookup value pairs, should be created.

You can always edit the fields already stored in the DBSheet-Column list by selecting a line and changing the values in the dropdowns.

You can change the order of fields by selecting a row and using the context menu up/down buttons:
![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/DBSheetsDefFullDefinedMoveMenu.PNG)

You can copy/paste the foreign lookup definitions between fields by pressing Ctrl-C on a field to be copied and Ctrl-V on the field where the definitions should be pasted. Everything except the field name and type is pasted.

Removing columns is possible by simply deleting a row, you can clear the whole DBsheet definitions by clicking "reset DBSheet creation".

When dealing with foreign column lookups or other lookup restrictions (see below), you can edit the definition of the lookup directly by editing the lookup query field (testing is possible with the context menu in the foreign lookup column).

![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/DBSheetsDefFullDefinedContextMenu.PNG)

You can put any query into that, it just has to return the lookup value first and the ID to be looked up second. Duplicates should be strictly avoided in the return set of this query as they would lead to ambiguities and will produce error messages when generating the DBSheet.

Customizations of the restriction field should observe a few rules to be used efficiently (thereby not forcing the DBSheet creating person to do unnecessary double work): First, any reference to the foreign table itself has to use the configured template placeholder (here: !T!), which is then replaced by an actual table enumerator (T2..TN, T1 always being the primary table).  

Lookup value columns that differ from the table field's name must have an alias associated, which has to be the table field's name. If that is not the case, DBAddin won't be able to associate the foreign column in the main table with the lookup id, and thus displays following error message when creating the DBSheet query:

![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/DBSheetsDefForLookupErrorMsg.PNG)

The connection between the lookup queries and the main query is as follows:

1.  The first column part of the lookup query select statement is copied into the respective field in the main table query (therefore the above restriction)
2.  The foreign lookup table and all further additional tables needed for the lookup query are joined into the main query in the same way as they are defined in the lookup (inner/outer joins), `WHERE` clauses are added to those joins with `AND`.

Following diagram should clarify the connection between lookup queries and the main query:

![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/DBSheetsDefForLookupRelation.PNG)

You can always test the foreign lookup query by selecting the context menu item "test lookup query" in the lookup column. This opens an excel sheet with the results of the lookup query being inserted. This Testsheet can be closed again either by simply closing it, or by selecting context menu item "remove lookup query test".

You can also have a lookup field without defining a foreign table relation at all. This is done by simply defining the lookup query for that field. These lookup queries can either have one or two columns but only the **first column** is used for the lookup:

![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/DBSheetsDefLookupSelect.PNG)

Also remember that lookups always check for uniqueness, so in case there are duplicate lines to be expected, an additional "distinct" clause will avoid the consequential error messages: "`select distinct lookupCol, lookupCol from someTable…`" (this approach is not to be used with foreign key lookups, as the exact/correct id should always be found out. Instead try to find a way to make the lookup values reflect their uniqueness, e.g. by concatenating/joining further identifiers, as in "`select lookupCol+additionalLookup, lookupID…`" )

Even a lookup column without a lookup query is possible by just listing the possible values after the in the restriction separated by `||`, e.g.: `Yes||No||Maybe`:

![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/DBSheetsDefLookupList.PNG)

The DBSheet definition is created in four steps:

1.  First, all column definitions should have been entered and additional lookups either generated or entered/edited directly.
2.  Then the main query for retrieving the fields to be edited has to be generated. It can also be further customized, if needed.  
    However bear in mind that every change in the columns requires either an overwriting of the customizations and subsequent redoing them (cleaner) or constantly keeping the two synchronized. For customizing the data restriction part (Where Parameter Clause), a separate text input field can be used that allows the query to be regenerated without any intervening. Simply enter the restriction part (the Where clause without the "Where") and create the Query again.
3.  Then the DBSheet Definition needs to be stored, simply click "save DBsheet def", which allows you to choose a filename (if it hasn't been already saved. The file choice dialog can always be achieved by clicking "save DBSheet def As..."). Afterwards the information currently contained in the DBSheet columns, the DBsheet query and other information is stored in a DBSheet definition file (extension: xml)
4.  Finally, the DBSheet definition can be assigned to an Excel Worksheet with "assign DBsheet definition", creating a CUD enabled DBMapper with the selected DBSheet definition file to the currently opened Excel Worksheet (this only works if a workbook/sheet is active).

![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/DBSheetsDefinitionButton.PNG)

You can always test the main table query by clicking on "test DBSheet Query" next to the query definition. This opens an excel sheet with the results of the main table query being inserted. This Testsheet can be closed again either by simply closing it, or quicker by clicking on the same button again (that now changed its caption to "remove Testsheet").

#### Settings

Following Settings in DBAddin.xll.config or the referred DBAddinCentral.config affect behaviour of DBSheet definitions:
```xml
    <add key="ConfigName3" value="MSSQL"/>
    <add key="DBSheetConnString3" value="DRIVER=SQL SERVER;Server=Lenovo-PC;UID=sa;PWD=;Database=pubs;"/>
    <add key="DBidentifierCCS3" value="Database="/>
    <add key="DBSheetDefinitions3" value="C:\\dev\\DBAddin.NET\\definitions"/>
    <add key="dbGetAll3" value="sp_helpdb"/>
    <add key="dbGetAllFieldName3" value="name"/>
    <add key="ownerQualifier3" value=".dbo."/>
    <add key="dbPwdSpec3" value="PWD="/>
```


Explanation:
*   `ConfigName`**N**: freely definable name for your environment (e.g. Production or your database instance)
*   `DBSheetConnString`**N**: the connection string used to connect to the database for the DBSheet definitions
*   `DBidentifierCCS`**N**: used to identify the database within DBSheetConnString
*   `DBSheetDefinitions`**N**: path to the stored DBSheetdefinitions (default directory of assign DBsheet definitions and load/save DBSheet Defintions)
*   `dbGetAll`**N**: command for retrieving all databases/schemas from the database can be entered (for (MS) SQL server this is "`sp_helpdb`" for Oracle its "`select username from sys.all_users`".
*   `dbGetAllFieldName`**N**: If the result of above command has more than one column (like in sqlserver), you have to give the fieldname where the databases can be retrieved from.
*   `ownerQualifier`**N**: default owner qualifier for table when loading DBSheet definitions, if table name is not fully qualified (legacy DBSheet definitions)
*   `dbPwdSpec`**N**: Password entry specifier within DBSheetConnString

The entries DBisUserscheme and dbneedPwd are for Oracle databases where DBAddin has to switch to the scheme and therefore needs a password (Oracle has not been tested with the new DBAddin).
