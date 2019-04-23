## DB Functions

There are five ways to query data with DBAddin:

1.  A (fast) list-oriented way using `DBListFetch`.  
    Here the values are entered into a rectangular list starting from the TargetRange cell (similar to MS-Query, actually the `QueryTables` Object is used to fill the data into the Worksheet).
2.  A record-oriented way using `DBRowFetch`  
    Here the values are entered into several ranges given in the Parameter list `TargetArray`. Each of these ranges is filled in order of appearance with the results of the query.
3.  A "sentence" oriented way using `DBCellFetch`  
    Here the returned values are concatenated into one "sentence" by using specific column separators and row separators. Additionally "final" column separators and row separators can be defined to present data in a natural sentence way (e.g.: 10, 20, 30 and 25\. Here ", " is the column separator and " and " is the last row separator in a single column query.)
4.  Loading data into userforms (Listbox and Dropdown) using `DBMakeControl`  
    This is required for selecting values from a database. Max. 10 columns (incl. the invisible key column) can be defined in the associated query.
5.  Setting the Query of a ListObject (new since Excel 2007) or a PivotTable to a defined query using `DBSetQuery`  
    This requires an existing object (e.g. a DataList created from a DB query/connection or a pivot table) and sets the target's queryobject to the desired one.

Four of those user-defined functions (`DBListFetch`, `DBRowFetch`, `DBSetQuery` and `DBMakeControl`) insert the queried data outside their calling cell context, which means that the target ranges can be put anywhere in the workbook (even outside of the workbook).

Additionally, some helper functions are available:

*   `chainCells`, which concatenates the values in the given range together by using "," as separator, thus making the creation of the select field clause easier.
*   `concatCells` simply concatenating cells (making the "&" operator obsolete)
*   `concatCellsText` above Function using the Text property of the cells, therefore getting the displayed values.
*   `concatCellsSep` concatenating cells with given separator.
*   `concatCellsSepText` above Function using the Text property of the cells, therefore getting the displayed values.
*   `DBString`, building a quoted string from an open ended parameter list given in the argument. This can also be used to easily build wildcards into the String.
*   `DBinClause`, building an SQL in clause from an open ended parameter list given in the argument.
*   `DBDate`, building a quoted Date string (format YYYYMMDD) from the date value given in the argument.

There is also a "jump" feature that allows to move the focus from the DB function's cell to the data area and from the data area back to the DB function's cell (useful in complex workbooks with lots of remote (not on same-sheet) target ranges)

### Using the Functions

#### DBListFetch

<pre lang="vb.net">DBListFetch (Query, ConnectionString(optional), TargetRange,   
 FormulaRange(optional), ExtendDataArea(optional),   
 HeaderInfo(optional), AutoFit(optional),   
 AutoFormat(optional), ShowRowNum(optional))</pre>

The select statement for querying the values is given as a text string in parameter "Query". This text string can be a dynamic formula, i.e. parameters are easily given by concatenating the query together from other cells, e.g. `"select * from TestTable where TestID = "&A1`

The query parameter can also be a Range, which means that the Query itself is taken as a concatenation of all cells comprising that Range, separating the single cells with blanks. This is useful to avoid the problems associated with packing a large (parameterized) Query in one cell, leading to "Formula is too long" errors. German readers might be interested in [XLimits](http://www.xlam.ch/xlimits/xllimit4.htm), describing lots (if not all) the limits Excel faces.  

The connection string is either given in the formula, or for standard configuration can be left out and is then set globally in the registry key `[HKEY_CURRENT_USER\Software\VB and VBA Program Settings\DBAddin\Settings\ConstConnString]`

The returned list values are written into the Range denoted by "TargetRange". This can be  

*   just any range, resulting data being copied beginning with the left-uppermost cell
*   a self-defined named range (of any size) as TargetRange, which resizes the named range to the output size. This named range can be defined (and set as function parameter) either before or after results have been queried.

There is an additional `FormulaRange` that can be specified to fill “associated” formulas (can be put anywhere (even in other workbooks), though it only allowed outside of the data area). This `FormulaRange` can be

*   either a one dimensional row-like range or
*   a self-defined named range (of any size extent, columns have to include all calculated/filled down cells), which resizes the named range to the output size. This named range can be defined (and set as function parameter) either before or after results have been queried. Watch out when giving just one cell as the named range, this won't work as it's not possible in VBA to retrieve another assigned name of a cell and a hidden name us used to store the last extent of the formula range. The workaround is to assign at least two cells (columns or rows) to that name.

with formulas usually referring to cell-values fetched within the data area. All Formulas contained in this area are filled down to the bottom row of the `TargetRange`. In case the `FormulaRange` starts lower than the topmost row of `TargetRange`, then any formulas above are left untouched (e.g. enabling possibly different calculations from the rest of the data). If the `FormulaRange` starts above the `TargetRange`, then an error is given and no formulas are being refreshed down. If a `FormulaRange` is assigned within the data area, an error is given as well.

In case TargetRange is a named range and the FormulaRange is adjacent, the TargetRange is automatically extended to cover the FormulaRange as well. This is especially useful when using the compound TargetRange as a lookup reference (Vlookup).

The next parameter "`ExtendDataArea`" defines how DBListFetch should behave when the queried data extends or shortens:

*   0: DBListFetch just overwrites any existing data below the current `TargetRange`.
*   1: inserts cells of just the width of the `TargetRange` below the current TargetRange, thus preserving any existing data. However any data right to the target range is not shifted down along the inserted data. Beware in combination with a `FormulaRange` that the cells below the `FormulaRange` are not shifted along in the current version !!
*   2: inserts whole rows below the current `TargetRange`, thus preserving any existing data. Data right to the target range is now shifted down along the inserted data. This option is working safely for cells below the `FormulaRange`.

The parameter headerInfo defines whether Field Headers should be displayed (`TRUE`) in the returned list or not (`FALSE` = Default).

The parameter AutoFit defines whether Rows and Columns should be autofitted to the data content (`TRUE`) or not (`FALSE` = Default). There is an issue with multiple autofitted target ranges below each other, here the autofitting is not predictable (due to the unpredictable nature of the calculation order), resulting in not fitted columns sometimes.

The parameter `AutoFormat` defines whether the first data row's format information should be autofilled down to be reflected in all rows (`TRUE`) or not (`FALSE` = Default).

The parameter ShowRowNums defines whether Row numbers should be displayed in the first column (`TRUE`) or not (`FALSE` = Default).

##### <a name="Connection_String_Special_Settings"></a>Connection String Special Settings:

In case the "normal" connection string's driver (usually OLEDB) has problems in displaying data with DBListFetch and the problem is not existing in conventional MS-Query based query tables (using special ODBC connection strings that can't be used with DB functions), then following special connection string setting can be used:  

<pre lang="vb.net">StandardConnectionString;SpecialODBCConnectionString</pre>

Example:  

<pre lang="vb.net">provider=SQLOLEDB;Server=LENOVO-PC;Trusted_Connection=Yes;Database=pubs;ODBC;DRIVER=SQL Server;SERVER=LENOVO-PC;DATABASE=pubs;Trusted_Connection=Yes</pre>

This works around the issue with displaying GUID columns in SQL-Server.  

#### DBRowFetch

<pre lang="vb.net">DBRowFetch (Query, ConnectionString(optional),   
 headerInfo(optional/ contained in paramArray), TargetRange(paramArray))</pre>

For the query and the connection string the same applies as mentioned for "DBListFetch".  
The value targets are given in an open ended parameter array after the query, the connection string and an optional headerInfo parameter. These parameter arguments contain ranges (either single cells or larger ranges) that are filled sequentially in order of appearance with the result of the query.  
For example:  

<pre lang="vb.net">DBRowFetch("select job_desc, min_lvl, max_lvl, job_id from jobs " & "where job_id = 1",,A1,A8:A9,C8:D8)</pre>

would insert the first returned field (job_desc) of the given query into A1, then min_lvl, max_lvl into A8 and A9 and finally job_id into C8.  

The optional headerInfo parameter (after the query and the connection string) defines, whether field headers should be filled into the target areas before data is being filled.  
For example:  

<pre lang="vb.net">DBRowFetch("select job_desc, min_lvl, max_lvl, job_id from jobs","",TRUE,B8:E8, B9:E20)</pre>

would insert the the headers (`job_desc`, `min_lvl`, `max_lvl`, `job_id`) of the given query into B8:E8, then the data into B9:E20, row by row.  

The orientation of the filled rows is always determined by the first range within the `TargetRange` parameter array: if this range has more columns than rows, data is filled by rows, else data is filled by columns.  
For example:  

<pre>DBRowFetch("select job_desc, min_lvl, max_lvl, job_id from jobs","",TRUE,A5:A8,B5:I8)</pre>

would fill the same data as above (including a header), however column-wise. Typically this first range is used as a header range in conjunction with the headerInfo parameter.  

Beware that filling of data is much slower than with DBlistFetch, so use DBRowFetch only with smaller data-sets.  

#### DBCellFetch

<pre lang="vb.net">DBCellFetch (Query, ConnectionString (optional),   
 headerInfo(optional), colSep(optional), rowSep(optional), lastColSep(optional), lastRowSep(optional))</pre>

For the query and the connection string the same applies as mentioned for "DBListFetch".  
headerInfo specifies if additionally the headings of query should be filled (default = none). colSep is the usual column separator (default: ","); rowSep is the usual row separator (default CR); lastColSep (if given) specifies how the last column is differently separated from others; and finally lastRowSep how the last row is differently separated from others. lastColSep and lastRowSep are replaced by colSep and rowSep resp. if not given.  

Following Example should clarify everything (carriage returns are here just for clarity of display, actually this is rowSep, resp. lastRowSep):  

<pre lang="vb.net">header1 (colSep) header2 (colSep) header3 (colSep)... (colSep/lastColSep) headerN (rowSep)  
value11 (colSep) value12 (colSep) value13 (colSep)... (colSep/lastColSep) value1N (rowSep)  
...  
v(M-1)1 (colSep) v(M-1)2 (colSep) v(M-1)3 ........ (colSep/lastColSep) v(M-1)N (rowSep/lastRowSep)  
valueM1 (colSep) valueM2 (colSep) valueM3 (colSep)... (colSep/lastColSep) valueMN</pre>

#### DBMakeControl

<pre lang="vb.net">DBMakeControl (Query, ConnectionString(optional),   
 controlType(optional), headerInfo(optional), autoArrange(optional),  
 controlName(optional), dataTargetRange(optional), controlLocation(optional))</pre>

DBMakeControl creates DB bound Listboxes and Comboboxes (aka. Dropdowns). This is specified in the parameter controlType (0(default) is the Listbox and 1 the Combobox).

headerInfo (default = True) additionally creates a gray textbox above the Listbox/Combobox with header information. Removing headers is only possible by setting headerInfo to False, if they are simply deleted they will be recreated as long as headerInfo is True.

autoArrange (default = False) resizes the field widths and total width automatically each time a recalc is done (This requires the use of a special font and font size, so setting autoArrange to True will override most of your customizations of the control).

Setting the controlName enables you to give a special Name (which has to be unique in the workbook however) during creation of the control. This is useful, if you further want to use the control in VBA (event) procedures. This is only regarded when creating the control with the first call of the function, any subsequent changes will not yield anything. If you want to "rename" an existing control, just delete the existing control (incl. the header), the next function call will create a control having the given controlName.

For DB bound controls the dataTargetRange corresponds to the `LinkedCell` of the control, here the selection value of the control is placed. The dataTargetRange must be given as a String literal, as using a range would lead to a direct circular dependency (indirect circular dependencies are allowed, there are some quirks with Listboxes however - see Known Issues). The dataTargetRange can also specify a different sheet using "<Tablename>!<CellAddress>".

controlLocation finally specifies similar to above the placement of the control. This also has to be a String literal, additionally it has to be the same sheet as dataTargetRange (requirement of MS forms).

Following is an example of a dropdown and a listbox created with following functions respectively:  

<pre lang="vb.net">Listbox:  
=DBMakeControl("SELECT pub_id, pub_name, city, state, country FROM publishers";  
	activeConnString;;TRUE;TRUE)  

Combobox:  
=DBMakeControl("SELECT emp_id, fname, minit, lname, job_id,  
	convert(varchar(10),job_lvl) job_lvl, pub_id, hire_date FROM  
	employee";activeConnString;1;TRUE)  
</pre>

![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/DBAddin-Dateien/clip_image004.jpg)  

The DB bound control retain the selection during a refresh (e.g. saving/reopening the workbook).  

#### DBSetQuery

<pre lang="vb.net">DBSetQuery (Query, ConnectionString(optional), TargetRange)</pre>

Stores a query into an Object defined in TargetRange (an embedded MS Query/Listobject, Pivot table, etc.)


### Additional Helper Functions

#### chainCells(Range)

<pre lang="vb.net">chainCells(ParameterList)</pre>

chainCells "chains" the values in the given range together by using "," as separator. It's use is mainly to facilitate the creation of the select field clause in the `Query` parameter, e.g.

<pre lang="vb.net">DBRowFetch("select " & chainCells(E1:E4) & " from jobs where job_id = 1","",A1,A8:A9,C8:D8)</pre>

Where cells E1:E4 contain job_desc, min_lvl, max_lvl, job_id respectively.

#### concatCells

<pre lang="vb.net">concatCells(ParameterList)</pre>

`concatCells` concatenates the values in the given range together. It's use is mainly to facilitate the building of very long and complex queries:

<pre lang="vb.net">DBRowFetch(concatCells(E1:E4),"",A1,A8:A9,C8:D8)</pre>

Where cells E1:E4 contain the constituents of the query respectively.  

#### concatCellsSep

<pre lang="vb.net">concatCellsSep(separator, ParameterList)</pre>

`concatCellsSep` does the same as concatCells, however inserting a separator between the concatenated values. It's use is the building of long and complex queries, too:

<pre lang="vb.net">DBRowFetch(concatCellsSep(E1:E4),"",A1,A8:A9,C8:D8)</pre>

Where cells E1:E4 contain the constituents of the query respectively.

All three concatenation functions (chainCells, concatCells and concatCellsSep) work with matrix conditionals, i.e. matrix functions of the form: `{=chainCells(IF(C2:C65535="Value";A2:A65535;""))}` that only chain/concat values from column A if the respective cell in column C contains "Value".

Both `concatCells` and `concatCellsSep` have a "Text" sibling that essentially does the same, except that it concats the displayed Values, not the true Values. So if you want to concatenate what you see, then `concatCellsText` and `concatCellsSepText` are the functions you need.

#### DBinClause

<pre lang="vb.net">DBinClause(ParameterList)</pre>

Creates an in clause from cell values, strings are created with quotation marks, dates are created with DBDate (see there for details, formatting is 0).

<pre lang="vb.net">DBinClause("ABC", 1, DateRange)</pre>

Would return `”('ABC',1,'20070115')”`, if DateRange contained `15/01/2007` as a date value.  

#### DBString

<pre lang="vb.net">DBString(ParameterList)</pre>

This builds a Database compliant string (quoted) from the open ended parameter list given in the argument. This can also be used to easily build wildcards into the String, like

<pre lang="vb.net">DBString("\_",E1,"%")</pre>

When E1 contains "test", this results in '\_test%', thus matching in a like clause the strings 'stestString', 'atestAnotherString', etc.

#### DBDate

<pre lang="vb.net">DBDate(DateValue, formatting (optional))</pre>

This builds from the date/datetime/time value given in the argument based on parameter `formatting` either

1.  (default formatting = DefaultDBDateFormatting Setting) A simple datestring (format `'YYYYMMDD``'`), datetime values are converted to `'YYYYMMDD` `HH:MM:SS'` and time values are converted to `'HH:MM:SS'`.
2.  (formatting = 1) An ANSI compliant Date string (format `date 'YYYY-MM-DD'`), datetime values are converted to `timestamp` `'YYYY-MM-DD` `HH:MM:SS'` and time values are converted to time `time 'HH:MM:SS'`.
3.  (formatting = 2) An ODBC compliant Date string (format `{d 'YYYY-MM-DD'}`), datetime values are converted to `{ts 'YYYY-MM-DD HH:MM:SS'}` and time values are converted to `{t 'HH:MM:SS'}`.

An Example is give below:

<pre lang="vb.net">DBDate(E1)</pre>

*   When E1 contains the excel native date 18/04/2005, this results in : `'20050418'` (ANSI: `date '2005-04-18'`, ODBC: `{d '2005-04-18'}`).
*   When E1 contains the excel native date/time value 10/01/2004 08:05, this results in: `'20040110` `08:05:00``'` (ANSI: `timestamp '2004-01-10` `08:05:00``'`, ODBC: `{ts '2004-01-10` `08:05:00``'}`)
*   When E1 contains the excel native time value 08:05:05, this results in `'``08:05:05'` (ANSI: `time '``08:05:05'`, ODBC: `{t '``08:05:05'``}`)

Of course you can also change the default setting for formatting by changing the setting "`DefaultDBDateFormatting`" in the global settings

<pre lang="vb.net">"DefaultDBDateFormatting"="0"</pre>

### Modifications of DBFunc Behaviour

There are some options to modify  

*   the refreshing and
*   the storing

of DB functions data area within a Workbook.  

You can set these options in Excel's Custom Properties (Menu File/Properties, Tab "Customize"):  

#### Skipping Data Refresh when opening Workbook

To disable refreshing of DBFunctions when opening the workbook create a boolean custom property "DBFSkip" set to "Yes" (set to "No" to disable skipping).  

#### Prevent Storing of retrieved Data in the Workbook

To prevent storing of the contents of a DBListFetch or DBRowFetch when saving the workbook create a boolean custom property "DBFCC(DBFunctionSheet!DBFunctionAddress)" set to "Yes" (set to "No" to reenable storing). This clears the data area of the respective DB function before storing and refreshes it afterwards (Note: If the custom property "DBFSkip" is set to "Yes", then this refreshing is skipped like when opening the Workbook)  

Example: The boolean custom property "DBFCCTable1!A1" would clear the contents of the data area for the DBFunction entered in Table1, cell "A1".  

To prevent storing of the contents for all DBFunctions create a boolean Custom Property "DBFCC*" set to "Yes".  

Excel however doesn't fill only contents when filling the data area, there's also formatting being filled along, which takes notable amounts of space (and saving time) in a workbook. So to really have a small/quick saving workbook, create a boolean custom property "DBFCA(DBFunctionSheet!DBFunctionAddress)" set to "Yes" (set to "No" to reenable storing). This clears everything in the the data area of the res  

Example: The boolean custom property "DBFCATable1!A1" would clear everything from the data area for the DBFunction entered in Table1, cell "A1".  

To prevent storing of everything (incl. formats) for all DBFunctions create a boolean Custom Property "DBFCA*" set to "Yes".  


#### Global Connection Definition and Query Builder with MSQuery

There are two possibilities of connection strings: ODBC or OLEDB. ODBC hast the advantage to seamlessly work with MS-Query, native OLEDB is said to be faster and more reliable (there is also a generic OLEDB over ODBC by Microsoft, which emulates OLEDB if you have just a native ODBC driver available).

Now, if using **ODBC** connection strings (those containing "Driver="), there is a straightforward way to redefine queries directly from the cell containing the DB function: just right click on the function cell and select "build DBfunc query". Then MS-query will allow you to redefine the query which you can use to overwrite the function's query.

If using **OLEDB** connection strings, MS-query will try to connect using a system DSN named like the database as identified after the DBidentifierCCS given in the standard settings (see Installation section).

The DBidentifierCCS is used to identify the database within the standard default connection string, The DBidentifierODBC is used to identify the database within the connection definition returned by MS-Query (to compare and possibly allow to insert a custom connection definition within the DB function/control). Usually these identifiers are called "Database=" (all SQLservers, MySQL), "location=" (PostgreSQL), "User ID/UID" (oracle), "Data source=" (DB2)

Additionally the timeout (CnnTimeout, which can't be given in the functions) is also defined in the standard settings.

### Supporting Tool Cell Config Deployment

To easen the distribution of complex DB functions (resp. queries), there is a special deployment mechanism in the DBAddin Commandbar: Saving of DB function configurations can be done with the button "Save Config", whereas for loading there are two possibilities: The button "Load Config" (displaying a simple load file dialog) and a tree-dropdown menu below "DBConfigs" that displays the file hierarchy beneath ConfigStoreFolder for easy retrieval of the configs.  

#### Creating configs

"Save Config" asks you to select cells you want to store for others to import into their sheet. This is done by either selecting one contiguous area or by Ctrl+clicking the single cells you want to add to the distribution. Finally a Save Dialog asks you for the filename where these cell contents/formulas should be stored. If you choose an existing file, you're asked whether the config should be appended to that file.  

Other users can simply look up those config files either with "Load Config" or the hierarchical menu "DBConfigs", which is hierarchically showing all config files under the ConfigStoreFolder (set in the global settings). Using folders, you can build categorizations of any depth here.  

There is a helping script ("createTableViewConfigs.vbs") to create a DBListFetch with a "select TOP 10000 * from ..." for all tables and views in a given database (In order for that script to work, the ADO driver has to support the "OpenSchema" method of the connection object). The working of that script is quite simple: It takes the name of the folder it is located in, derives from that the database name by excluding the first character and opens the schema information of that database to retrieve all view and table names from that. These names are used to build the Excel and Word config files (XCL/WRD).  

DBAddin has a convenient feature to hierarchically order those config files furthermore, if they are consistently named. For this to work, there either has to be a separation character between "grouping" prefixes (like "\_" in "Customer_Customers", "Customer_Addresses", "Customer_Pets", etc.) for grouping similar objects (tables, views) together or "CamelCase" Notation is used for that purpose (e.g. "CustomerCustomers", "CustomerAddresses", "CustomerPets").  

There is one registry setting and two registry setting groups to configure this further hierarchical ordering:  

<pre lang="vb.net">Windows Registry Editor Version 5.00  

[HKEY_CURRENT_USER\\Software\\VB and VBA Program Settings\\DBAddin\\Settings]  
"specialConfigStoreFolders"="(pathName):.pubs:.Northwind"  
"(pathName)MaxDepth"="1"  
"(pathName)Separator"=""  
"(pathName)FirstLetterLevel"="True"</pre>

If you add the (sub) foldername to "specialConfigStoreFolders" (colon separated list) then this subfolder is regarded as needing special grouping of object names. The separator ("\_" or similar) can be given in  "(pathName)Separator", where (pathName) denotes the path name used above in "specialConfigStoreFolders". If this is not given then CamelCase is assumed to be the separating criterion.  

The maximum depth of the sub menus can be stated in "(pathName)MaxDepth", which denotes the depth of hierarchies below the uppermost in the (pathName) folder (default value is 10000, so practically infinite depth).  

You can add another hierarchy layer by setting "(pathName)FirstLetterLevel" to "True", which adds the first letter as the top level hierarchy.  

You can decide for each subfolder whether it's contents should be hierarchically organized by entering the relative path from ConfigStoreFolder for each subfolder in "specialConfigStoreFolders", or you can decide for all subfolders of that folder by just entering the topmost folder in "specialConfigStoreFolders". Beware that the backslash (path separator) in (pathName) needs to be entered quoted (two "\" !) to be recognized when importing the registry key files!  

In some instances the Windows API doesn't return the directories/filenames in alphabetical order (e.g. in some SAMBA networks), so you'd need to set the following registry setting to "True", which does a sort on the config folder's content before putting it into the tree dropwdown:  

<pre lang="vb.net">"sortConfigStoreFolders"="False"</pre>

#### Inserting configs

If the user finds/loads the relevant configuration, a warning is shown and then the configured cells are entered into the active workbook as defined in the config, relative to the current selection. The reference cell during saving is always the left/uppermost cell (A1), so anything chosen in other cells will be placed relatively right/downward.  

Cells in other worksheets are also filled, these are also taking the reference relative to the current selection (when loading) or cell A1 (when saving). If the worksheet doesn't exist it is created.  

Currently there are no checks (except for Excels sheet boundaries) as whether any cells are overwritten !  

#### Refreshing the config tree

To save time when starting up DBAddin/Excel, the previously used way of refreshing the config tree (when starting up Excel/DBAddin) was abandoned in favour of a better one: now each user can decide when to refresh the config tree for himself with the button "refresh Configs" right below the menu button "DBConfigs".  

This refreshing also restores "theDBSheetAppHandler" object as a side effect, which re-enables DBSheet treatment after sporadic invalidation of this object (needed until the reason for this has been found out).  

### Installation

#### Dependencies

*   Office and Excel Object Libraries (well, as it is an Excel Addin, this should be expected on the target system)
*   ADO 2.5 or higher (usually distributed with Windows)

If any of these is missing, please install yourself before starting DBAddin.

After installation you'd want to adapt the standard default connection string (ConstConnString) that is globally applied if no function-specific connection string is given. This can be done by modifying and importing DBAddinSettings.reg into your registry.

<pre lang="vb.net">Windows Registry Editor Version 5.00  

[HKEY_CURRENT_USER\\Software\\VB and VBA Program Settings\\DBAddin\\Settings]  
"ConstConnString"="provider=SQLOLEDB;Server=LENOVO_PC;Trusted_Connection=Yes;Database=InfoDB;Packet Size=32767"  
"DBidentifierCCS"="Database="  
"DBidentifierODBC"="Database="  
"CnnTimeout"="15"  
"DefaultDBDateFormatting"="0"  
"ConfigStoreFolder"="(YourPathToTheConfigStore)\\ConfigStore"  
"LocalHelp"="(YourPathToTheDocumentation)\\HelpFrameset.htm"  
</pre>

The other settings:

*   `DBidentifierCCS`: used to identify the database within the standard default connection string
*   `DBidentifierODBC`: used to identify the database within the connection definition returned by MS-Query
*   `CnnTimeout:` the default timeout for connecting
*   `DefaultDBDateFormatting: `default formatting choice for DBDate
*   `ConfigStoreFolder: `all config files (xcl/wrd) under this folder are shown in a hierarchical manner in "load config"
*   `LocalHelp: `the path to the local help files downloadable [here](doc.zip). To include it into the standard installation, put the contained documentation folder into the DBAddin Application folder (e.g. C:\Program Files\RK\DBAddin)

When starting the Testworkbook, after waiting for the – probable – connection error, you have to change the connection string(s) to suit your needs (see below for explanations).

![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/DBAddin-Dateien/clip_image005.jpg)

Several connection strings for "DBFuncsTest.xls" are placed to the right of the black line, the actual connection is then selected by choosing the appropriate shortname (dropdown) in the yellow input field. After the connection has been changed don't forget to refresh the queries/DBforms by right clicking and selecting "refresh data".

### Points of Interest

The basic principle behind returning results into an area external to the Database query functions, is the utilisation of the calculation event (as mentioned in and inspired by the excelmvf project, see [http://www.codeproject.com/macro/excelmvf.asp](http://www.codeproject.com/macro/excelmvf.asp) for further details), as Excel won't allow ANY side-effects inside a UDF.

There is lots of information to be carried between the function call and the event (and back for status information). This is achieved by utilising a so-called "`calcContainer`" and a "`statusMsgContainer`", basically being VBA classes abused as a simple structure that are stored into global collections called "`allCalcContainers`" and "`allStatusContainers`". The references of the correct calcContainers and statusMsgContainers are the Workbook-name, the table name and the cell address of the calling functions which is quite a unique description of a function call (this description is called the "`callID`" in the code).

Below diagram should clarify the process:

![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/DBAddin-Dateien/clip_image001.gif)  

The real trick is to find out when resp. where to get rid of the calc containers, considering Excel's intricate way of invoking functions and the calc event handler (the above diagram is simplifying matters a bit as the chain of invocation is by no way linear in the calculations in the dependency tree).

Excel sometimes does additional calculations to take shortcuts and this makes the order of invocation basically unpredictable, so you have to take great care to just work on every function once and then remove the `calcContainer.`

After every calculation event the working `calcContainers` are removed, if there are no more `calcContainers` left, then `allCalcContainers` is reset to "Nothing", being ready for changes in input data or function layout. Good resources for more details on the calculation order/backgrounds is Decision Model's Excel Pages, especially Calculation Secrets ([http://www.decisionmodels.com/calcsecretsc.htm](http://www.decisionmodels.com/calcsecretsc.htm)).

The DBListFetch's target areas' extent is stored in hidden named ranges assigned both to the calling function cell (DBFsource(Key)) and the target.

There is a procedure in the Functions module, which may be used to "purge" these hidden named ranges in case of any strange behaviour due to multiple name assignments to the same area. This behaviour usually arises, when redefining/adding dbfunctions that refer to the same target area as the original/other dbfunction. The procedure is called "purge" and may be invoked from the VBA IDE as follows:

<pre lang="vb.net">Sub purge()  
  Set dbfuncs = CreateObject("DBAddin.Functions")  
  dbfuncs.purge  
End Sub  
</pre>

### Known Issues / Limitations

*   DBMakeControl

*   Listbox control selection shifts uncontrolled upwards when using mouse. This only occurs when the DBMakeControl is either dependent on calculation results that are themselves dependent on the control's target cell (kind of "circular dependency") or if there is a dependent DBListfetch in the same sheet with extendDataArea != 0\. In this case Excel somehow always sets other cells to "dirty" status, thus enforcing a recalculation of the DBMakeControl. The workaround is to either use up/down keys after the initial click for precise selecting or place the DBMakeControl function in a different sheet than the DBListFetch's target area.  

*   Listbox control selection is not visibly retained during refresh (context menu resfresh data or full calc (Ctrl-Alt-F9)), i.e. the data target is kept, but the entry's highlighting is gone. This is the case when the control is placed on the same sheet as the DBMakeControl function. Workaround: put DBMakeControl in a different sheet than the control itself.  

*   All DB getting functions (DBListfetch, DBRowFetch, etc....)

*   A fundamental restriction for these function is that they should only exist alone in a cell with no other DB getters. This is needed because linking the functions with their cell targets is done via a hidden name in the function cell (created on first invocation)  

*   DBListFetch:

*   formulaRange and extendArea = 1 or 2: Don't place content in cells directly below the formula Range as this will be deleted when doing recalculations. One cell below is OK.
*   In Worksheets with names like Cell references (Letter + number + blank + something else, eg. 'C701 Country') this leads to a fundamental error with the names used for the data target. Avoid using those sheet names in conjunction with DBListFetch, i.e. do not use a blank between the 'cell reference' and the rest (eg. 'C701Country' instead of 'C701 Country').
*   GUID Columns are not displayed when working with the standard data fetching method used by DBListFetch (using an opened recordset for adding a - temporary - querytable). A workaround has been built that circumvents this problem by adding the querytable the way that excel does (using the connection string and query directly when adding the querytable). This however implicitly opens another connection, so is more resource intensive. For details see [Connection String Special Settings](#Connection_String_Special_Settings)

*   Query composition: Composing Queries (as these sometimes tends to be quite long) can become challenging, especially when handling parameters coming from cells. There is a simple way to avoid lots of trouble by placing
*   When invoking an Excel Workbook from the commandline (from a cmd script or the task scheduler) Excel may register (call the connect method of the Add-in) the Add-in later than invoking the calculation which leads to an uninitialized host application object and therefore a non-functional dbfunctions (they all rely on the caller object of the Excel application to retrieve their calling cell's address). I'm still investigating into this.
*   The Workbook containing the DB functions may not have any VBA Code with Compile Errors, or it will return an "Application defined Error". This relates to Excel not passing the Application.Caller object correctly to UDFs when having compile errors in VBA-Code.


## Mapper

Mapper is an object that you can use to save Excel Range data to database table(s). The function that is used to do that is

#### saveRangeToDB

<pre lang="vb.net">saveRangeToDB(DataRange As Excel.Range, tableNamesStr As String, _  
 primKeysStr As String, primKeyColumnsStr As String, startDataColumn As Integer, _  
 connID As String, ParamArray optionalArray() As Variant) As Boolean</pre>

Dumps table given in DataRange to database table(s) in tableNamesStr in a database specified by connID.  
Parameters:  

*   DataRange As Excel.Range = The Excel Range to be saved to the database
*   tableNamesStr As String = the table(s) where the range is to be saved to. If more than one table is given then use comma as a separator.
*   primKeysStr As String = The primary key(s) used for identifying the rows in above table(s). If more than one primary key is given then use comma as a separator. The count of primary keys has to be the same as the tables.
*   primKeyColumnsStr As String = The columns where the primary key value(s) are found in the Excel Range. If more than one primary key is given then use comma as a separator. The count of primary key columns has to be the same as the tables/primary keys.
*   startDataColumn As Integer = The starting column for the actual data (excluding the primary column(s)).
*   connID As String = The connection ID specified in DBConns.xml (see `DBConnFileName` in [Installation)](#Installation)
*   if given in the optionalArray:
*   [insertIfMissing] As Boolean = if True then insert row into table if primary key is missing there. Default = False (only update)
*   [excludeColumns] As String = The column numbers of those columns that should be excluded from the storing (helper columns,etc.). If more than one column is to be excluded then use comma as a separator.
*   [specialExceptionTable] As String = The table in whose specialExceptionField the update values should be checked whether they are of specialExceptionValue or result in an error. If that is the case then the rest of the row is not being stored, and the values given in specialExceptionDefaults are stored into the fields and tables given in specialExceptionDefaults.
*   [specialExceptionField] As String = The field that is checked whether it's value is a specialExceptionValue (typically an undesired value, like empty or zero) or whether it is an error.
*   [specialExceptionValue] As Variant = The special value that is checked (typically an undesired value, like empty or zero)
*   [specialExceptionDefaults] As Variant() = Array of Array(""Tablename"",""Fieldname"",""COL/VAL"",value): The table/field/Column/Value combinations that are used for determining the special default values to be stored in the special fields of special tables in case of an above found exception. In case the third part of a specialExceptionDefaultArray is "COL" then the default values are taken from the column given in the fourth part (value). In case the third part of a specialExceptionDefaultArray is "VAL" then the default values are taken directly from the fourth part.

The function returns True if successful, false in case of errors.  

Example:  

<pre lang="vb.net">Sub testdumpRangeToDB()  
  Set mapper = CreateObject("DBAddin.Mapper")  
  mapper.saveRangeToDB Range("employee"), "employee", "emp_id", 1, 2, "MSSQLPUBSTest", True  
End Sub</pre>

Here the data in range `employee` are stored to the table `employee` using primary key `emp_id`, which is located in column 1. Data is starting in column 2, the connection information is stored in connection id "MSSQLPUBSTest" and any missing employees are inserted with a new emp_id key if it is missing.  

An Example for the Mapper object/saveRangeToDB usage can also be found in the DBFuncsTest.xls Workbook, Sheet "MapperExamples".
