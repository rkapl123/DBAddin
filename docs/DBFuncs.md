## DB Functions

There are three ways to query data with DBAddin:

1.  A list-oriented way using `DBListFetch`:  
    Here the values are entered into a rectangular list starting from the TargetRange cell (similar to MS-Query, actually the `QueryTables` Object is used to fill the data into the Worksheet).
2.  A record-oriented way using `DBRowFetch`:  
    Here the values are entered into several ranges given in the Parameter list `TargetArray`. Each of these ranges is filled in order of appearance with the results of the query.
3.  Setting the Query of a ListObject (new since Excel 2007) or a PivotTable to a defined query using `DBSetQuery`:  
    This requires an existing object (e.g. a ListObject created from a DB query/connection or a pivot table) and sets the target's queryobject to the desired one.

All these functions insert the queried data outside their calling cell context, which means that the target ranges can be put anywhere in the workbook (even outside of the workbook).

Additionally, some helper functions are available:

*   `chainCells`, which concatenates the values in the given range together by using "," as separator, thus making the creation of the select field clause easier.
*   `concatCells` simply concatenating cells (making the "&" operator obsolete)
*   `concatCellsText` above Function using the Text property of the cells, therefore getting the displayed values.
*   `concatCellsSep` concatenating cells with given separator.
*   `concatCellsSepText` above Function using the Text property of the cells, therefore getting the displayed values.
*   `DBString`, building a quoted string from an open ended parameter list given in the argument. This can also be used to easily build wildcards into the String.
*   `DBinClause`, building an SQL "in" clause from an open ended parameter list given in the argument.
*   `DBDate`, building a quoted Date string (standard format YYYYMMDD, but other formats can be chosen) from the date value given in the argument.

An additional cell context menu is available:

![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/ContextMenu.PNG)  

It provides:
*   A "jump" feature that allows to move the focus from the DB function's cell to the data area and from the data area back to the DB function's cell (useful in complex workbooks with lots of remote (not on same-sheet) target ranges)
*   Refreshing the currently selected DB function or it's data area. If no DB function or a corresponding data area is selected, then all DB Functions are refreshed.
*   [Creation of DB Functions](#create-db-functions)  

### Using the Functions

#### DBListFetch

<pre lang="vb">DBListFetch (Query, ConnectionString(optional), TargetRange,   
 FormulaRange(optional), ExtendDataArea(optional),   
 HeaderInfo(optional), AutoFit(optional),   
 AutoFormat(optional), ShowRowNum(optional))</pre>

The select statement for querying the values is given as a text string in parameter "Query". This text string can be a dynamic formula, i.e. parameters are easily given by concatenating the query together from other cells, e.g. `"select * from TestTable where TestID = "&A1`

The query parameter can also be a Range, which means that the Query itself is taken as a concatenation of all cells comprising that Range, separating the single cells with blanks. This is useful to avoid the problems associated with packing a large (parameterized) Query in one cell, leading to "Formula is too long" errors. German readers might be interested in [XLimits](http://www.xlam.ch/xlimits/xllimit4.htm), describing lots (if not all) the limits Excel faces.  

The connection string is either given in the formula, or can be left out. The connection string is then taken from the standard configuration settings from the key `ConstConnString`**N** of the set environment **N**. You can also fix this environment in the formula by passing **N** as the "connection string".

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

##### Connection String Special Settings

In case the "normal" connection string's driver (usually OLEDB) has problems in displaying data with DBListFetch and the problem is not existing in conventional MS-Query based query tables (using special ODBC connection strings that can't be used with DB functions), then following special connection string setting can be used:  

`StandardConnectionString;SpecialODBCConnectionString`

Example:  

`provider=SQLOLEDB;Server=LENOVO-PC;Trusted_Connection=Yes;Database=pubs;ODBC;DRIVER=SQL Server;SERVER=LENOVO-PC;DATABASE=pubs;Trusted_Connection=Yes`

This works around the issue with displaying GUID columns in SQL-Server.  

#### DBRowFetch

<pre lang="vb">DBRowFetch (Query, ConnectionString(optional),   
 headerInfo(optional/ contained in paramArray), TargetRange(paramArray))</pre>

For the query and the connection string the same applies as mentioned for "DBListFetch".  
The value targets are given in an open ended parameter array after the query, the connection string and an optional headerInfo parameter. These parameter arguments contain ranges (either single cells or larger ranges) that are filled sequentially in order of appearance with the result of the query.  
For example:  

<pre lang="vb">DBRowFetch("select job_desc, min_lvl, max_lvl, job_id from jobs " & "where job_id = 1",,A1,A8:A9,C8:D8)</pre>

would insert the first returned field (job_desc) of the given query into A1, then min_lvl, max_lvl into A8 and A9 and finally job_id into C8.  

The optional headerInfo parameter (after the query and the connection string) defines, whether field headers should be filled into the target areas before data is being filled.  
For example:  

<pre lang="vb">DBRowFetch("select job_desc, min_lvl, max_lvl, job_id from jobs","",TRUE,B8:E8, B9:E20)</pre>

would insert the the headers (`job_desc`, `min_lvl`, `max_lvl`, `job_id`) of the given query into B8:E8, then the data into B9:E20, row by row.  

The orientation of the filled rows is always determined by the first range within the `TargetRange` parameter array: if this range has more columns than rows, data is filled by rows, else data is filled by columns.  
For example:  

<pre>DBRowFetch("select job_desc, min_lvl, max_lvl, job_id from jobs","",TRUE,A5:A8,B5:I8)</pre>

would fill the same data as above (including a header), however column-wise. Typically this first range is used as a header range in conjunction with the headerInfo parameter.  

Beware that filling of data is much slower than with DBlistFetch, so use DBRowFetch only with smaller data-sets.  

#### DBSetQuery

<pre lang="vb">DBSetQuery (Query, ConnectionString(optional), TargetRange)</pre>

Stores a query into an Object defined in TargetRange (an embedded MS Query/Listobject, Pivot table, etc.)


### Additional Helper Functions

#### chainCells(Range)

<pre lang="vb">chainCells(ParameterList)</pre>

chainCells "chains" the values in the given range together by using "," as separator. It's use is mainly to facilitate the creation of the select field clause in the `Query` parameter, e.g.

<pre lang="vb">DBRowFetch("select " & chainCells(E1:E4) & " from jobs where job_id = 1","",A1,A8:A9,C8:D8)</pre>

Where cells E1:E4 contain job_desc, min_lvl, max_lvl, job_id respectively.

#### concatCells

<pre lang="vb">concatCells(ParameterList)</pre>

`concatCells` concatenates the values in the given range together. It's use is mainly to facilitate the building of very long and complex queries:

<pre lang="vb">DBRowFetch(concatCells(E1:E4),"",A1,A8:A9,C8:D8)</pre>

Where cells E1:E4 contain the constituents of the query respectively.  

#### concatCellsSep

<pre lang="vb">concatCellsSep(separator, ParameterList)</pre>

`concatCellsSep` does the same as concatCells, however inserting a separator between the concatenated values. It's use is the building of long and complex queries, too:

<pre lang="vbnet">DBRowFetch(concatCellsSep(E1:E4),"",A1,A8:A9,C8:D8)</pre>

Where cells E1:E4 contain the constituents of the query respectively.

All three concatenation functions (chainCells, concatCells and concatCellsSep) work with matrix conditionals, i.e. matrix functions of the form: `{=chainCells(IF(C2:C65535="Value";A2:A65535;""))}` that only chain/concat values from column A if the respective cell in column C contains "Value".

Both `concatCells` and `concatCellsSep` have a "Text" sibling that essentially does the same, except that it concats the displayed Values, not the true Values. So if you want to concatenate what you see, then `concatCellsText` and `concatCellsSepText` are the functions you need.

#### DBinClause

<pre lang="vbnet">DBinClause(ParameterList)</pre>
<pre lang="vbnet">DBinClauseStr(ParameterList)</pre>
<pre lang="vbnet">DBinClauseDate(ParameterList)</pre>

Creates an in clause from cell values, strings are created using `DBinClauseStr` with quotation marks, dates are created using `DBinClauseDate` using default date formatting (see DBDate for details).

<pre lang="vbnet">DBinClause("ABC", 1, DateRange)</pre>

Would return `”('ABC',1,'20070115')”`, if DateRange contained `15/01/2007` as a date value.  

#### DBString

<pre lang="vb">DBString(ParameterList)</pre>

This builds a Database compliant string (quoted) from the open ended parameter list given in the argument. This can also be used to easily build wildcards into the String, like

<pre lang="vb">DBString("_",E1,"%")</pre>

When E1 contains "test", this results in '\_test%', thus matching in a like clause the strings 'stestString', 'atestAnotherString', etc.

#### DBDate

<pre lang="vb">DBDate(DateValue, formatting (optional))</pre>

This builds from the date/datetime/time value given in the argument based on parameter `formatting` either

1.  (default formatting = DefaultDBDateFormatting Setting) A simple datestring (format `'YYYYMMDD``'`), datetime values are converted to `'YYYYMMDD` `HH:MM:SS'` and time values are converted to `'HH:MM:SS'`.
2.  (formatting = 1) An ANSI compliant Date string (format `date 'YYYY-MM-DD'`), datetime values are converted to `timestamp` `'YYYY-MM-DD` `HH:MM:SS'` and time values are converted to time `time 'HH:MM:SS'`.
3.  (formatting = 2) An ODBC compliant Date string (format `{d 'YYYY-MM-DD'}`), datetime values are converted to `{ts 'YYYY-MM-DD HH:MM:SS'}` and time values are converted to `{t 'HH:MM:SS'}`.
4.  (formatting = 3) An Access/JetDB compliant Date string (format `#YYYY-MM-DD#`), datetime values are converted to `#YYYY-MM-DD HH:MM:SS#` and time values are converted to `#HH:MM:SS#`.

An Example is give below:

<pre lang="vb">DBDate(E1)</pre>

*   When E1 contains the excel native date 18/04/2005, this results in : `'20050418'` (ANSI: `date '2005-04-18'`, ODBC: `{d '2005-04-18'}`).
*   When E1 contains the excel native date/time value 10/01/2004 08:05, this results in: `'20040110` `08:05:00``'` (ANSI: `timestamp '2004-01-10` `08:05:00``'`, ODBC: `{ts '2004-01-10` `08:05:00``'}`)
*   When E1 contains the excel native time value 08:05:05, this results in `'``08:05:05'` (ANSI: `time '``08:05:05'`, ODBC: `{t '``08:05:05'``}`)

Of course you can also change the default setting for formatting by changing the setting "`DefaultDBDateFormatting`" in the Addin settings

```xml
    <add key="DefaultDBDateFormatting" value="0"/>
```

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

Additionally the connection timeout (CnnTimeout, which can't be given in the functions) is also defined in the DBAddin settings.

### Cell Config Deployment

To easen the distribution of complex DB functions (especially queries), there is a config file mechanism in DBAddin: DB function (actually any Excel formula) configurations can be created in config files having extension XCL and are displayed with a tree-dropdown menu below "DB Configs" that displays the file hierarchy beneath ConfigStoreFolder for easy retrieval of the configs.  

The layout of these files is a pairwise, tab separated instruction where to fill (first element) Excel formulas (starting with "=" and being in R1C1 representation) or values (second element). Values are simple literal values to be inserted into Excel (numbers, strings, dates (should be interpretable by Excel !)), formulas are best taken from the return of ActiveCell.FormulaR1C1 !

#### Creating configs

There is a helping script ("createTableViewConfigs.vbs") to create a DBListFetch with a "select TOP 10000 * from ..." for all tables and views in a given database (In order for that script to work, the ADO driver has to support the "OpenSchema" method of the connection object). The working of that script is quite simple: It takes the name of the folder it is located in, derives from that the database name by excluding the first character and opens the schema information of that database to retrieve all view and table names from that. These names are used to build the config files (XCL).  

Other users can simply look up those config files with the hierarchical menu "DBConfigs", which is showing all config files under the ConfigStoreFolder (set in the global settings). Using folders, you can build categorizations of any depth here.  

DBAddin has a convenient feature to hierarchically order those config files further, if they are consistently named. For this to work, there either has to be a separation character between "grouping" prefixes (like "\_" in "Customer_Customers", "Customer_Addresses", "Customer_Pets", etc.) for grouping similar objects (tables, views) together or "CamelCase" Notation is used for that purpose (e.g. "CustomerCustomers", "CustomerAddresses", "CustomerPets").  

There is one setting key and three setting key groups to configure this further hierarchical ordering:  

```xml
    <add key="specialConfigStoreFolders" value="_pubs:_Northwind"/>
    <add key="_pubsMaxDepth" value="1"/>
    <add key="_pubsSeparator" value=""/>
    <add key="_NorthwindMaxDepth" value="1"/>
    <add key="_NorthwindSeparator" value="."/>
    <add key="_NorthwindFirstLetterLevel" value="True"/>
```

If you add the (sub) foldername to "specialConfigStoreFolders" (colon separated list) then this subfolder is regarded as needing special grouping of object names. The separator ("\_" or similar) can be given in  "(pathName)Separator", where (pathName) denotes the path name used above in "specialConfigStoreFolders". If this is not given then CamelCase is assumed to be the separating criterion.  

The maximum depth of the sub menus can be stated in "(pathName)MaxDepth", which denotes the depth of hierarchies below the uppermost in the (pathName) folder (default value is 10000, so practically infinite depth).  

You can add another hierarchy layer by setting "(pathName)FirstLetterLevel" to "True", which adds the first letter as the top level hierarchy.  

You can decide for each subfolder whether it's contents should be hierarchically organized by entering the relative path from ConfigStoreFolder for each subfolder in "specialConfigStoreFolders", or you can decide for all subfolders of that folder by just entering the topmost folder in "specialConfigStoreFolders".

#### Inserting configs

If the user retrieves the relevant configuration, a warning is shown and then the configured cells are entered into the active workbook as defined in the config, relative to the current selection.  

Cells in other worksheets are also filled, these are also taking the reference relative to the current selection. If the worksheet doesn't exist it is created.  

There are no checks (except for Excels sheet boundaries) especially whether any cells are overwritten !  

If the setting `ConfigSelect` (or any other ConfigSelect, see [Settings continued](https://rkapl123.github.io/DBAddin#settings-continued) is found in the settings, then the query template given there (e.g. `SELECT TOP 10 * FROM !Table!`) is used instead of the standard config (currently `SELECT TOP 10000 * FROM <Table>`) when inserting cell configurations. The respective Table is being replaced into `!Table!`.

#### Refreshing the config tree

To save time when starting up DBAddin/Excel, refreshing the config tree is only done when you open the Config Menu and click "refresh DB Config Tree".  
![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/DBConfigsRefresh.PNG)

### Create DB Functions

You can create the three DB Functions by using the cell context menu:  
![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/ContextMenu.PNG)  

The DB functions are created with an empty query string and full feature settings (e.g. Headers displayed, autosize and autoformat switched on) and target cells directly below the current active cell (except DBSetQuery for ListObjects, the ListObjects are placed to the right).

Below the results for a DB Function created in Cell A1:
*   DBListFetch:  `=DBListFetch("";"";A2;;;WAHR;WAHR;WAHR)`
*   DBRowFetch:   `=DBRowFetch("";"";WAHR;A2:K2)`
*   DBSetQueryPivot:   `=DBSetQuery("";"";A2)`
*   DBSetQueryListObject:   `=DBSetQuery("";"";B1)`

DBSetQuery also creates the target Object (a Pivot Table or a ListObject) below respectively to the right of the DB Function, so it is easier to start with. 
In case you want to insert DB Configurations (see [Cell Config Deployment](#Cell-Config-Deployment)), just place the selection on the inserted DB function cell and select your config, the stored query will replace the empty query in the created DB function.
For pivot tables the excel version of the created pivot table can be set with the user setting `ExcelVersionForPivot` (the numbers corresponding to the versions are: 0=2000, 1=2002, 2=2003, 3=2007, 4=2010, 5=2013, 6=2016, 7=2019=default if not set).
This is important to either provide backward compatibility with other users excels versions or to use the latest features.
#### Settings

Following Settings in DBAddin.xll.config or the referred DBAddinCentral.config affect the behaviour of DB functions:
```xml
    <add key="CnnTimeout" value="15" />
    <add key="DefaultEnvironment" value="3" />
    <add key="DontChangeEnvironment" value="False" />
    <add key="DebugAddin" value="False" />
```

Explanation:

*   `CnnTimeout`: the default timeout for connecting
*   `DefaultEnvironment`: default selected environment on startup
*   `DontChangeEnvironment`: prevent changing the environment selector (Non-Production environments might confuse some people)
*   `DebugAddin`: activate Info messages to debug addin

### Known Issues / Limitations

*  All DB getting functions (DBListfetch, DBRowFetch, etc....)
	*   A fundamental restriction for DB functions is that they should only exist alone in a cell (with no other DB functions). This is needed because linking the functions with their cell targets is done via a hidden name in the function cell (created on first invocation)  
	*   Query composition: Composing Queries (as these sometimes tend to be quite long) can become challenging, especially when handling parameters coming from cells. There is a simple way to avoid lots of trouble by placing the parts of a query in different lines/cells and putting all these cells together as a range in the DB functions first argument (query).
	*   When invoking an Excel Workbook from the commandline using CmdLogAddin (from a cmd script or the task scheduler), Excel may initialize the Add-in later than invoking the calculation of the DB function, which leads to an uninitialized host application object and therefore non-functional db functions (they all rely on the caller object of the Excel application to retrieve their calling cell's address). I'm still investigating into this.

* DBListFetch:
	*   formulaRange and extendArea = 1 or 2: Don't place content in cells directly below the formula Range as this will be deleted when doing recalculations. One cell below is OK.
	*   Worksheets with names like Cell references (Letter + number + blank + something else, eg. 'C701 Country') lead to a fundamental error with the names used for the data target. Avoid using those sheet names in conjunction with DBListFetch, i.e. do not use a blank between the 'cell reference' and the rest (eg. 'C701Country' instead of 'C701 Country').
	*   GUID Columns are not displayed when working with the standard data fetching method used by DBListFetch (using an opened recordset for adding a - temporary - querytable). A workaround has been built that circumvents this problem by adding the querytable the way that excel does (using the connection string and query directly when adding the querytable). This however implicitly opens another connection, so is more resource intensive. For details see [Connection String Special Settings](#Connection_String_Special_Settings)

* DBSetQuery
	* in DBSetQuery the underlying ListObject sometimes doesn't work with the SQLOLEDB provider, so there is a mechanism to change the provider part to something that works better. You can define a searched part of the connection string and its replacement in the settings of the environment (here environment 3):

```xml
    <add key="ConnStringSearch3" value="provider=SQLOLEDB"/>
    <add key="ConnStringReplace3" value="driver=SQL SERVER"/>
```
