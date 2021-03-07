## DBAddin

DBAddin is a ExcelDNA based Add-in for Database interoperability.

First, DBaddin provides DB Functions (see [DBFuncs Userdoc](DBFuncs.md)), which are an alternative to the Excel built-in MSQuery (integrated statically into worksheets having severe limitations in terms of querying and constructing parameterized queries (MS-Query allows parameterized queries only in simple queries that can be displayed graphically)).  

Next, methods for working with database data ([DBModifications](DBModif.md): DBMapper, DBAction and DBSequences) are included. This also includes a row entry oriented way to modify data in so called DBSheets (see [DBSheets](DBSheets.md)).

DBAddin.NET is the successor to the VB6 based Office Database Addin.

### Installation

* Dependencies
	* .NET 4.7 or higher (usually distributed with Windows)
	* Excel (minimum 2003)
	* ADO 2.5 or higher (usually distributed with Windows)

If any of these is missing, please install yourself before starting DBAddin.

Download the latest zip package in [https://github.com/rkapl123/DBAddin/tags](https://github.com/rkapl123/DBAddin/tags), unzip to any location and run deployAddin.cmd in the folder Distribution.
This copies DBAddin.xll, DBAddin.xll.config, DBaddinUser.config and DBAddinCentral.config to your %appdata%\Microsoft\AddIns folder and starts Excel for activating DBAddin (adding it to the registered Addins).

### Settings

Settings are done in three config files:

* DBAddin.xll.config (expected in the same %appdata%\Microsoft\AddIns folder as the DBAddin.xll) in section appSettings. DBAddin.xll.config has a reference
	* in the file attribute of the appSettings element to DBAddinCentral.config and
	* in the configSource attribute of the UserSettings element to DBaddinUser.config
* DBAddinCentral.config (this is a reference copy of the appSettings section, where the key/value pairs override the settings in DBAddin.xll.config)
* DBaddinUser.config (this is a reference copy of the UserSettings section, where the key/value pairs override both the settings in DBAddin.xll.config and DBAddinCentral.config)

In the DBAddin settings Group, there is a dropdown "settings", where you can modify these three settings inside Excel.

After installation you'd want to adapt the connection strings (ConstConnString**N**) that are globally applied if no function-specific connection string is given and environment **N** is selected.
This can be done by modifying DBAddin.xll.config or the referred DBaddinUser.config or DBAddinCentral.config (in this example the settings apply to environment 3):

```xml
<appSettings>
    <add key="ConfigName3" value="MSSQL"/>
    <add key="ConstConnString3" value="provider=SQLOLEDB;Server=Lenovo-PC;Trusted_Connection=Yes;Database=pubs;Packet Size=32767"/>
    <add key="ConfigStoreFolder3" value="C:\\dev\\DBAddin.NET\\source\\ConfigStore"/>
    <add key="DBidentifierCCS3" value="Database="/>
    <add key="DBSheetDefinitions3" value="C:\\dev\\DBAddin.NET\\definitions"/>
    <add key="dbGetAll3" value="sp_helpdb"/>
    <add key="dbGetAllFieldName3" value="name"/>
    <add key="ownerQualifier3" value=".dbo."/>
    <add key="DBSheetConnString3" value="DRIVER=SQL SERVER;Server=Lenovo-PC;UID=sa;PWD=;Database=pubs;"/>
    <add key="dbPwdSpec3" value="PWD="/>
    <add key="DBSetQueryListObjConnStringSearch3" value="provider=SQLOLEDB"/>
    <add key="DBSetQueryListObjConnStringReplace3" value="driver=SQL SERVER"/>
</appSettings>
```

Explanation:
*   `ConfigName`**N**: freely definable name for your environment (e.g. Production or your database instance)
*   `ConstConnString`**N**: the standard connection string used to connect to the database
*   `ConfigStoreFolder`**N**: all config files (*.xcl) under this folder are shown in a hierarchical manner in "load config"
*   `DBidentifierCCS`**N**: used to identify the database within the standard connection string or within DBSheetConnString
*   `DBSheetDefinitions`**N**: path to the stored DBSheetdefinitions (default directory of assign DBsheet definitions and load/save DBSheet Defintions)
*   `dbGetAll`**N**: command for retrieving all databases/schemas from the database can be entered (for (MS) SQL server this is "`sp_helpdb`" for Oracle its "`select username from sys.all_users`".
*   `dbGetAllFieldName`**N**: If the result of above command has more than one column (like in sqlserver), you have to give the fieldname where the databases can be retrieved from.
*   `ownerQualifier`**N**: default owner qualifier for table when loading DBSheet definitions, if table name is not fully qualified (legacy DBSheet definitions)
*   `DBSheetConnString`**N**: the connection string used to connect to the database for the DBSheet definitions
*   `dbPwdSpec`**N**: Password entry specifier within DBSheetConnString
*   `ConnStringSearch`**N**: part to be searched for replacement within the standard connection string in DBSetQuery
*   `ConnStringReplace`**N**: replacement for above

Other settings possible in DBAddin.xll.config (or DBAddinCentral.config):
```xml
    <add key="LocalHelp" value="C:\dev\DBAddin.NET\docs\doc\rkapl123.github.io\DBAddin\index.html"/>
    <add key="CmdTimeout" value="30" />
    <add key="CnnTimeout" value="15" />
    <add key="DefaultDBDateFormatting" value="0" />
    <add key="DefaultEnvironment" value="3" />
    <add key="DontChangeEnvironment" value="False" />
    <add key="maxCellCount" value="300000" />
    <add key="maxCellCountIgnore" value="False" />
    <add key="DebugAddin" value="False" />
    <add key="DBMapperCUDFlagStyle" value="TableStyleLight11" />
    <add key="DBMapperStandardStyle" value="TableStyleLight9" />
    <add key="maxNumberMassChange" value="10" />
    <add key="connIDPrefixDBtype" value="MSSQL" />
    <add key="DBSheetAutoname" value="True" />
    <add key="disableSettingsDisplay" value="addin"/>
		<add key="ConfigSelect" value="SELECT (SELECT Count(*) FROM !Table!) Anzahl, TOP 10 * FROM !Table!" />
```

Explanation:
*   `LocalHelp`: the path to local help files downloadable [here](doc.zip). To include it, extract the package into the respective folder and assign the file accordingly.
*   `CmdTimeout`: the default timeout for a command to execute
*   `CnnTimeout`: the default timeout for connecting
*   `DefaultDBDateFormatting`: default formatting choice for DBDate
*   `DefaultEnvironment`: default selected environment on startup
*   `DontChangeEnvironment`: prevent changing the environment selector (Non-Production environments might confuse some people)
*   `maxCellCount`: Cells being filled in Excel Workbook to issue a warning for refreshing DBFunctions (searching for them might take a long time...)
*   `maxCellCountIgnore`: Don't issue a warning, therefore ignore above setting.
*   `DebugAddin`: activate Info messages to debug addin
*   `DBMapperCUDFlagStyle`: Style for setting Excel data tables when having CUD Flags set on DBMappers
*   `DBMapperStandardStyle`: Style for setting Excel data tables when not having CUD Flags set on DBMappers
*   `maxNumberMassChange`: Threshold of Number of changes in CUDFlag DBMappers to issue a warning.
*   `connIDPrefixDBtype`: legacy DBSheet definitions have a Prefix in `connID` before the database that needs to be removed, this is the String to remove ...
*   `DBSheetAutoname`: When inserting DBSheet Definitions, automatically name Worksheet to the table name, if this is set
*   `disableSettingsDisplay`: put the settings that should not be available for viewing/editing to the user here (addin: DBAddin.xll.config, central: DBAddinCentral.config and user: DBaddinUser.config)
*   `ConfigSelect`: Use this template instead of standard config (currently `SELECT TOP 10000 * FROM <Table>`) when inserting cell configurations. The respective Table is being replaced into `!Table!`.

To change the settings, there is also a dropdown called "settings", where you can modify the DBAddin.xll.config and the referred DBAddinCentral.config including XML validation. You can have multiple same named entries (e.g. `ConfigSelect`) in your settings files, always the last one is taken as the active setting.

### Testing

Testing for MS SQL Server and other databases (MySQL, Oracle, PostgreSQL, DB2, Sybase and Access) can be done using the Testing Workbook "DBFuncsTest.xls".
To use that Testing Workbook you'll need the pubs database, where I have scripts available for Oracle, Sybase, DB2, PostgreSQL and MySql [here](PUBS_database_scripts.zip) (the MS-SQLserver version can be downloaded [here](https://www.microsoft.com/en-us/download/details.aspx?id=23654)). I've also added a pubs.mdb Access database in the test folder.

When starting the Testworkbook, after waiting for the – probable – connection error, you have to change the connection string(s) to suit your needs (see below for explanations).

![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/DBFunctionsTest.PNG)

Several connection strings for "DBFuncsTest.xls" are placed to the right of the black line, the actual connection is then selected by choosing the appropriate shortname (dropdown) in the yellow input field. After the connection has been changed don't forget to refresh the queries/DBforms by right clicking and selecting "refresh data".

### AboutBox, Logs, Purge hidden names and other tools

The Aboutbox can be reached by clicking the small dialogBox Launcher in the right bottom corner of the General Settings group of the DBAddin Ribbon.
![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/AboutBox.PNG)  

There is a possibility to set the future log events displayed (starting values are set in the config file). You can also fix legacy DBAddin functions, in case you decided to skip the possibility offered on opening a Workbook.

To see the Log, there is a separate Button in the log group of the DBAddin ribbon that also indicates the existence of warning log entries with a red exclamation mark.

![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/ConfigMenu.PNG)  

The DBListFetch's and DBRowFetch's target areas' extent is stored in hidden named ranges assigned both to the calling function cell (DBFsource(Key)) and the target (DBFtarget(Key)). These hidden names are used to keep track of the previous content to prevent overwriting, clearing old values, etc.
Sometimes during copying and pasting DB Functions, these names can get mixed up, leading to strange results or non-functioning of the "jump" function. In these cases, there is a tool in the DB Addin tools group, which may be used to "purge" these hidden named ranges in case of any strange behaviour due to multiple name assignments to the same area.  

The tool "Buttons" is used for switching designmode for DBModifier Buttons (identical to standard Excel button "Design Mode" in Ribbon "Developer tab", Group "Controls")

In the DBAddin settings Group, there is a dropdown "settings", where you can modify the DBAddin.xll.config and the referred DBAddinUser.config and DBAddinCentral.config including XML validation.

Right besides that dropdown, there is a shortcut to the Workbook properties (being the standard dialog Advanced Properties, accessible via File/Info) that allows you to change custom properties settings for DBAddin.
A green check shows that custom property DBFskip is not set to true for this workbook, therefore refreshing DB functions on opening the Workbook.

### Building

All packages necessary for building are contained, simply open DBaddin.sln and build the solution. The script deployForTest.cmd can be used to deploy the built xll and configs to %appdata%\Microsoft\AddIns

### docfx generated API documentation
[DBFuncs API documentation](api/index.html).
