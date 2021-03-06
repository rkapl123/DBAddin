## DBModifications

DBModifications can be used to
* save Excel Range data to database table(s): DBMapper
* modify DB Data using Data Manipulation SQL (update, delete, etc.): DBAction
* and creating sequences of these activites: DBSequence.

The target data referred to by DBMapper and DBAction (data is the DML SQL statement(s)) is specified by special Range names, any other definitions (environment, target database, etc.) is stored in a custom property of the workbook having the same name as the target range.

Examples for the usage of DBMapper can be found in the DBMapperTests.xlsx Workbook.

## Create DBModifiers

You can create the three DB Modifiers by using the cell context menu:  
![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/ContextMenu.PNG)  

The DBModifier creation/editing is shown below (examples already filled, when activated on a blank cell all entries are empty):  
(one feature that can not be set in the dialogs is a customized confirmation text for the "Ask for execution" dialog, this is done with Edit DBModifier Definitions, see below)

### DB Mappers are created/edited with the following dialog:  
![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/DbMapperCreate.PNG)  

*   DBMapper Name: Enter the name for the selected Range (containing the Data including header fields) that will be used to identify the DBMap in the "Execute DBModifier" Group dropdowns. If no name is given here, then UnnamedDBMapper will be used to identify it.
*   Tablename: Database Table, where Data is to be stored.
*   Primary Keys: String containing primary Key names for updating table data, comma separated.
*   Database: Database to store DBMaps Data into.
*   Ignore Columns: columns to be ignored (e.g. helper columns), comma separated.
*   Additional Stored Procedure: additional stored procedure to be executed after saving.
*   Insert If Missing: if set, then insert row into table if primary key is missing there. Default = False (only update).
*   Store DBMap on Save: should DBMap also be saved on Excel Workbook Saving? (default no). If multiple DBModifiers (DB Mappers and DB Actions, DB Sequences are not bound to a worksheet) are defined to be stored/executed on save, then the DBModifiers being on the active worksheet are done first (without any specific order), then those on the other worksheets are done (also without any specific order).
*   Environment: The Environment, where connection id should be taken from (if not existing, take from selected Environment in DB Addin General Settings Group).
*   Exec on Save: Should the DBMap be executed when the workbook is being saved?
*   Ask for execution: Before execution of the DBMap, ask for confirmation. A custom text can be given in the CustomXML definition element confirmText (see below).
*   C/U/D Flags: special mode used for row-by-row editing (inserting, updating and deleting rows). Only edited rows will be done when executing. Deleting rows is node with the special context menu item "delete Row" (or pressing Ctrl-Shift-D).
*   Ignore data errors: replace excel errors like #VALUE! with null on updating/inserting. Otherwise an error message is passed and execution is skipped for that row.
*   Auto Increment: Allow empty primary column values (only for a single primary key!) in use with tables that have the IsAutoIncrement property set for this primary column (typically because of an identity specification for that column in the Database).
*   Create CB: create a commandbutton for the DB Sequence in the current Worksheet.
*   Hyperlink: click on it to highlight/select the DB Mapper area.

You can always edit these parameters by selecting a cell in the DB Mapper area and invoking the context menu again.

The range that is used for holding the data to be stored can be identified in three different ways:
*   A plain address: Here the range is automatically extended by using the first column and the first row (header row), if no Auto Increment or C/U/D Flags are set. In this case empty primary column(s) are allowed so automatic extension won't work.
*   A named offset formula (that is used to dynamically assign the data range).
*   A data list object (especially in use for C/U/D Flags).

The clickable Hyperlink shows the range address of the data range, a named offset formula is displayed after the address:  
![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/DbMapperCreateOffsetFormula.PNG)  

### DB Actions are created/edited with following dialog:  
![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/DbActionCreate.PNG)  

*   DBAction Name: Enter the name for the selected cell or range that will be used to identify the DBAction in the "Execute DBModifier" Group dropdowns. If no name is given here, then UnnamedDBAction will be used to identify it.
*   Database: Database to do the DBAction in.
*   Environment: The Environment, where connection id should be taken from (if not existing, take from selected Environment in DB Addin General Settings Group).
*   Exec on Save: Should the DBAction be executed when the workbook is being saved?
*   Ask for execution: Before execution of the DBAction, ask for confirmation. A custom text can be given in the CustomXML definition element confirmText (see below).
*   The actual DBAction to be done is defined in a range that is named like the DBAction definition (the hyperlink takes you there). This range can be dynamically computed as all ranges in excel.
*   Create CB: create a commandbutton for the DB Sequence in the current Worksheet.
*   Hyperlink: click on it to highlight/select the DB Action range.

You can always edit these parameters by selecting a cell in the range of the DB Action area and invoking the context menu again.

### DB Sequences are created/edited with following dialog:  
![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/DbSequenceCreate.PNG)  

*   DBSequence Name: Enter the name for the selected Range that will be used to identify the DBSequence in the "Execute DBModifier" Group dropdowns. If no name is given here, then UnnamedDBAction will be used to identify it.
*   Exec on Save: Should the DBSequence be executed when the workbook is being saved?
*   Ask for execution: Before execution of the DBSequence, ask for confirmation. A custom text can be given in the CustomXML definition element confirmText (see below).
*   Sequence Step Datagrid: here the available DBMappers, DBActions and DBFunctions (DBListfetch/DBRowFetch/DBSetQuery) can be added that are then executed in Sequence. If you are executing all sequence steps in the same environment, its possible to run the sequence in a transaction context by placing DBBegin at the top and DBCommitRollback at the bottom of the sequence.
*   move the sequence steps up and down by selecting a row and using the context menu (right mouse button).
*   Create CB: create a commandbutton for the DB Sequence in the current Worksheet.

As DB Sequences have no Range with data/definitions, invoking the context menu always creates new DB Sequences. You can edit existing DB Sequences
by Ctrl-Shift clicking the Execute DBModifier Groups dropdown menus or by Ctrl-Shift clicking the created commandbuttons.

## Edit DBModifier Definitions  

All DBModifier definitions (done in XML) can be viewed by clicking the dialogBox Launcher on the right bottom corner of the Execute DBModifier Ribbon Group
together with Ctrl and Shift. This opens the Edit DBModifier Definitions Window:  
![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/EditDBModifDefinitions.PNG)

Here you can edit the definitions directly and also insert hidden features like the customized confirmation text in the element `confirmText`.

The DBModifiers can be executed either  

... using the Execute DBModifier Groups dropdown menus..  
![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/ConfigMenu.PNG)

... or using commandbuttons that were generated with the creation dialogs (the name of the control box has to be the same as the DBModifier definition/DBModifier Range)..  

... or be done on saving the Workbook.  

You can edit the DBModifiers either by Ctrl-Shift clicking the Execute DBModifier Groups dropdown menus..  
.. or by Ctrl-Shift clicking the created commandbuttons.  
.. or by using the Insert/Edit DBFunc/DBModif context menu within a DBMapper or DBAction range.  

## Hidden settings (not available in creation dialogs)

Following Settings of DBModifiers can only be edited in the Edit DBModifier Definitions Window:

*  `confirmText` (all DB Modifiers, String): an alternative text displayed in the confirmation of DB Modifiers.
*  `avoidFill` (only DBMappers, Boolean): prevent filling of whole table during execution of DB Mappers, this is useful for very large tables that are incrementally filled and would take unnecessary long time to start the DB Mapper. If set to true then each record is searched independently by going to the database. If the records to be stored are not too many, then this is more efficient than loading a very large table.

## Settings

Following Settings in DBAddin.xll.config or the referred DBAddinCentral.config affect behaviour of DBModifiers:
```xml
    <add key="CmdTimeout" value="30" />
    <add key="CnnTimeout" value="15" />
    <add key="DefaultEnvironment" value="3" />
    <add key="DontChangeEnvironment" value="False" />
    <add key="DBMapperCUDFlagStyle" value="TableStyleLight11" />
    <add key="DBMapperStandardStyle" value="TableStyleLight9" />
    <add key="DebugAddin" value="False" />
    <add key="maxNumberMassChange" value="10" />
    <add key="connIDPrefixDBtype" value="MSSQL" />
    <add key="DBSheetAutoname" value="True" />
```

Explanation:

*   `CmdTimeout`: the default timeout for a command to execute.
*   `CnnTimeout`: the default timeout for connecting.
*   `DefaultEnvironment`: default selected environment on startup.
*   `DontChangeEnvironment`: prevent changing the environment selector (Non-Production environments might confuse some people).
*   `DBMapperCUDFlagStyle`: Style for setting Excel data tables when having CUD Flags set on DBMappers (to find the correct name for this enter `? ActiveCell.ListObject.TableStyle` in the VBE direct window having selected a cell in the desirably formatted data table).
*   `DBMapperStandardStyle`: Style for setting Excel data tables when not having CUD Flags set on DBMappers.
*   `DebugAddin`: activate Info messages to debug addin.
*   `maxNumberMassChange`: Threshold of Number of changes in CUDFlag DBMappers to issue a warning.
*   `connIDPrefixDBtype`: Sometimes, legacy DBSheet definitions have a Prefix, this is the prefix to remove.
*   `DBSheetAutoname`: When inserting DBSheet Definitions, automatically name Worksheet to the table name, if this is set.
