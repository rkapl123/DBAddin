## DBModifications

DBModifications can be used to 
* save Excel Range data to database table(s): DBMapper
* modify DB Data using Data Manipulation SQL (update, delete, etc.): DBAction
* and creating sequences of these activites: DBSequence.

The target data referred to by DBMapper and DBAction (data is the DML SQL statement(s)) is specified by special Range names, any other definitions (environment, target database, etc.) is stored in a custom property of the workbook having the same name as the target range.

Examples for the usage of DBMapper can be found in the DBMapperTests.xlsx Workbook.

### Create DBModifiers

You can create the three DB Modifiers by using the cell context menu:  
![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/ContextMenu.PNG)  

The DBModifier creation/editing is shown below (examples already filled, when clicked on a blank cell all entries are empty):  

#### DB Mappers are created/edited with following dialog:  
![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/DbMapperCreate.PNG)  

*   DBMapper Name: Enter the name for the selected Range that will be used to identify the DBMap in the "Execute DBModifier" Group dropdowns. If no name is given here, then UnnamedDBMapper will be used to identify it.
*   Tablename: Database Table, where Data is to be stored.
*   Primary Keys: String containing primary Key names for updating table data, comma separated.
*   Database: Database to store DBMaps Data into
*   Ignore Columns: columns to be ignored (e.g. helper columns), comma separated.
*   Additional Stored Procedure: additional stored procedure to be executed after saving
*   Insert If Missing: if set, then insert row into table if primary key is missing there. Default = False (only update)
*   Store DBMap on Save: should DBMap also be saved on Excel Workbook Saving? (default no)
*   Environment: The Environment, where connection id should be taken from (if not existing, take from selected Environment in DB Addin General Settings Group)
*   Exec on Save: Should the DBMap be executed when the workbook is being saved?
*   Ask for execution: Before execution of the DBMap, ask for confirmation? A custom text can be given in the CustomXML definition element confirmText (see below)
*   C/U/D Flags: special mode used for row-by-row editing (inserting, updating and deleting rows). Only edited rows will be done when executing. Deleting rows is node with the special context menu item "delete Row" (or pressing Ctrl-Shift-D)
*   Ignore data errors: replace excel errors like #VALUE! with null on updating/inserting. Otherwise an error message is passed and execution is skipped for that row.
*   Create CB: create a commandbutton for the DB Sequence in the current Worksheet.
*   Hyperlink: click on it to highlight/select the DB Mapper area

You can always edit these parameters by selecting a cell in the DB Mapper area and invoking the context menu again.

#### DB Actions are created/edited with following dialog:  
![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/DbActionCreate.PNG)  

*   DBAction Name: Enter the name for the selected Range that will be used to identify the DBAction in the "Execute DBModifier" Group dropdowns. If no name is given here, then UnnamedDBAction will be used to identify it.
*   Database: Database to do the DBAction in
*   Environment: The Environment, where connection id should be taken from (if not existing, take from selected Environment in DB Addin General Settings Group)
*   Exec on Save: Should the DBAction be executed when the workbook is being saved?
*   Ask for execution: Before execution of the DBAction, ask for confirmation? A custom text can be given in the CustomXML definition element confirmText (see below)
*   The actual DBAction to be done is defined in a range that is named like the DBAction definition (the hyperlink takes you there). This range can be dynamically computed as all ranges in excel.
*   Create CB: create a commandbutton for the DB Sequence in the current Worksheet.
*   Hyperlink: click on it to highlight/select the DB Action range

You can always edit these parameters by selecting the Range of the DB Action area and invoking the context menu again.

#### DB Sequences are created/edited with following dialog:  
![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/DbSequenceCreate.PNG)  

*   DBSequence Name: Enter the name for the selected Range that will be used to identify the DBSequence in the "Execute DBModifier" Group dropdowns. If no name is given here, then UnnamedDBAction will be used to identify it.
*   Exec on Save: Should the DBSequence be executed when the workbook is being saved?
*   Ask for execution: Before execution of the DBSequence, ask for confirmation? A custom text can be given in the CustomXML definition element confirmText (see below)
*   Sequence Step Datagrid: here the available DBMappers, DBActions and DBFunctions (DBListfetch/DBRowFetch/DBSetQuery) can be added that are then executed in Sequence. If you are executing all sequence steps in the same environment, its possible to run the sequence in a transaction context by placing DBBegin at the top and DBCommitRollback at the bottom of the sequence.
*   ^/v buttons: used to move the sequence steps up and down.
*   Create CB: create a commandbutton for the DB Sequence in the current Worksheet.

As DB Sequences have no Range with data/definitions, invoking the context menu always creates new DB Sequences. You can edit existing DB Sequences 
by Ctrl-Shift clicking the Execute DBModifier Groups dropdown menus or by Ctrl-Shift clicking the created commandbuttons.

#### Edit DBModifier Definitions  

All DBModifier definitions (done in XML) can be viewed by clicking the dialogBox Launcher on the right bottom corner of the Execute DBModifier Ribbon Group
together with Ctrl and Shift. This opens the Edit DBModifier Definitions Window:  
![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/EditDBModifDefinitions.PNG)

Here you can edit the definitions directly and also insert hidden features like the confirmText. 

The DBModifiers can be executed either  

... using the Execute DBModifier Groups dropdown menus..  
![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/DBModifierMenu.PNG)

... or using commandbuttons that were generated with the creation dialogs (the name of the control box has to be the same as the DBModifier definition/DBModifier Range)..  

... or be done on saving the Workbook.  

You can edit the DBModifiers either by Ctrl-Shift clicking the Execute DBModifier Groups dropdown menus..  
.. or by Ctrl-Shift clicking the created commandbuttons  
.. or by using the Insert/Edit DBFunc/DBModif context menu within a DBMapper or DBAction range  
