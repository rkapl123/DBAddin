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

The DBModifier (in this case a DBMapper) creation starts following dialog (already filled, when clicked on a blank cell all entries are empty):  
![image](https://raw.githubusercontent.com/rkapl123/DBAddin/master/docs/image/DbMapperCreate.PNG)  

*   DBMapper Name: Enter the name for the selected Range that will be used to identify the DBMap in the "store DBMapper Data" Group dropdowns. If no name is given here, then UnnamedDBMapper will be used to identify it.
*   Tablename: Database Table, where Data is to be stored.
*   Primary Keys: String containing primary Key names for updating table data, comma separated.
*   Database: Database to store DBMaps Data into
*   Ignore Columns: columns to be ignored (e.g. helper columns), comma separated.
*   Additional Stored Procedure: additional stored procedure to be executed after saving
*   Insert If Missing: if set, then insert row into table if primary key is missing there. Default = False (only update)
*   Store DBMap on Save: should DBMap also be saved on Excel Workbook Saving? (default no)
*   Environment: The Environment, where connection id should be taken from (if not existing, take from selected Environment in DB Addin General Settings Group)

You can always edit these parameters by selecting this cell and invoking the context menu again.

So for the parameters shown in above creation dialog, following definition parameter is created in a custom property DBMapperTest (MSSQL is environment 3 in my settings):
`def(3,"TestTable","TestId,TestId2","TestDB",True,"TestProc","TestHelper,Lookup,Dummy",True)` 
