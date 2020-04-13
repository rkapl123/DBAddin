
# Overview

DBAddin has three main areas and the Addin global code area, which are mapped to the following source files:

## DB Functions and their config files
* Functions.vb - Contains the public callable DB functions and helper functions and a data structure for transporting information back from the calculation action procedure to the calling function
* ConfigFiles.vb - procedures used for loading config files (containing DBFunctions and general sheet content) and building the config menu

## DB Modifiers
* DBModif.vb - DBModif Class: Abstraction of a DB Modification Object and descendant concrete classes DBMapper, DBAction or DBSeqnce; also contains global helper functions for DBModifiers
* DBModifCreate.vb  - Dialog for creating DB Modifier configurations
* EditDBModifDef.vb  - Dialog used to display and edit the CustomXMLPart utilized for storing the DBModif definitions, reused to also show DBAddin settings

## DBSheet definition creation and assignment
* DBSheetConfig.vb  - Helper module  for easier manipulation of DBSheet definition / Connection configuration data
* DBSheetCreateForm.vb  - Form for defining/creating DBSheet definitions

## Addin global code
* Globals.vb - Global variables and functions for DB Addin
* MenuHandler.vb - handles all Menu related aspects (context menu for building/refreshing, "DBAddin"/"Load Config" tree menu for retrieving stored configuration files, etc.)
* AboutBox.vb - About box: used to provide information about version/buildtime and links for local help and project homepage
* AddInEvents.vb - AddIn Connection class, also handling Events from Excel (Open, Close, Activate)

