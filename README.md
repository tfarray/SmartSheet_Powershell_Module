# SmartSheet_Powershell_Module
SmartSheet module for Powershell
I built this repository to share, but honestly, I'm not sure to have time to improve it
It is not Perfect
It is not implementing all the functions
Not sure to be able to bring in new features

This module intent is to be able to use the most basic functions directly from a powershell command line.

To get this module working you will need :
- Powershell v3 or above
- A smartsheet API token (get it from smartsheet)

To be able to use this module you need :
- To copy directory SmartSheet either :
  (1) To any path
  (2) To a built in module path. To get thoses paths, type the powershell command > $env:psmodulePath -split ";"
- You might need to run this command to get the module working : 
  > Set-ExecutionPolicy Bypass
- To import the module
 (1) > import-module FQDN_Path\SmartSheet\smartsheet.psm1
 (2) > import-module SmartSheet
 - Once the module is imported, you need to store it (it will store encrypted using windows Data protection API)
  > Set-SmartSheetAPIToken [your Token]
  
  To use the Module :
  > get-smartsheet
  Will show you all the smartsheets you are able to read.
  
  > $MySS = get-smartsheet MySmartsheet*
  This will load (only if there is a single SmartSheet corresponding to the name) the smartsheet to the variable $MySS
  
  > $MySS = get-smartsheet -ID 123412341234
  This will load the EXACT smartsheet to the variable.
  
  > $MySS.table
  This will show you your smartsheet
  
  If you have hiearchy starting at line 3
  > $MySS.table[2].__Childnode
  This will show the Child nodes
  
  > $MySS.table[5].ColumnName = "New Value"
  > $MySS.table[5].update()
  Those 2 commands will update the smartsheet at line 6, the column called "ColumnName" and set the cell to the value "New Value"
  
  
  Regards to all
  Thomas
