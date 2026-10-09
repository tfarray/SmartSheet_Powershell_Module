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
  
  Please check the wiki for instructions
  
  
  Regards to all
  Thomas
