This file lists the pre-requisites to run this process

First - you will need Import-Excel and dbatools imported into Powershell 

see:
https://www.sqlshack.com/dbatools-powershell-module-for-sql-server/ 
https://github.com/dfinke/ImportExcel


open powershell as administrator

then run command:
Install-Module -Name DBATools
say yes at any prompts

then run command:
set-executionpolicy remotesigned

then run command:
Install-Module -Name ImportExcel -RequiredVersion 5.4.2

Then map P drive using at command prompt:

net use P: "\\myremotelocation\"

This must be the file containing the zipped excel files you require

Then run 1_RunImport.bat to import the files. some variables and paths will need amending for your environment




