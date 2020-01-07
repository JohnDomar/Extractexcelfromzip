
rem starts powershell that unzips files to D:\JDB\Scripts\proj_LabResults\1_staging_1\ and 2 etc
rem ready to import

Powershell.exe -executionpolicy remotesigned -File D:\JDB\Scripts\proj_LabResults\scripts\2_UnzipperALL_V2.ps1

rem - now itterates through the folder and imports each xls sheet if it exists 
Powershell.exe -executionpolicy remotesigned -File D:\JDB\Scripts\proj_LabResults\scripts\3_ImportExcelDirect_V2.ps1

rem tidies up folder ready for next files.  

move P:\*.zip P:\Imported\
copy D:\JDB\Scripts\proj_LabResults\3_staging_all\*.* D:\JDB\Scripts\proj_LabResults\4_files_processed\

del /Q D:\JDB\Scripts\proj_LabResults\1_staging_1\*.*
del /Q D:\JDB\Scripts\proj_LabResults\2_staging_2\*.*
del /Q D:\JDB\Scripts\proj_LabResults\3_staging_all\*.*
