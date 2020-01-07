set-alias sz "$env:ProgramFiles\7-Zip\7z.exe"
#input path is the source location of the zip files
$Inputpath = "P:"
$Password1 = "password1" #password 
$UnzipPath1 = "D:\JDB\Scripts\proj_LabResults\1_staging_1"
$Password2 = "password2" #password 
$UnzipPath2 = "D:\JDB\Scripts\proj_LabResults\2_staging_2"
$UnzipPathALL = "D:\JDB\Scripts\proj_LabResults\3_staging_all"

$CheckFile = Get-ChildItem -Path $Inputpath -Filter *.zip -Recurse 

if ($CheckFile) 
{

Write-Output "It exists"
#Tries to unzip into outputfolder 
sz x $Inputpath\*.zip "-p$Password1" "-o$UnzipPath1"
sz x $Inputpath\*.zip "-p$Password2" "-o$UnzipPath2"

#copies files greater than 0kb to output folder

Get-ChildItem $UnzipPath1 |where { $_.Length -gt 0KB}| copy-item -Destination $UnzipPathALL
Get-ChildItem $UnzipPath2 |where { $_.Length -gt 0KB}| copy-item -Destination $UnzipPathALL


}

else 

{
Write-Output "It doesn't exist" 
exit
}





