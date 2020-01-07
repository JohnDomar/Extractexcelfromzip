$TargetServer = "servername"
$TargetDb = 'databasename'
$ConnectedDb = Get-DbaDatabase -SqlInstance $TargetServer -Database $TargetDb      
$Folder = "D:\JDB\Scripts\proj_LabResults\3_staging_all\"



$Files = Get-ChildItem -Path $Folder -Filter *.xlsx  

If ($Files)
{

foreach ($f in $Files){
    $outfile = $f.FullName + "out" 
    Get-Content $f.FullName | Where-Object { ($_ -match 'step4' -or $_ -match 'step9') } | Set-Content $outfile
    Write-Output $f.FullName

    $ConvertToSqlParams = @{
            TableName = "schema.tablename"
            Path = $f.FullName
            ConvertEmptyStringsToNull = $true
            UseMSSQLSyntax = $true
            }


$SqlInsertStmts = ConvertFrom-ExcelToSqlInsert @ConvertToSqlParams

$ConnectedDb.Query($SqlInsertStmts)
#Write-Output $ConvertToSqlParams


}


$smtpServer = "SMTP.your address"
$smtpFrom = "yourfromaddress@your domain"
$smtpTo = "yourtoaddress@whereitsgoing"
$messageSubject = "Lab Test Results"
$messageBody = "New proj_LabResults Results Imported"
$smtp = New-Object Net.Mail.SmtpClient($smtpServer)
$smtp.Send($smtpFrom,$smtpTo,$messagesubject,$messagebody)

}

else 
{
$smtpServer = "SMTP.bradfordhospitals.int"
$smtpFrom = "BHTS-CONYDEVWD2@bthft.nhs.uk"
$smtpTo = "John.Birkinshaw@bthft.nhs.uk, Qamar.Zaman@bthft.nhs.uk, Alex.Newsham@bthft.nhs.uk"
$messageSubject = "Lab Test Results"
$messageBody = "No New proj_LabResults Results"
$smtp = New-Object Net.Mail.SmtpClient($smtpServer)
$smtp.Send($smtpFrom,$smtpTo,$messagesubject,$messagebody)
exit
}