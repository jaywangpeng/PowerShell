$ExportLocation = '<UNC path>'
$FileName = "ADUserReport-$(Get-Date -Format yyyyMMdd).csv"

Get-ADUser -SearchBase "OU=People,OU=Accounts,OU=PROD,DC=caanz,DC=com" `
    -Filter { (Enabled -eq $True) -and (PasswordNeverExpires -eq $True) } `
    -Properties Givenname,Surname,Manager,Department,Description,City |
Select-Object -Property Name,SAMAccountName,Description,Department,City,Manager |
Export-Csv -Path "$ExportLocation$FileName"