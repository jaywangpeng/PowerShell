#RefUser is the one who you want to check what groups he/she belongs to.
$RefUser = Read-Host "What is the user's Full Name?"
$User = Get-ADUser -Identity $RefUser -Properties *
$GroupMemberShip = ($User.Memberof | %{(Get-ADGroup $_).Name;}) -Join ';'
$GroupMemberShip 
