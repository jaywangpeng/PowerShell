Import-Module ActiveDirectory

#RefUser is the one who you want to check what groups he/she belongs to.
$RefUserInput = Read-Host "What is the account of the referenced user(e.g. jwang)?"
$RefUser = Get-ADUser -Identity $RefUserInput -Properties SAMAccountName,Memberof

#TargetUser is the one who you are adding to groups.
$TargetUserInput = Read-Host "What is the account of the user you want to add?"
$TargetUser = Get-ADUser -Identity $TargetUserInput -Properties SAMAccountName,Memberof

#Add the target user to the ref user's Memberof groups
$GroupMemberShip = ($RefUser.Memberof | Foreach-Object {(Get-ADGroup $_).Name;})
$GroupMemberShip | Foreach-Object {Add-ADGroupMember -Identity $_ -Member $TargetUser.SAMAccountName}
