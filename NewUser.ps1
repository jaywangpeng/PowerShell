#Where the new will be created
$OUPath = 'OU=People,OU=Accounts,OU=UAT,DC=caanz,DC=com'

#Functions
function CreateNewUser(){
     New-ADUser `
    -SamAccountName $FirstName `
    -GivenName $FirstName `
    -Surname $LastName `
    -DisplayName $FirstName $LastName `
    -Path $OUofUser
}


$FirstName = Read-Host -Prompt "Please enter the user's First Name (Given Name)"
$LastName = Read-Host -Prompt "Please enter the user's Last Name (Surname)"
#Validate the name
if ($FirstName -match "^[a-zA-Z ,.'-]+*") {
    if ($LastName -match "^[a-zA-Z ,.'-]+*")
    }
else {

}
