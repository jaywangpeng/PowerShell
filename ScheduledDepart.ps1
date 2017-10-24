$Today = Get-Date
$FormatedDate = Get-Date -Format G

# Read the users file which has all the to-be-departed users
try
{
    $csv = Import-Csv .\users.csv
}
catch
{
    "`n$FormatedDate`: Cannot find users.csv in this script folder." >> .\logs.txt
    exit
}

# Write a log which shows it has been run.
"`n$FormatedDate`: The script is started." >> .\logs.txt

# Function to disable the mailbox of a user
function Disable-UserMailbox ($Name)
{
    $ConnectionUri = 'http://ecp-exch-v01.caanz.com/powershell/'
    $SessionExch = New-PSSession -ConfigurationName Microsoft.Exchange `
        -ConnectionUri $ConnectionUri `
        -Credential $Cred `
        -ErrorAction Stop
    Import-PSSession $SessionExch

    # Disable the mailbox
    Disable-Mailbox -Identity $Name -Confirm:$true
    "$FormatedDate`: Disabled mailbox $Name" >> .\logs.txt
    Remove-PSSession $SessionExch
}

# Function to remove a user from Lync server
function Remove-LyncUser ($Name)
{
    # Make the connection to Lync
    $ConnectionUri = 'https://lyncpool.caanz.com/ocspowershell'
    $SessionLync = New-PSSession -ConnectionUri $ConnectionUri `
        -Credential $Cred `
        -ErrorAction Stop
    Import-PSSession $SessionLync -ErrorAction Stop

    # Delete Lync user
    Disable-CsUser -Identity $Name
    "$FormatedDate`: Removed Lync user $Name" >> .\logs.txt
    Remove-PSSession $SessionLync
}

# Get the password from the file
try {
    $Password = Get-Content .\securepassword.txt
    $Securestring = $Password | ConvertTo-SecureString
    $Cred = New-Object System.Management.Automation.PSCredential `
        -ArgumentList 'caanz\svc_runningscript',$Securestring
}
catch {
    "$_" >> .\logs.txt
}

# Check each line of the csv to perform actions based on the date and Y
foreach ($Line in $csv)
{
    if( ($Line.DepartDay.ToInt32($null) -eq $Today.Day) -and
        ($Line.DepartMonth.ToInt32($null) -eq $Today.Month) -and
        ($Line.DepartYear.ToInt32($null) -eq $Today.Year) )
    {
        # Check Y to disable mailbox
        if($Line.FwdEmail -eq 'Y')
        {
            try { Disable-UserMailbox -Name $Line.Alias }
            catch { "$_" >> .\logs.txt }
        }

        # Check Y to remove Lync user
        if($Line.FwdLync -eq 'Y')
        {
            try { Remove-LyncUser -Name $Line.Alias }
            catch { "$_" >> .\logs.txt }
        }

        # Check if both are N, write a log. No actions.
        if( ($Line.FwdEmail -eq 'N') -and ($Line.FwdLync -eq 'N') )
        {
            "User $($Line.Alias) is not set Y to FwdEmail or FwdLync." >> .\logs.txt
        }
    }
}

"$FormatedDate`: The script is completed." >> .\logs.txt
