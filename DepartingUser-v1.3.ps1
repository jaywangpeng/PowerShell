<# Departing User script
Developed by Jay Wang
Version: 1.3
Release Date: 10/11/2017
#>

# Disable and move AD account
function Disable-ADaccount([string]$Name)
{
# Step 1: Disable account
    Set-ADUser -Identity $Name -Enabled $False
    Write-Host "Disabled account $($objAccount.Name)" -ForegroundColor Green

# Step 2: Move to the OU which is named with the month of 3 months later
    $OU = 'OU=Disabled,OU=Accounts,OU=PROD,DC=caanz,DC=com'
    $OUPrefix = 'OU=To Be Deleted -'
    switch((Get-Date).AddMonths(3).Month)
    {
        1 {$MonthOU = "$OUPrefix January,$OU"}
        2 {$MonthOU = "$OUPrefix February,$OU"}
        3 {$MonthOU = "$OUPrefix March,$OU"}
        4 {$MonthOU = "$OUPrefix April,$OU"}
        5 {$MonthOU = "$OUPrefix May,$OU"}
        6 {$MonthOU = "$OUPrefix June,$OU"}
        7 {$MonthOU = "$OUPrefix July,$OU"}
        8 {$MonthOU = "$OUPrefix August,$OU"}
        9 {$MonthOU = "$OUPrefix September,$OU"}
        10 {$MonthOU = "$OUPrefix October,$OU"}
        11 {$MonthOU = "$OUPrefix November,$OU"}
        12 {$MonthOU = "$OUPrefix December,$OU"}
        default {
            Write-Host "Cannot get the correct date from the system" `
                -ForegroundColor Red
            break
        }
    }
# Move the account
    Move-ADObject -Identity $objAccount.DistinguishedName -TargetPath $MonthOU
    Write-Host "Moved account $($objAccount.Name) to " -NoNewLine `
        -ForegroundColor Green
    Write-Host "$MonthOU" -ForegroundColor Green

# Step 3: Get all the groups of this account
    $UserGroups = $objAccount.MemberOf
# Remove all groups. Domain Users group will not be deleted by default.
    foreach($Group in $UserGroups)
    {
        Get-ADGroup $Group |
        Remove-ADGroupMember -Members $Name -Confirm:$False
    }
    Write-Host -ForegroundColor Green `
        "Removed the group membership of account $($objAccount.Name)"

# Step 4: Remove the manager in Orgnization tab
    try
    {
        $objManager = Get-ADUser $objAccount.Manager -ErrorAction Stop
        Set-ADUser -Identity $Name -Manager $Null -ErrorAction Stop
        Write-Host -ForegroundColor Green `
        "Removed the manager $($objManager.Name) from $($objAccount.Name)"
    }
    catch
    {
        Write-Host "User $Name has no manager set to it."
    }
}

# Move folder and rename it. For Archive-Folders function
function Move-Folder
{
    [CmdletBinding()]param(
        [parameter(position=0)][string]$Src,
        [parameter(position=1)][string]$Dst,
        [parameter(position=2)][string]$Name,
        [parameter(position=3)][string]$Folder
    )
    Write-Host "Using robocopy.exe to move $Src\$Name..." `
        -ForegroundColor Cyan
    Robocopy.exe "$Src\$Name" "$Dst\$Name\$Name" /MOVE /E
    Rename-Item "$Dst\$Name\$Name" -NewName $Folder -ErrorAction Stop
}

# Backup userdata folders to archive location #
function Archive-Folders([string]$Name)
{
# Source base path PROD
# Homedir folders
    $HomedirSYD = '\\caanz.com\userdata\homedirs'
    $HomedirMEL = '\\MEL-FIL-V06\homedirs$'
    $HomedirADL = '\\ADL-FIL-V05\homedirs$'
    $HomedirBNE = '\\BNE-FIL-V04\homedirs$'
    $HomedirCBR = '\\CBR-FIL-V07\homedirs$'
    $HomedirPER = '\\PER-FIL-V08\homedirs$'
    $HomedirNZD = '\\NZD-FIL-V09\home$'
# Documents folders
    $DocumentsSYD = '\\caanz.com\userdata\Documents'
    $DocumentsMEL = '\\MEL-FIL-V06\documents$'
    $DocumentsADL = '\\ADL-FIL-V05\documents$'
    $DocumentsBNE = '\\BNE-FIL-V04\documents$'
    $DocumentsCBR = '\\CBR-FIL-V07\documents$'
    $DocumentsPER = '\\PER-FIL-V08\documents$'

# Desktop folders
    $DesktopSYD = '\\caanz.com\userdata\desktops'
    $DesktopMEL = '\\MEL-FIL-V06\desktop$'
    $DesktopADL = '\\ADL-FIL-V05\desktop$'
    $DesktopBNE = '\\BNE-FIL-V04\desktop$'
    $DesktopCBR = '\\CBR-FIL-V07\desktop$'
    $DesktopPER = '\\PER-FIL-V08\desktop$'

# Destination base path PROD
    $DestBase = '\\syd-fil-v01\ManualArchive\ArchivedUserData'

    if(Test-Path -Path "$DestBase\$Name")
    {
        Write-Host "$DestBase\$Name already exists. Operation terminated." `
            -ForegroundColor Red
        break
    }
    else
    {
        try
        {
            New-Item -Path $DestBase -Name $Name -ItemType Directory `
                -ErrorAction Stop
            $FlagNewFolder = $True
        }
        catch
        {
            Write-Host "Error creating new folder $DestBase\$Name" `
                -ForegroundColor Red
            $FlagNewFolder = $False
            break
        }
        if( $FlagNewFolder -eq $True )
        {
            Write-Host "Folder $DestBase\$Name has been created" `
                -ForegroundColor Cyan
            Write-Host "Archiving userdata folders to $DestBase\$Name" `
                -ForegroundColor Cyan
            try
            {
                # Set the move flag to false before moving anything
                $FlagMove = $False
                switch( $($objAccount.HomeDirectory.Substring(2, 3)) )
                {
                    'CAA' {
                        Move-Folder $HomedirSYD $DestBase $Name 'homedir' `
                            -ErrorAction Stop
                        Move-Folder $DocumentsSYD $DestBase $Name 'documents' `
                            -ErrorAction Stop
                        Move-Folder $DesktopSYD $DestBase $Name 'desktop' `
                            -ErrorAction Stop
                    }
                    'MEL' {
                        Move-Folder $HomedirMEL $DestBase $Name 'homedir' `
                            -ErrorAction Stop
                        Move-Folder $DocumentsMEL $DestBase $Name 'documents' `
                            -ErrorAction Stop
                        Move-Folder $DesktopMEL $DestBase $Name 'desktop' `
                            -ErrorAction Stop
                    }
                    'ADL' {
                        Move-Folder $HomedirADL $DestBase $Name 'homedir' `
                            -ErrorAction Stop
                        Move-Folder $DocumentsADL $DestBase $Name 'documents' `
                            -ErrorAction Stop
                        Move-Folder $DesktopADL $DestBase $Name 'desktop' `
                            -ErrorAction Stop
                    }
                    'BNE' {
                        Move-Folder $HomedirBNE $DestBase $Name 'homedir' `
                            -ErrorAction Stop
                        Move-Folder $DocumentsBNE $DestBase $Name 'documents' `
                            -ErrorAction Stop
                        Move-Folder $DesktopBNE $DestBase $Name 'desktop' `
                            -ErrorAction Stop
                    }
                    'CBR' {
                        Move-Folder $HomedirCBR $DestBase $Name 'homedir' `
                            -ErrorAction Stop
                        Move-Folder $DocumentsCBR $DestBase $Name 'documents' `
                            -ErrorAction Stop
                        Move-Folder $DesktopCBR $DestBase $Name 'desktop' `
                            -ErrorAction Stop
                    }
                    'PER' {
                        Move-Folder $HomedirPER $DestBase $Name 'homedir' `
                            -ErrorAction Stop
                        Move-Folder $DocumentsPER $DestBase $Name 'documents' `
                            -ErrorAction Stop
                        Move-Folder $DesktopPER $DestBase $Name 'desktop' `
                            -ErrorAction Stop
                    }
                    'NZD' {
                        Move-Folder $HomedirNZD $DestBase $Name 'home' `
                            -ErrorAction Stop
                    }
                }
# Flag the move as true when completed
                $FlagMove = $True
            }
            catch
            {
                $FlagMove = $False
                Write-Host "$_" -ForegroundColor Red
                Write-Host "Moving folders failed. Check the folders." `
                    -ForegroundColor Red
                break
            }
            if( $FlagMove -eq $True )
            {
                Write-Host "Archiving folders has been completed." `
                    -ForegroundColor Cyan
            }
        }
    }
}

# Connect to Exchange
function Connect-Exchange
{
    $ConnectionUri = 'http://ecp-exch-v01.caanz.com/powershell/'
    $Cred = Get-Credential -UserName $ENV:USERNAME `
        -Message 'Enter your Exchange admin credential'
    Write-Host "Connecting to Exchange at ECP-EXCH-V01..." -ForegroundColor Cyan
    $Script:Session = New-PSSession -ConfigurationName Microsoft.Exchange `
        -ConnectionUri $ConnectionUri `
        -Authentication Kerberos `
        -Credential $Cred `
        -ErrorAction Stop
}

# Connect to Lync
function Connect-Lync
{
    $ConnectionURI = 'https://lyncpool.caanz.com/ocspowershell'
    $Cred = Get-Credential -UserName $ENV:USERNAME `
        -Message 'Enter your Lync admin credential'
    Write-Host "Connecting to Lync at lyncpool.caanz.com..." -ForegroundColor Cyan
    $Script:Session = New-PSSession -ConnectionUri $ConnectionURI `
        -Credential $Cred `
        -ErrorAction Stop
}

# Connect to Exchange and export a user mailbox to pst file
function Export-PST ([string]$Name)
{
    $PSTPath = '\\syd-fil-v01\ManualArchive\ArchivedUserPSTs'
    try
    {
        Connect-Exchange
        Import-PSSession $Session -ErrorAction Stop 3> $null
        Write-Host "Connected to Exchange at ECP-EXCH-V01." `
            -ForegroundColor Green

        New-MailboxExportRequest -Mailbox $Name `
            -FilePath "$PSTPath\$($objAccount.Name).pst"

        Write-Host "The export job has been created. Find the file at:" `
            -ForegroundColor Green
        Write-Host "$PSTPath\$($objAccount.Name).pst" -ForegroundColor Green
        Write-Host "No notification will be sent to you once it's completed." `
            -ForegroundColor Yellow
    }
    catch
    {
        Write-Host "$_" -ForegroundColor Red
    }
    finally
    {
        Remove-PSSession $Session
        Write-Host "Disconnected from Exchange." -ForegroundColor Green
    }
}

# Disable the user's mailbox from Exchange
function Disable-UserMailbox ($Name)
{
    Connect-Exchange
    Import-PSSession $Session -ErrorAction Stop 3> $null

    try {
        # Disable Unified Messaging
        Disable-UMMailbox -Identity $Name -Confirm:$false -ErrorAction Stop
        Write-Host "Unified Mailbox of $Name has been disabled." -ForegroundColor Green
        # Disable the mailbox
        Disable-Mailbox -Identity $Name -Confirm:$false -ErrorAction Stop
        Write-Host "Mailbox of $Name has been disabled." -ForegroundColor Green
    }
    catch {
        Write-Host "$_" -ForegroundColor Red
    }
    finally {
        Remove-PSSession $Session
    }
}

# Function to remove an account from Lync server
function Remove-LyncUser ($Name)
{
    Connect-Lync
    Import-PSSession $Session -ErrorAction Stop 3> $null

    try {
        # Delete Lync user
        Write-Host "Start disable-csuser" -ForegroundColor Cyan
        Disable-CsUser -Identity $Name -Confirm:$false -ErrorAction Stop
        Write-Host "Removed Lync user $Name" -ForegroundColor Green
    }
    catch {
        Write-Host "$_" -ForegroundColor Red
    }
    finally {
        Remove-PSSession $Session
    }
    Write-Host "Disconnected from Lync"
}

# Set the Name to forward emails to FwdName
function Forward-Email ([string]$Name)
{
# Get the logon name of forwarded-to user
    while($true)
    {
        try
        {
# Get the AD user object
            [string]$FwdName = Read-Host "The logon name to be forwarded to?" `
                -ErrorAction Stop
            $objFwdAccount = Get-ADUser -Identity $FwdName -Properties Mail `
                -ErrorAction Stop

# Check whether the input is correct
            Write-Host "Please confirm: Forward" -ForegroundColor Cyan -NoNewLine
            Write-Host " $Name " -ForegroundColor Green -NoNewLine
            Write-Host "to" -ForegroundColor Cyan -NoNewLine
            Write-Host " $($objFwdAccount.Mail)" -ForegroundColor Green
            [string]$FwdConfirm = Read-Host "Enter Y/y or N/n"
            if($FwdConfirm -eq 'Y')
            {
                break
            }
        }
        catch
        {
            Write-Host "$_" -ForegroundColor Red
            Write-Host "Please enter again" -ForegroundColor Cyan
        }
    }
# Forward the email to that user
    try
    {
        Connect-Exchange
        Import-PSSession $Session -ErrorAction Stop 3> $null
        Write-Host "Connected to Exchange at ECP-EXCH-V01." `
            -ForegroundColor Green
        Write-Host "Forwarding email to $($objFwdAccount.Mail)" `
            -ForegroundColor Cyan
        Get-Mailbox -Identity $Name |
            Set-Mailbox -ForwardingAddress "$($objFwdAccount.Mail)"
    }
    catch
    {
        Write-Host "$_" -ForegroundColor Red
        Write-Host "Please check and rerun the script." -ForegroundColor Red
        break
    }
    finally
    {
        Remove-PSSession $Session
        Write-Host "Disconnected from Exchange." -ForegroundColor Green
    }
}

function Forward-Lync
{
# Asking for who to forward to
    [string]$LyncFwdNumber = Read-Host "Enter the phone number to be forwarded to"

# Lync SEFAUtil.exe tool parameters and combine to one command $Cmd
    $LyncTool = "'C:\Program Files\Microsoft Lync Server 2013\ResKit\SEFAUtil.exe'"
    $Param1 = '/server:lyncpool.caanz.com'
    $Param2 = "$($objAccount.EmailAddress)"
    $Param3 = "/setfwddestination:$LyncFwdNumber"
    $Param4 = '/enablefwdimmediate'
    $Cmd = "& $LyncTool $Param1 $Param2 $Param3 $Param4"
    $Cmd | clip.exe
# Tell user to copy the following command to an locally elevated window
# on Lync server. Use clip.exe to paste to Clipboard.
    $Comments = "
The following command is copied to clipboard.
Please RDP to ECP-UC-V01 and start an elevated PowerShell. Paste and it will automatically execute.`n"
    Write-Host $Comments -ForegroundColor Yellow
    Write-Host $Cmd -ForegroundColor Cyan
}

# Execute Option 1,2,3
function Normal-Depart( [string]$Name )
{
    Write-Host "Executing step 1..." -ForegroundColor Cyan
    Disable-ADaccount $Name
    Write-Host "Executing step 2..." -ForegroundColor Cyan
    Archive-Folders $Name
    Write-Host "Executing step 3..." -ForegroundColor Cyan
    Export-PST $Name
    Write-Host "Step 6 is skipped. Please run it separately after PST is exported." -ForegroundColor Yellow
    Start-Sleep -Seconds 5
    Write-Host "Executing step 7..." -ForegroundColor Cyan
    Remove-LyncUser $Name
}

############-----> Script starts from here <-----############
# Get the username from input and verify username
while($True)
{
    try
    {
        Write-Host "What's the user logon name to be departed? (Type q to exit)"
        [string]$InputName = Read-Host -ErrorAction Stop
    }
    catch
    {
        Write-Host "$_" -ForegroundColor Red
    }
    if($InputName -match '^[qQ]$')
    {
        Write-Host "Script exited." -ForegroundColor Cyan
        exit
    }
    elseif($InputName.ToLower() -match '^[a-z]+[-_]?[a-z]+[1-9]*$')
    {
        try
        {
            $Script:objAccount = Get-ADUser $InputName -Properties * `
                -ErrorAction Stop
            break
        }
        catch
        {
            Write-Host "Cannot find the specified user in AD." `
                -ForegroundColor Red
        }
    }
    else
    {
        Write-Host "Username is invalid. Type q to exit." `
            -ForegroundColor Red
    }
}

####### Departing User Menu #######
$Menu = "
0. Exit
1. AD account operations:
    Disable AD account
    Move to Monthly Disabled OU (3 months)
    Remove group membership
    Remove manager
2. Archive userdata (homedir, documents, desktop)
3. Export PST to archive location
4. Forward Email to someone
5. Forward Lync call to someone
6. Disable the user's Unified Messaging and mailbox
7. Remove the user from Lync
8. Execute step: 1,2,3,6,7"

$Caution = "
Caution!!
Steps here are irreversible and non-repeatable.
Not all departing user steps can be done here.
Don't forget other steps in Departing User process!!"

Write-Host "`nYou are departing: " -NoNewLine -ForegroundColor Cyan
Write-Host "$($objAccount.Name)" -ForegroundColor Green
Write-Host "`nDeparting User Menu:" -ForegroundColor Cyan
Write-Host $Caution -ForegroundColor Yellow
Write-Host $Menu -ForegroundColor Cyan

while($True)
{
    try
    {
        [int]$Option = Read-Host "Please enter a number from the menu" `
                -ErrorAction Continue
    }
    catch
    {
        Write-Host "$_" -ForegroundColor Red
    }
    try
    {
        switch($Option)
        {
            0 { Write-Host "Script exited." -ForegroundColor Cyan
                exit
              }
            1 { Disable-ADaccount $objAccount.SAMAccountName }
            2 { Archive-Folders $objAccount.SAMAccountName }
            3 { Export-PST $objAccount.SAMAccountName }
            4 { Forward-Email $objAccount.SAMAccountName }
            5 { Forward-Lync }
            6 { Disable-UserMailbox $objAccount.SAMAccountName }
            7 { Remove-LyncUser $objAccount.UserPrincipalName }
            8 { Normal-Depart $objAccount.SAMAccountName }
            default { Write-Host "Input must be a number from 0 to 8" `
                        -ForegroundColor Cyan
                    }
        }
    }
    catch
    {
        Write-Host "$_" -ForegroundColor Red
    }
    Write-Host $Menu -ForegroundColor Cyan
}