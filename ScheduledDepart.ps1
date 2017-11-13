# Written by Jay Wang
# Released date: 09/11/2017

# PST path to be exported to
$PSTPath = '\\syd-fil-v01\ManualArchive\ArchivedUserPSTs'
# Used to get the date in users.csv
$Today = Get-Date
# Count how many users to depart for today
$UserCount = 0
# Account used to connect to Lync
$AdmName = 'caanz\svc_runningscript'
# Write a log which shows it has been run.
"
$(Get-Date -Format G): The script is started." >> .\logs.txt

# Read the users file which has all the to-be-departed users
try {
    $csv = Import-Csv .\users.csv
}
catch {
    "$_" >> .\logs.txt
    "$(Get-Date -Format G): Cannot find users.csv in this script folder." >> .\logs.txt
    Send-Email
    exit
}

# Check if there are users to depart and count them
foreach ($Line in $csv) {
    if (($Line.DepartDay.ToInt32($null) -eq $Today.Day) -and
        ($Line.DepartMonth.ToInt32($null) -eq $Today.Month) -and
        ($Line.DepartYear.ToInt32($null) -eq $Today.Year))
    {
        $UserCount += 1
    }
}

# Get the password from the file. Password is saved in SecretServer.
# The securepassword.txt is converted from the plain text password.
# This is only used for Lync remote connection.
try {
    $Password = Get-Content .\securepassword.txt
    $Securestring = $Password | ConvertTo-SecureString
    $Cred = New-Object System.Management.Automation.PSCredential `
        -ArgumentList $AdmName,$Securestring
}
catch {
    "$_" >> .\logs.txt
    "$(Get-Date -Format G): Cannot read the password from securepassword.txt" >> .\logs.txt
    Send-Email
    exit
}

# Email function to be used when task failed
function Send-Email {
    $From = 'ScheduledTask@charteredaccountantsanz.com'
    $To = 'jay.wang@charteredaccountantsanz.com'
    $Subject = 'Scheduled Task Failed'
    $Body = "The task Scheduled Depart on ECP-RMA-V03 is having errors.
    Please check the attachment or at
    \\ECP-RMA-V03\C$\ScheduledTasks\ScheduledDepart\logs.txt"
    $File = "$PWD\logs.txt"
    $SMTPServer = 'ecp-exch-v01'
    $Attachment = New-Object System.Net.Mail.Attachment($File)
    $MailMsg = New-Object System.Net.Mail.MailMessage
    $MailMsg.From = $From
    $MailMsg.To.Add($To)
    $MailMsg.Subject = $Subject
    $MailMsg.Body = $Body
    $MailMsg.Attachments.Add($Attachment)
    $SMTP = New-Object Net.Mail.SMTPclient($SMTPServer)
    $SMTP.Send($MailMsg)
}

# Function to connect to Exchange
# Kerberos is okay to use to connect to Exchange, but not Lync
function Connect-ExchangeAndLync {
    $ConnectionUri = 'http://ecp-exch-v01.caanz.com/powershell/'
    $SessionExch = New-PSSession -ConfigurationName Microsoft.Exchange `
        -ConnectionUri $ConnectionUri `
        -Authentication Kerberos `
        -ErrorAction Stop
    Import-PSSession $SessionExch -ErrorAction Stop

    # Lync has to use credential to connect.
    $ConnectionUri = 'https://lyncpool.caanz.com/ocspowershell'
    $SessionLync = New-PSSession -ConnectionUri $ConnectionUri `
        -Credential $Cred `
        -ErrorAction Stop
    Import-PSSession $SessionLync -ErrorAction Stop
}

# Function to disable the mailbox of a user and export PST as needed
function Disable-UserMailbox ($Name) {
    # Disable Unified Messaging
    Disable-UMMailbox -Identity $Name -Confirm:$false -ErrorAction Stop
    "$(Get-Date -Format G): Disabled UM mailbox $Name" >> .\logs.txt
    # Disable the mailbox
    Disable-Mailbox -Identity $Name -Confirm:$false -ErrorAction Stop
    "$(Get-Date -Format G): Disabled mailbox $Name" >> .\logs.txt
}

# Function to remove a user from Lync server
# Connect PS remote session to Lync are inclusive
# Disconnect PS remote session to Lync must be written outside this function
function Remove-LyncUser ($Name) {
    # Delete Lync user
    Disable-CsUser -Identity $Name -Confirm:$false -ErrorAction Stop
    "$(Get-Date -Format G): Removed Lync user $Name" >> .\logs.txt
}


# Establish connection if there are users to be departed today
if ($UserCount -gt 0) {
    try {
        Connect-ExchangeAndLync
        $ConnectFlag = $true
    }
    catch {
        $ConnectFlag = $false
    }
} else {
    $ConnectFlag = $false
}

if (($UserCount -gt 0) -and ($ConnectFlag -eq $true)) {
    # Check each line of the csv to perform actions based on the date and Y
    foreach ($Line in $csv) {
        if (($Line.DepartDay.ToInt32($null) -eq $Today.Day) -and
            ($Line.DepartMonth.ToInt32($null) -eq $Today.Month) -and
            ($Line.DepartYear.ToInt32($null) -eq $Today.Year))
        {
            # 1. When the mailbox is shared to others without any forwarding
            if (($Line.ExportPST -eq 'Y') -and ($Line.FwdEmail -eq 'N')) {
                try {
                # Create the export request
                $ExportRequest = New-MailboxExportRequest `
                    -Mailbox $Line.Alias `
                    -FilePath "$PSTPath\$($Line.Name).pst" `
                    -ErrorAction Stop
                }
                catch {
                    "$_" >> .\logs.txt
                    "$(Get-Date -Format G): $($Line.Alias) Failed to create export request" >> .\logs.txt
                    Send-Email
                    continue
                }
                # Check for completion
                do {
                    # Wait 60 seconds to let the export request run
                    Start-Sleep -Seconds 60

                    $Request = Get-MailboxExportRequest -Identity $ExportRequest.Identity
                    # Quit the loop once the export is failed
                    if ($Request.Status -in 'Completed','CompletedWithWarning') {
                        # Write log
                        "$(Get-Date -Format G): Exported PST to $PSTPath\$($Line.Name).pst" >> .\logs.txt
                        break
                    } elseif ($Request.Status -in 'Failed','AutoSuspended','None','Suspended','Synced') {
                        # When the export job status is failed or pending
                        "$(Get-Date -Format G): ExportPST failed for $($Line.Alias)." >> .\logs.txt
                        $ExportFlag = $false
                        Send-Email
                        break
                    } else {
                        continue
                    }
                } while ($true)

                if ($ExportFlag -eq $false) {
                    continue
                }

                try {
                    # Disable the mailbox after the export is completed
                    Disable-UserMailbox -Name $Line.Alias
                }
                catch {
                    "$_" >> .\logs.txt
                    "$(Get-Date -Format G): Failed to disable mailbox $($Line.Alias)" >> .\logs.txt
                    Send-Email
                }
            }
            # 2. When mailbox is forwared but not requiring exporting PST which
            # has already been done during normal departing user process
            elseif (($Line.ExportPST -eq 'N') -and ($Line.FwdEmail -eq 'Y')) {
                try {
                    Disable-UserMailbox -Name $Line.Alias
                }
                catch {
                    "$_" >> .\logs.txt
                    "$(Get-Date -Format G): Failed to disable mailbox $($Line.Alias)" >> .\logs.txt
                    Send-Email
                }
            }
            # 3. When mailbox is set to Export and Forwarding. Not expected
            elseif (($Line.ExportPST -eq 'Y') -and ($Line.FwdEmail -eq 'Y')) {
                "$(Get-Date -Format G): $($Line.Name)Cannot have both ExportPST and FwdEmail set to Y.
                        Export must be done when enabling email forwarding.
                        Please manually complete the departing." >> .\logs.txt
            }
            # 4. When both is set to N, nothing is processed
            elseif (($Line.ExportPST -eq 'N') -and ($Line.FwdEmail -eq 'N')) {
                "$(Get-Date -Format G): $($Line.Alias) Both ExportPST and FwdEmail are set to N" >> .\logs.txt
            }

            # Check Y to remove Lync user
            if ($Line.FwdLync -eq 'Y') {
                try {
                    $UPN = (Get-ADUser -Identity $Line.Alias).UserPrincipalName
                    Remove-LyncUser -Name $UPN
                }
                catch {
                    "$_" >> .\logs.txt
                    "$(Get-Date -Format G): Failed to disable Lync user $($Line.Alias)" >> .\logs.txt
                    Send-Email
                }
            }
        }
    }
}

if ($ConnectFlag) {
    try {
        Remove-PSSession -Session $SessionExch,$SessionLync -ErrorAction Stop
    }
    catch {
        Get-PSSession | Remove-PSSession
    }
    finally {
        "$(Get-Date -Format G): The script is completed." >> .\logs.txt
    }
} else {
    "$(Get-Date -Format G): No user to be departed" >> .\logs.txt
}