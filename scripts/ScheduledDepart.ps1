# Written by Jay Wang
# Released date: 14/11/2017
# Version 1.0

# Used to get the date in users.csv
$Today = Get-Date
# Counters of users, mailboxes, Lync accounts to depart for today
$UserCount = 0
$MailboxCount = 0
$LyncCount = 0
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

# Check how many users, mailboxes, lync accounts to depart and count them
foreach ($Line in $csv) {
    if (($Line.DepartDay.ToInt32($null) -eq $Today.Day) -and
        ($Line.DepartMonth.ToInt32($null) -eq $Today.Month) -and
        ($Line.DepartYear.ToInt32($null) -eq $Today.Year))
    {
        $UserCount += 1
        if (($Line.ExportPST -eq 'Y') -or ($Line.FwdEmail -eq 'Y')) {
            $MailboxCount += 1
        }
        if ($Line.FwdLync -eq 'Y') {
            $LyncCount += 1
        }
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
    $To = 'ICTOperationsTeam@charteredaccountantsanz.com'
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
    if ($MailboxCount -gt 0) {
        $ConnectionUri = 'http://ecp-exch-v01.caanz.com/powershell/'
        $SessionExch = New-PSSession -ConfigurationName Microsoft.Exchange `
            -ConnectionUri $ConnectionUri `
            -Authentication Kerberos `
            -ErrorAction Stop
        Import-PSSession $SessionExch -ErrorAction Stop
    }

    if ($LyncCount -gt 0) {
        # Lync has to use credential to connect.
        $ConnectionUri = 'https://lyncpool.caanz.com/ocspowershell'
        $SessionLync = New-PSSession -ConnectionUri $ConnectionUri `
            -Credential $Cred `
            -ErrorAction Stop
        Import-PSSession $SessionLync -ErrorAction Stop
    }
}

function Export-PST ($Name, $FullName) {
    # PST path to be exported to
    $PSTPath = '\\syd-fil-v01\ManualArchive\ArchivedUserPSTs'
    try {
        # Create the export request
        $ExportRequest = New-MailboxExportRequest -Mailbox $Name -FilePath "$PSTPath\$FullName.pst" `
            -ErrorAction Stop
    }
    catch {
        "$(Get-Date -Format G): Failed to create export request $Name" >> .\logs.txt
        Send-Email
        return $false
    }
    do {
        # Wait 60 seconds to let the export request run
        Start-Sleep -Seconds 60
        # Query the export status
        $Request = Get-MailboxExportRequest -Identity $ExportRequest.Identity
        if ($Request.Status -in 'Completed','CompletedWithWarning') {
            # Write log
            "$(Get-Date -Format G): Exported PST to $PSTPath\$FullName.pst" >> .\logs.txt
            return $true
        } elseif ($Request.Status -in 'Failed','AutoSuspended','None','Suspended','Synced') {
            # When the export job status is failed or pending
            "$(Get-Date -Format G): $Name Export PST failed" >> .\logs.txt
            Send-Email
            return $false
        } else {
            continue
        }
    } while ($true)
}

# Establish connection if there are users to depart today
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
            # 1. Check ExportPST. Export if Y and wait until it's completed.
            if ($Line.ExportPST -eq 'Y') {
                $ExportFlag = Export-PST -Name $Line.Alias -FullName $Line.Name
            }

            # 2. Once Step 1 is completed, check FwdEmail. Remove UM and mailbox as needed.
            if ($Line.FwdEmail -eq 'Y') {
                # Either Export is complete or there is no export required, mailbox is safe to delete.
                if (($ExportFlag -eq $true) -or ($Line.ExportPST -eq 'N')) {
                    try {
                        Disable-UMMailbox -Identity $Line.Alias -Confirm:$false -ErrorAction Stop
                        "$(Get-Date -Format G): Disabled UM mailbox $($Line.Alias)" >> .\logs.txt
                    }
                    catch {
                        "$_" >> .\logs.txt
                        "$(Get-Date -Format G): Failed to disable UM mailbox $($Line.Alias)" >> .\logs.txt
                        Send-Email
                    }
                    try {
                        Disable-Mailbox -Identity $Line.Alias -Confirm:$false -ErrorAction Stop
                        "$(Get-Date -Format G): Disabled mailbox $($Line.Alias)" >> .\logs.txt
                    }
                    catch {
                        "$_" >> .\logs.txt
                        "$(Get-Date -Format G): Failed to disable mailbox $($Line.Alias)" >> .\logs.txt
                        Send-Email
                    }
                } else {
                    "$(Get-Date -Format G): Export PST is not completed $($Line.Alias)"
                    Send-Email
                }
            }

            # 3. Check FwdLync to remove Lync user
            if ($Line.FwdLync -eq 'Y') {
                try {
                    $UPN = (Get-ADUser -Identity $Line.Alias).UserPrincipalName
                    Disable-CsUser -Identity $UPN -Confirm:$false -ErrorAction Stop
                    "$(Get-Date -Format G): Removed Lync user $($Line.Alias)" >> .\logs.txt
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