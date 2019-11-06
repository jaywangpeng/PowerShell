Set-Location 'C:\Repos\'

function prompt {
    Write-Host "$(Get-Date -Format 'dd/MM HH:mm:ss')" -NoNewLine -ForegroundColor Gray
    Write-Host "[$(kubectl.exe config current-context)]" -NoNewLine -ForegroundColor Cyan
    if (git status 2>$null) {
        $Git = git rev-parse --abbrev-ref HEAD
        Write-Host "[$Git]" -NoNewLine -ForegroundColor Magenta
    }
    Write-Host "$((Get-Location).Path)" -NoNewLine -ForegroundColor Green
    Write-Host '>' -NoNewLine -ForegroundColor White
    return ' '
}

function Edit-Profile {
    vim $PROFILE.CurrentUserCurrentHost
}

function Edit-ISEProfile {
    vim "$HOME\Documents\WindowsPowerShell\Microsoft.PowerShellISE_profile.ps1"
}

function Edit-Vimrc {
    vim "$HOME\_vimrc"
}

function Edit-Hosts {
    vim 'C:\Windows\System32\drivers\etc\hosts'
}

function ConvertFrom-Base64 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)][string]$string
    )
    $bytes = [System.Convert]::FromBase64String($string)
    [System.Text.Encoding]::UTF8.GetString($bytes)
}

function ConvertTo-Base64 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)][string]$string
    )
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($string)
    [System.Convert]::ToBase64String($bytes)
}

function Get-PemKey {
    param (
        $BucketName,
        $StackName,
        $Path,
        $AWSAccount
    )
    Read-S3Object -BucketName $BucketName `
                  -Key "$AWSAccount/$StackName.pem" `
                  -File "$Path\$StackName.pem"
}

function Connect-EC2 {
    [CmdletBinding()]
    param(
        $InstanceId,
        $StackName,
        $AWSAccount
    )

    if (!($PemFile = Get-ChildItem "$StackName.pem" -ErrorAction 'silentlycontinue')) {
        if (!($PemFile = Get-PemKey -StackName $StackName -AWSAccount $AWSAccount -ErrorAction 'silentlycontinue')) {
            throw "PEM key file is not found"
        }
    }
    $Instance = Get-EC2Instance -InstanceId "i-$InstanceId" -ProfileName $AWSAccount
    if ($Instance) {
        # Get Name, Password, and IP address
        $InstanceName = ($Instance.RunningInstance.Tags | Where-Object { $_.Key -eq 'Name' }).Value
        # $InstanceID is composed of 'i-' and the long ID.
        $InstancePassword = Get-EC2PasswordData -InstanceId "i-$InstanceId" -PemFile $PemFile
        $InstanceIpAddress = $Instance.RunningInstance.PrivateIpAddress
        # RDP using cmdkey
        cmdkey.exe /generic:$InstanceIpAddress /user:administrator /pass:$InstancePassword
        mstsc.exe /v:$InstanceIpAddress
        # Give mstsc some time to log on
        Start-Sleep -Seconds 2
        # Delete the credential after connected
        cmdkey.exe /delete:$InstanceIpAddress
    }
    else {
        throw "The instance is not found"
    }
}

function Connect-Okta {
    [CmdletBinding()]
    param(
        $IDPAccount = 'default',
        $Region = 'ap-southeast-2'
    )
    saml2aws.exe login -a $IDPAccount --force --skip-prompt
    $AWSAccount = if ($IDPAccount -eq 'default') { '<defaultaccountname>' } else { $IDPAccount }
    $ENV:AWSEnv = $Region
    $ENV:VSEnv = $Region.Split('-') -join ''
    $creds = Get-AWSCredential -ProfileName $AWSAccount
    Initialize-AWSDefaultConfiguration -Region $Region `
                                       -AccessKey $creds.GetCredentials().AccessKey `
                                       -SecretKey $creds.GetCredentials().SecretKey `
                                       -SessionToken $creds.GetCredentials().Token
    Initialize-AWSDefaultConfiguration -Region $Region -ProfileName $AWSAccount
    Set-DefaultAWSRegion -Region $Region
    Write-Output "AWS credential default to profile $AWSAccount in $ENV:AWSEnv"
    Set-VSCredential -Region $ENV:VSEnv -ProfileName $AWSAccount
    Write-Output "VS credential default to profile $AWSAccount in $ENV:VSEnv"
}

New-Alias -Name 'vim' -Value 'C:\Program Files (x86)\Vim\vim81\gvim.exe'
New-Alias -Name 'vi' -Value 'C:\Program Files (x86)\Vim\vim81\gvim.exe'
New-Alias -Name 'll' -Value 'Get-ChildItem'
New-Alias -Name 'grep' -Value 'Select-String'
Set-Alias -Name 'kc' -Value 'kubectl.exe'

# Chocolatey profile
$ChocolateyProfile = "$env:ChocolateyInstall\helpers\chocolateyProfile.psm1"
if (Test-Path($ChocolateyProfile)) {
  Import-Module "$ChocolateyProfile"
}
