function prompt {
    Write-Host "$(Get-Date -Format 'HH:mm:ss')" -NoNewLine -ForegroundColor Gray
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

function Sync-Config {
    param (
        [ValidateSet('Vim', 'PS')][Parameter(Mandatory=$true)]$VimOrPS,
        [ValidateSet('Push', 'Pull')][Parameter(Mandatory=$true)]$PushOrPull
    )
    if ($VimOrPS -eq 'Vim') {
        $RepoPath = 'E:\Repos\VimConfigs\gvim'
        $FileName = '_vimrc'
        $ConfigPath = 'C:\Users\jaywa'
    }
    else {
        $RepoPath = 'E:\Repos\PowerShell\profile'
        $FileName = 'Microsoft.PowerShell_profile.ps1'
        $ConfigPath = 'E:\Documents\PowerShell'
    }
    Push-Location
    Set-Location $RepoPath
    if ($PushOrPull -eq 'Push') {
        Write-Host "Copying $Src to $RepoPath"
        Copy-Item "$ConfigPath\$FileName" $RepoPath -Force -Confirm:$false
        git pull
        git add .
        git commit -m 'update config'
        git push
    }
    else {
        git pull
        Copy-Item "$RepoPath\$FileName" $ConfigPath -Force -Confirm:$false
    }
    Pop-Location
    Write-Host 'Done' -ForegroundColor Green
}

New-Alias -Name 'vim' -Value 'C:\Program Files (x86)\Vim\vim82\gvim.exe'
New-Alias -Name 'vi' -Value 'C:\Program Files (x86)\Vim\vim82\gvim.exe'
New-Alias -Name 'll' -Value 'Get-ChildItem'
New-Alias -Name 'grep' -Value 'Select-String'
chcp 65001
cd 'E:\Repos'
