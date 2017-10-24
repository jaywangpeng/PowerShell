# This script is to delete the cache folder on all YSOFT servers with
# the restart of all the Ysoft services.
# Written by Jay Wang
# Release Date: Oct. 6th, 2017
# Version 1.0

$Servers = 'ADL-PRNT-V04','BNE-PRNT-V03','CBR-PRNT-V06','MEL-PRNT-V05',`
'NZD-PRNT-V08','PER-PRNT-V07','SYD-PRNT-V02','ECP-PRNT-V01'

# wildcard all YSoft services
$ServiceName = 'YSoft*'

$FlagDone = $false

function Clear-Cache($Server,$Service)
{
    # Stop the services
    try
    {
        Write-Host "Stopping YSoft services on $Server"
        $Services = Get-Service -ComputerName $Server -ServiceName $Service
        $Services | Stop-Service -ErrorAction Stop
        $FlagDone = $true
    }
    catch
    {
        Write-Host "Error stopping services. Script exited." `
            -ForegroundColor Red
        exit
    }

    # Delete cache folder except ECP-PRNT-V01 console server
    if($Server -ne 'ECP-PRNT-V01')
    {
        Write-Host "Deleting cache folder on $Server"
        Get-ChildItem "\\$Server\c$\SafeQORS\server\cache" -Recurse | Remove-Item `
            -Recurse -Force
    }
    else
    {
        Write-Host "Skipped deleting cache folder on $Server"
    }

    # Start the services
    try
    {
        Write-Host "Starting YSoft services on $Server"
        $Services | Start-Service -ErrorAction Stop
        $FlagDone = $true
    }
    catch
    {
        Write-Host "Error starting services. Script exited." `
            -ForegroundColor Red
        exit
    }
}

$Menu = "Which environment to restart the YSoft services?
    ADL
    BNE
    CBR
    MEL
    NZD
    PER
    SYD
    Console"
Write-Host "`n$Menu"
$Input = Read-Host "Please enter one of the above(case-insensitive)"
switch($Input)
{
    'ADL' { $ServerName = $Servers[0] }
    'BNE' { $ServerName = $Servers[1] }
    'CBR' { $ServerName = $Servers[2] }
    'MEL' { $ServerName = $Servers[3] }
    'NZD' { $ServerName = $Servers[4] }
    'PER' { $ServerName = $Servers[5] }
    'SYD' { $ServerName = $Servers[6] }
    'Console' { $ServerName = $Servers[7] }
    default { Write-Host 'Wrong input. Please rerun the script.'; exit }
}

# Execute the function with the select param
Clear-Cache -Server $ServerName -Service $ServiceName

# Output the result
if($FlagDone -eq $true)
{
    Write-Host -ForegroundColor Green `
    "$ServerName Cache folders are cleaned. YSoft services are restarted."
}
else
{
    Write-Host -ForegroundColor Red `
    "$ServerName Errors. Please check the services on the server"
}
