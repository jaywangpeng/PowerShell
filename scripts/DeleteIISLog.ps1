# This script is to delete IIS log files older than 30 days
# on the servers in servers.txt

$Servers = Get-Content 'C:\ScheduledTasks\DeleteIISLog\Servers.txt'

foreach ($Server in $Servers)
{
    $LogFiles = Get-ChildItem "\\$Server\C$\inetpub\logs\LogFiles\*.log" `
        -Recurse -ErrorAction Stop
    foreach ($LogFile in $LogFiles)
    {
        if([datetime]::ParseExact(
            $LogFile.BaseName.TrimStart('u_ex', 'u_in'), 'yyMMd',
            [System.Globalization.CultureInfo]::CurrentCulture) `
            -lt (Get-Date).AddDays(-31))
        {
            Remove-Item $LogFile
        }
    }
}