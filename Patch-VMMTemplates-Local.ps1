# Parameters for robocopy
$MaxAge = 5
$WSUSPath = '\\ecp-sccm-v03\sup$\ServerStage3VMM'
$ADLPatchFolder = '\\ADL-HYPV-P15\c$\VMTemplatePatch'
$AKLPatchFolder = '\\AKL-HYPV-P29\c$\VMTemplatePatch'
$BNEPatchFolder = '\\BNE-HYPV-P16\c$\VMTemplatePatch'
$CBRPatchFolder = '\\CBR-HYPV-P17\c$\VMTemplatePatch'
$CHCPatchFolder = '\\CHC-HYPV-P30\c$\VMTemplatePatch'
$ECPPatchFolder = '\\ECP-HYPV-P14\c$\VMTemplatePatch'
$MELPatchFolder = '\\MEL-HYPV-P18\c$\VMTemplatePatch'
$NZDPatchFolder = '\\NZD-HYPV-P31\c$\VMTemplatePatch'
$PERPatchFolder = '\\PER-HYPV-P19\c$\VMTemplatePatch'
$SYDPatchFolder = '\\SYD-HYPV-P25\c$\VMTemplatePatch'

# VM template locations
$ADLVHDX = 'C:\VMMLibrary\VM\ADL-VM-W2k12R2\ADL-VM-W2k12R2 - C Drive.vhdx'
$AKLVHDX = 'C:\AKL-VMMLibrary\VM\Windows 2012 R2\AKL-VM-W2k12R2 - C Drive.vhdx'
$BNEVHDX = 'C:\VMMLibrary\VM\BNE-VM-W2k12R2\BNE-VM-W2k12R2 - C Drive.vhdx'
$CBRVHDX = 'C:\VMMLibrary\VM\CBR-VM-W2k12R2\CBR-VM-W2k12R2 - C Drive.vhdx'
$CHCVHDX = 'C:\CHC-VMMLibrary\VM\Windows 2012 R2\CHC-VM-W2k12R2 - C Drive.vhdx'
$ECPVHDX = 'D:\VMMLibrary\VM\Windows 2012 R2\ECP-VM-W2k12R2 - C Drive.vhdx'
$MELVHDX = 'C:\VMMLibrary\VM\MEL-VM-W2k12R2\MEL-VM-W2k12R2 - C Drive.vhdx'
$NZDVHDX = 'C:\NZD-VMMLibrary\VM\Windows 2012 R2\NZD-VM-W2k12R2 - C Drive.vhdx'
$PERVHDX = 'C:\VMMLibrary\VM\PER-VM-W2k12R2\PER-VM-W2k12R2 - C Drive.vhdx'
$SYDVHDX = 'D:\VMMLibrary\VM\Windows 2012 R2\SYD-VM-W2k12R2 - C Drive.vhdx'

# Parameters for patch installation
$FileTypes = '*.cab','*.msu'

############### Script starts from here ###############
Write-Host "For which site are you patching VM template?
    ADL
    AKL
    BNE
    CBR
    CHC
    ECP
    MEL
    NZD
    PER
    SYD" -ForegroundColor Cyan
try
{
    [string]$InputSite = Read-Host "Type one of above"
    switch($InputSite)
    {
        'ADL' { $PatchFolder = $ADLPatchFolder; $VHDXPath = $ADLVHDX }
        'AKL' { $PatchFolder = $AKLPatchFolder; $VHDXPath = $AKLVHDX }
        'BNE' { $PatchFolder = $BNEPatchFolder; $VHDXPath = $BNEVHDX }
        'CBR' { $PatchFolder = $CBRPatchFolder; $VHDXPath = $CBRVHDX }
        'CHC' { $PatchFolder = $CHCPatchFolder; $VHDXPath = $CHCVHDX }
        'ECP' { $PatchFolder = $ECPPatchFolder; $VHDXPath = $ECPVHDX }
        'MEL' { $PatchFolder = $MELPatchFolder; $VHDXPath = $MELVHDX }
        'NZD' { $PatchFolder = $NZDPatchFolder; $VHDXPath = $NZDVHDX }
        'PER' { $PatchFolder = $PERPatchFolder; $VHDXPath = $PERVHDX }
        'SYD' { $PatchFolder = $SYDPatchFolder; $VHDXPath = $SYDVHDX }
        default {
            Write-Host "Entered the wrong site. Please rerun the script."
            exit
        }
    }
}
catch
{
    Write-Host "Entered the wrong site. Please rerun the script."
    exit
}

# Function to dismount disk
function Dismount-VHDX ($DiskPath)
{
    Write-Host "Dismounting disk $DiskPath" -ForegroundColor Cyan
    Dismount-DiskImage $DiskPath
}

# Empty the folder before copy
Write-Host "You are patching the VM template:" -ForegroundColor Cyan
Write-Host " "
Write-Host "$VHDXPath"  -ForegroundColor Cyan
Write-Host " "
Write-Host "Deleting all updates in Patch folder: $PatchFolder" -ForegroundColor Cyan
Get-ChildItem -Path $PatchFolder -Recurse -Force |
Remove-Item -Recurse -Force

Start-Sleep -s 2

# Copy the updates from last xx days to patch folders
robocopy.exe "$WSUSPath" "$PatchFolder" /s /maxage:"$MaxAge"

# Check the drive letters before mounting template VM
$DrivesBefore = Get-PSDrive -PSProvider 'FileSystem'

Write-Host "Mounting disk $VHDXPath" -ForegroundColor Cyan
Mount-DiskImage -ImagePath $VHDXPath -Access ReadWrite

Start-Sleep -s 3

# Get the drive letters after mounting template VM
$DrivesAfter = Get-PSDrive -PSProvider 'FileSystem'

# Get the difference of drive letters and assign to $Drive
foreach ($TempDrive in $DrivesAfter)
{
    if ($TempDrive -notin $DrivesBefore)
    {
        $Drive = $TempDrive
        Write-Host "$($Drive.Name) drive has been mounted." -ForegroundColor Green
    }
}

# Get all the update file paths with names as strings
$Updates = Get-ChildItem -Path $PatchFolder -Include $FileTypes -Recurse -File |
    Select-Object Fullname

# Check if the folder is empty or not
If( !($Updates) )
{
    Write-Host "Folder is empty" -ForegroundColor Red
    Dismount-VHDX -DiskPath $VHDXPath
    exit
}

# Install updates to VM templates
try
{
    Write-Host ""
    Write-Host "Installing updates" -ForegroundColor Cyan

    foreach ($Update in $Updates)
    {
        dism.exe /image:"$($Drive.Name)":\ /add-package /packagepath:"$($Update.FullName)"
    }
}
catch
{
    Write-Host "$_" -ForegroundColor Red
    Dismount-VHDX -DiskPath $VHDXPath
    exit
}

Dismount-VHDX -DiskPath $VHDXPath
Write-Host "All updates have been installed successfully." -ForegroundColor Green
