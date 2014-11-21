## IS 531 Group Project 3
## Powershell Administration Script
## Group 2B
## Carlos Filoteo, Nathan Dudley, Michael Helvey, Travis Selland

## ----- Script 1 - Create Active Directory User Accounts

#parameters

param([string]$csv_source="c:\is531\public\GroupProject3Usernames.csv")

# accepts a csv parameter
if ($csv_source) {
    $source_exists = Test-Path $csv_source 

    if($source_exists){
        $names = Import-Csv $csv_source
    } else {
        Write-Host "ERROR: Source path does not exist" -ForegroundColor Red
        break
    }
} else {
    Write-Host "ERROR: Source is not defined" -ForegroundColor Red
    break
}
#loop for each row
#retrieve params from imported csv

$csv_source
$names


Foreach($line in $names){
    $username = $line.username
    $surname = $line.lastname
    $givenname = $line.firstname

    New-ADUser $username
    $user = Get-ADUser $username
    $user.Surname = $surname
    $user.Givenname = $givenname

    Set-ADUser -instance $user
}




## ----- Script 2 - Gather information about Windows computers in your enterprise

##NEEDSWORK: fix method to get computerName or list of computers.
param([string]$computer = (Get-WmiObject -Class Win32_Desktop -ComputerName))

Function Set-Stats([string]$label, [string]$info) {
    $stats = "" | Select Label, Info
    $stats.label = $label
    $stats.info = $info

    return $stats
}

$computerSystem = Get-WmiObject Win32_ComputerSystem
$computerBIOS = Get-CimInstance CIM_BIOSElement
$computerOS = Get-WmiObject Win32_OperatingSystem
$computerCPU = Get-WmiObject Win32_Processor
$computerHDD = Get-WMIObject Win32_LogicalDisk -Filter "DeviceID = 'C:'"
$computerBattery = Get-WmiObject Win32_Battery
$lastBootUpTime = Get-CimInstance -ClassName win32_operatingsystem | select lastbootuptime
Clear-Host

$table = @()

Write-Host "System Information for: " $computerSystem.Name -BackgroundColor DarkCyan
$table += Set-Stats "Manufacturer" $computerSystem.Manufacturer
$table += Set-Stats "Model" $computerSystem.Model
$table += Set-Stats "Serial Number" $computerBIOS.SerialNumber
$table += Set-Stats "CPU" $computerCPU.Name

$hddCap = "{0:N2}" -f ($computerHDD.Size/1GB) + "GB"
$table += Set-Stats "HDD Capacity" $hddCap

$hddSpace = "{0:P2}" -f ($computerHDD.FreeSpace/$computerHDD.Size) + " Free (" + "{0:N2}" -f ($computerHDD.FreeSpace/1GB) + "GB)"
$table += Set-Stats "HDD Space" $hddSpace

$ram = "{0:N2}" -f ($computerSystem.TotalPhysicalMemory/1GB) + "GB"
$table += Set-Stats "RAM" $ram

$table += Set-Stats "Operating System" $computerOS.caption
$table += Set-Stats "Service Pack" $computerOS.ServicePackMajorVersion
$table += Set-Stats "Current User" $computerSystem.UserName
$table += Set-Stats "Last Reboot" $lastBootUpTime.lastbootuptime

if($computerBattery) {
    $percentRemaining = $computerBattery.EstimatedChargeRemaining
    $hoursRemaining = [Math]::Floor([decimal]($computerBattery.EstimatedRunTime / 60))
    $minutesRemaining = $computerBattery.EstimatedRunTime % 60
    
    $timeRemaining = "$percentRemaining% ($hoursRemaining hours $minutesRemaining minutes remaining)"
    $table+= Set-Stats "Battery Remaining" $timeRemaining

} else {
    $table += Set-Stats "Battery Remaining" "No battery connected"
}

$table