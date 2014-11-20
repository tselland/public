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

$computerSystem = Get-WmiObject Win32_ComputerSystem
$computerBIOS = Get-CimInstance CIM_BIOSElement
$computerOS = Get-WmiObject Win32_OperatingSystem
$computerCPU = Get-WmiObject Win32_Processor
$computerHDD = Get-WMIObject Win32_LogicalDisk -Filter "DeviceID = 'C:'"
#Clear-Host

Write-Host "System Information for: " $computerSystem.Name -BackgroundColor DarkCyan
"Manufacturer: " + $computerSystem.Manufacturer
"Model: " + $computerSystem.Model
"Serial Number: " + $computerBIOS.SerialNumber
"CPU: " + $computerCPU.Name
"HDD Capacity: "  + "{0:N2}" -f ($computerHDD.Size/1GB) + "GB"
"HDD Space: " + "{0:P2}" -f ($computerHDD.FreeSpace/$computerHDD.Size) + " Free (" + "{0:N2}" -f ($computerHDD.FreeSpace/1GB) + "GB)"
"RAM: " + "{0:N2}" -f ($computerSystem.TotalPhysicalMemory/1GB) + "GB"
"Operating System: " + $computerOS.caption + ", Service Pack: " + $computerOS.ServicePackMajorVersion
"User logged In: " + $computerSystem.UserName
"Last Reboot: " + $computerOS.LastBootUpTime

