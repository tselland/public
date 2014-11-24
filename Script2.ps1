## IS 531 Group Project 3
## Powershell Administration Script
## Group 2B
## Carlos Filoteo, Nathan Dudley, Michael Helvey, Travis Selland
## Script 2 - Gather information about Windows computers in your enterprise

## Parameters 
# $computer_name parameter will accept a fully qualified domain name (FQDN), a NetBIOS name, or an IP address.
# if no computer name is passed, $computer_name will default to localhost
param([string]$computer_name="localhost", [switch]$no_report)

#Function to append stats and info to the stats for each WMI Object
Function Set-Stats([string]$label, [string]$info) {
    $stats = "" | Select Label, Info
    $stats.label = $label
    $stats.info = $info

    return $stats
}

##could be removed
$computer = Get-WmiObject -Class Win32_Desktop -ComputerName $computer_name

#elements of the system are gathered and assigned to variables
$computerSystem = Get-WmiObject Win32_ComputerSystem -ComputerName $computer_name
$computerBIOS = Get-CimInstance CIM_BIOSElement -ComputerName $computer_name
$computerOS = Get-WmiObject Win32_OperatingSystem -ComputerName $computer_name
$computerCPU = Get-WmiObject Win32_Processor -ComputerName $computer_name
$computerDrives = Get-WmiObject Win32_CDROMDrive -ComputerName $computer_name
$computerHDD = Get-WMIObject Win32_LogicalDisk -ComputerName $computer_name -Filter "DeviceID = 'C:'"
$computerBattery = Get-WmiObject Win32_Battery -ComputerName $computer_name 
$lastBootUpTime = Get-CimInstance -ComputerName $computer_name -ClassName win32_operatingsystem | select lastbootuptime
$userSession = Get-WmiObject Win32_Session -ComputerName $computer_name
Clear-Host

#Create table object that will be populated with system information
$table = @()

#Gather basic information pertaining to Hard Drive, RAM, and Session time
$hddCap = "{0:N2}" -f ($computerHDD.Size/1GB) + "GB"
$hddSpace = "{0:P2}" -f ($computerHDD.FreeSpace/$computerHDD.Size) + " Free (" + "{0:N2}" -f ($computerHDD.FreeSpace/1GB) + "GB)"
$ram = "{0:N2}" -f ($computerSystem.TotalPhysicalMemory/1GB) + "GB"
$startTime = [System.Management.ManagementDateTimeConverter]::ToDateTime($userSession[0].StartTime)

#if there are drives associated with this machine, they will be listed here.
if($computerDrives){
    $drive = $computerDrives.Description
} else {
    $drive = "None"
}

#Output System information Header
Write-Host "System Information for: " $computerSystem.Name -BackgroundColor DarkCyan

#SPECIFICATIONS
$table += Set-Stats "Manufacturer" $computerSystem.Manufacturer
$table += Set-Stats "Model" $computerSystem.Model
$table += Set-Stats "Serial Number" $computerBIOS.SerialNumber

#HARDWARE
$table += Set-Stats "CPU" $computerCPU.Name
$table += Set-Stats "Processors" $computerSystem.NumberOfProcessors
$table += Set-Stats "HDD Capacity" $hddCap
$table += Set-Stats "HDD Space" $hddSpace
$table += Set-Stats "RAM" $ram
$table += Set-Stats "Optical Drive" $drive

#OPERATING SYSTEM
$table += Set-Stats "Operating System" $computerOS.caption
$table += Set-Stats "Service Pack" $computerOS.ServicePackMajorVersion

#USER INFORMATION
$table += Set-Stats "Current User" $computerSystem.UserName
$table += Set-Stats "Session Start" $startTime
$table += Set-Stats "Last Reboot" $lastBootUpTime.lastbootuptime

#BATTERY (if it exists)
if($computerBattery) {
    $percentRemaining = $computerBattery.EstimatedChargeRemaining
    $hoursRemaining = [Math]::Floor([decimal]($computerBattery.EstimatedRunTime / 60))
    $minutesRemaining = $computerBattery.EstimatedRunTime % 60
    $timeRemaining = "$percentRemaining% ($hoursRemaining hours $minutesRemaining minutes remaining)"
    
    $pluggedIn = gwmi -Class batterystatus -Namespace root\wmi
    $pluggedIn = $pluggedIn.PowerOnline
    if($pluggedIn){
        $table += Set-Stats "Battery Status" "Plugged-In, Charging"
    } else {
        $table += Set-Stats "Battery Status" "Unplugged"
    }
    
    $table+= Set-Stats "Battery Remaining" $timeRemaining
} else {
    $table += Set-Stats "Battery Remaining" "No battery connected"
}

#Output results table to console
$table

#Export Report to CSV and format for excel
if ($no_report -eq $false) {
    #Export report as CSV
    $fullName = "$PSScriptRoot\$($computerSystem.Name).csv"
    $table | Export-CSV -path $fullName -ErrorAction 'silentlycontinue'
    
    #Format report in Excel

    #Create new workbook from CSV
    $XL = New-Object -comobject Excel.Application
    $XL.Visible = $true   
    $wb = $XL.Workbooks.Open($fullName)
    $ws = $wb.Worksheets.Item(1)

    #Get Rid of the Automatic Header
    $ws.Cells.Item(1,"a").EntireRow.Delete() | Out-Null
    $ws.Cells.Item(1,"a").EntireRow.Delete() | Out-Null

    #Add Sub Headings
    #These headings are hard coded based on a fixed format
    $headingRows = @(1, 6, 14, 18, 23)
    $headingCaptions = @("Computer Specs", "Hardware", "Operating System", "User and Session", "Battery Life")
    $counter = 0

    #for each heading row, change color and formatting
    foreach($hr in $headingRows){
        $ws.Range("a$hr").EntireRow.Insert(-4142) | Out-Null #space out the different sections
        $ws.Range("a$hr").EntireRow.Insert(-4142) | Out-Null
        $hr+=1
        $range = "a$hr" + ":b$hr"
        $ws.Range($range).Merge() | Out-Null #merge the heading cells
        $ws.Cells.Item($hr, "a").Interior.ColorIndex = 44 #Orange interior
        $ws.Cells.Item($hr,"a").HorizontalAlignment = -4108 #Centered
        $ws.Cells.Item($hr, "a") = $headingCaptions[$counter] #Add the caption
        $counter++
    }
    
    #Format each table with borders and 
    foreach($hr in $headingRows){
        $hr+=2 #We're offset by two because of the rows added in the previous loop
        $dataRange = $ws.Cells.Item($hr, "a").CurrentRegion
        $dataRange.Borders.LineStyle = 1 #Straight line
        $dataRange.Borders.Weight = 2 #Typical weight, normal - not thick
        $dataRange.HorizontalAlignment = -4131 #Left
        $hr-=1
        $ws.Cells.Item($hr,"a").HorizontalAlignment = -4108 #Centered
    }

    #Add Primary Heading
    $ws.Range("a1").EntireRow.Insert(-4142) | Out-Null
    $ws.Cells.Item(1,"a") = "Summary Data for $($computerSystem.Name)"
    $ws.Cells.Item(1,1).Font.Size = 14
    $ws.Cells.Item(1,1).Font.Bold = $True
    $ws.Cells.Item(1,"a").Interior.ColorIndex = 6 #Yellow
    $ws.Range("a1:b1").Merge() | Out-Null
    $ws.Columns.Autofit() | Out-Null
    $ws.Cells.Item(1,1).HorizontalAlignment = -4108 #Center

    #Create another small table with raw data and header for the charts
    $ws.Cells.Item(3, "f") = "Free"
    $ws.Cells.Item(4, "f") = $computerHDD.FreeSpace/1GB
    $ws.Cells.Item(3, "g") = "Used"
    $ws.Cells.Item(4, "g") = $computerHDD.Size/1GB - $computerHDD.FreeSpace/1GB
    
    #Create an empty chart
    $chart = $ws.Shapes.AddChart().Chart
    $chart.ChartType = 5 #pie chart
    
    #Move the chart to the right of the data
    $ws.Shapes.Item("Chart 1").Top = 30
    $ws.Shapes.Item("Chart 1").Left = 350

    #Set the source to the small table we created
    $chart.SetSourceData($ws.Range("f3:g4"))
    $chart.HasTitle = $True
    $chart.ChartTitle.Text = "Hard Disk Allocation"
    $chart.ApplyLayout(6,69) #this layout shows numbers on the chart

    #Save the document as a .xlsx in the same folder as the current script
    $newFullName = $fullName.replace('.csv', '.xlsx')
    $wb.SaveAs($newFullName)
}
