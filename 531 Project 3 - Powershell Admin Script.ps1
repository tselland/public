## IS 531 Group Project 3
## Powershell Administration Script
## Group 2B
## Carlos Filoteo, Nathan Dudley, Michael Helvey, Travis Selland

## ----- Script 1 - Create Active Directory User Accounts

#parameters

param([string]$csv_source)

# accepts a csv parameter
if ($csv_source) {
    Import-Csv $csv_source
} else {
    Write-Host "ERROR: Source is not defined" -ForegroundColor Red
    break
}


## ----- Script 2 - Gather information about Windows computers in your enterprise