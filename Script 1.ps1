#parameters

param([string]$csv_source="c:\is531\public\GroupProject3Usernames.csv", [switch]$no_report, [string]$output_filename="output.csv" )

# accepts a csv parameter
if ($csv_source) {
    # Check is source exists
    $source_exists = Test-Path $csv_source -PathType Leaf
    
    if($source_exists){
        # Check that source is actually a .csv file
        if (gci $csv_source *.csv){
            Write-Host "Importing usernames..." -ForegroundColor Cyan
            $names = Import-Csv $csv_source
        } else {
            Write-Host "ERROR: File is not a .csv file" -ForegroundColor Red
            break
        }
    } else {
        # Display error
        Write-Host "ERROR: Source ($csv_source) does not exist or is a directory" -ForegroundColor Red
        break
    }
} else {
    # Should never reach this code
    Write-Host "ERROR: Source is not defined" -ForegroundColor Red
    break
}
#loop for each row
#retrieve params from imported csv

#$csv_source
#$names

Function Create-Users($table){
    Foreach($line in $table){
        $username = $line.username
        $surname = $line.lastname
        $givenname = $line.firstname

        New-ADUser $username
        $user = Get-ADUser $username
        $user.Surname = $surname
        $user.Givenname = $givenname

        Set-ADUser -instance $user
    }
}   

Create-Users $names
