#parameters

param([string]$csv_source="c:\is531\public\GroupProject3Usernames.csv")

# accepts a csv parameter
if ($csv_source) {
    $source_exists = Test-Path $csv_source -PathType Leaf

    if($source_exists){
        $names = Import-Csv $csv_source
    } else {
        Write-Host "ERROR: Source ($csv_source) does not exist or is a directory" -ForegroundColor Red
        break
    }
} else {
    Write-Host "ERROR: Source is not defined" -ForegroundColor Red
    break
}
#loop for each row
#retrieve params from imported csv

#$csv_source
#$names

#Function Create-Users([string]$names){
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
#}   

 