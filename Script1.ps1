#parameters
param([string]$csv_source="c:\is531\public\GroupProject3Usernames.csv", [switch]$no_report, [string]$output_filename="output.csv", [switch]$delete_users=$false )

# Check for a csv file in parameter
if ($csv_source) {
    # Check is source exists
    $source_exists = Test-Path $csv_source -PathType Leaf
    
    if($source_exists){
        # Check that source is actually a .csv file
        if (gci $csv_source *.csv){
            Write-Host "Importing usernames..." -ForegroundColor Cyan
            $names = Import-Csv $csv_source
            $sourcePath = $(gci $csv_source).DirectoryName
        } else {
            # Output errors
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

# Function to create users based on entries in .csv file.
Function Create-Users($table){
    Write-Host "`nCREATING USERS...`n"
    #loop for each row
    Foreach($line in $table){
    #retrieve params from imported csv
        $username = $line.username
        $surname = $line.lastname
        $givenname = $line.firstname
        
        try {
            # Check if user already exists
            $existing_user = Get-ADUser $username
            Write-Host "User $username already exists" -ForegroundColor Red
        }
        catch {
            # Create new user 
            New-ADUser $username
            $user = Get-ADUser $username
            $user.Surname = $surname
            $user.Givenname = $givenname
            
            # Set ADUser properties, notify user of success
            Set-ADUser -instance $user
            Write-Host "User $username created successfully!" -ForegroundColor DarkCyan
            
            # Add creation result to table 
            $line | Add-Member -name "SamAccountName" -value $user.SamAccountName -MemberType NoteProperty
            $line | Add-Member -name "SID" -value $user.SID -MemberType NoteProperty
        }
    }
    
    # Export resulting table results to CSV file
    if ($no_report -eq $false){
        Write-Host "`nExporting results as $output_filename at directory $sourcePath `n" -ForegroundColor Magenta
        $table | Export-Csv "$sourcePath\$output_filename"
    }
    else{
        Write-Host "`nReport not created`n" -ForegroundColor Magenta
    }
}

# Function to delete users based on entries in .csv file
Function Delete-Users($table){
    Write-Host "`nDeleting USERS...`n"
    Foreach($line in $table){
        
        try {
            # Check if user is in Active Directory
            $existing_user = Get-ADUser $line.username
            if ($existing_user){
                # Delete user 
                Write-Host "Deleting $existing_user" -ForegroundColor DarkCyan
                Remove-ADuser $existing_user -Confirm:$false
            }
        }
        catch {
            # Output error if user is not in Active Directory
            Write-Host "User $line.username does not exist!" -ForegroundColor Red
        }
    }
} 

# Create or delete users based on user input of switch
if ($delete_users -eq $false) {
    Create-Users $names
} else {
    Delete-Users $names
}
