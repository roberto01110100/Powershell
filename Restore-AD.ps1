# Ensure the Active Directory module is loaded
Import-Module ActiveDirectory

# Set up the necessary variables
$ouName = "Finance"
$domain = "dc=dname,dc=rdname,dc=com"
$ouPath = "ou=$ouName,$domain"

# Use $PSScriptRoot to get the directory where the script is located and construct the CSV file path
$csvFilePath = Join-Path $PSScriptRoot "\financePersonnel.csv"

# Step 1: Check if the Finance Organizational Unit (OU) already exists
Write-Host "Checking if the '$ouName' Organizational Unit (OU) exists..."
$existingOU = Get-ADOrganizationalUnit -Filter "Name -eq '$ouName'" -SearchBase $domain -ErrorAction SilentlyContinue

if ($existingOU) {
    Write-Host "The OU '$ouName' exists." -ForegroundColor Yellow 
    
    # Remove protection against accidental deletion
    try {
        Set-ADOrganizationalUnit -Identity $ouPath -ProtectedFromAccidentalDeletion $false
        
        #Check the Object Items and delete them if necessary
        $OUObjects = Get-ADObject -Filter {objectclass -eq "user"} -SearchBase $ouPath
        if ($OUObjects) {
            foreach ($object in $OUObjects)  {
                Remove-ADObject -Identity $object.DistinguishedName -Confirm:$false
            }
        }
        # to remove the OU
        Get-ADOrganizationalUnit -Filter {Name -eq "Finance"} | Remove-ADOrganizationalUnit -Confirm:$false
        Write-Host "The '$ouName' OU has been successfully removed."

    } catch {
        Write-Host "Error: Unable to remove the '$ouName' OU or disable accidental deletion protection: $_"
    }
} else {
    Write-Host "The '$ouName' Organizational Unit does not exist." -ForegroundColor Yellow
}


# Step 2: Create the Finance Organizational Unit (OU)
try {
    New-ADOrganizationalUnit -Name $ouName -Path $domain -ErrorAction Stop
    Write-Host "The '$ouName' OU has been created." -ForegroundColor Yellow
} catch {
    Write-Host "Error creating the '$ouName' OU: $_"
}

$financeOUPath = "OU=Finance,DC=dname,DC=rdname,DC=com"
# Step 3: Import user data from the CSV and add users to the Finance OU
Write-Host "Importing user data from $csvFilePath..."
if (Test-Path $csvFilePath) {
    $users = Import-Csv -Path $csvFilePath
    foreach ($user in $users) {
        # Build the display name
        $displayName = "$($user.First_Name) $($user.Last_Name)"
        $fullName = $displayName  # Set the Name field to the full name

        # Secure Password
        $securePassword = ConvertTo-SecureString "DefaultPassword123!" -AsPlainText -Force

        # Add the user to the Finance OU
        try {
            Write-Host "Creating user: $displayName"

            # Create the user account in Active Directory
            New-ADUser -SamAccountName $user.samAccount `
                       -UserPrincipalName "$($user.First_Name).$($user.Last_Name)@consultingfirm.com" `
                       -GivenName $user.First_Name `
                       -Surname $user.Last_Name `
                       -DisplayName $displayName `
                       -Name $fullName `
                       -PostalCode $user.PostalCode `
                       -MobilePhone $user.MobilePhone `
                       -AccountPassword $securePassword `
                       -Enabled $true `
                       -Path $financeOUPath `
                       -OfficePhone $user.OfficePhone `  # Set OfficePhone to the correct field
                      

            Write-Host "User '$displayName' was successfully created."
        } catch {
            Write-Host "Failed to create user '$displayName': $_"
        }
    }
} else {
    Write-Host "Error: Could not find the CSV file at $csvFilePath. Please check the file path."
}



# Step 4: Export the list of users in the Finance OU to a text file
Write-Host "Exporting user details to AdResults.txt..."
Get-ADUser -Filter * -SearchBase $ouPath -Properties DisplayName, PostalCode, OfficePhone, MobilePhone |
    Select-Object DisplayName, PostalCode, OfficePhone, MobilePhone |
    Out-File -FilePath "$PSScriptRoot\AdResults.txt"

Write-Host "User details have been successfully exported to 'AdResults.txt'."
