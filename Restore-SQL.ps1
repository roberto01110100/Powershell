$ServerInstance = "Client1"
$DatabaseName = "TestDB"
$TableName = "Client_A_Contacts"
$CsvPath = Join-Path $PSScriptroot "NewClientData.csv"

# 1. Check for the existence of the database named ClientDB using error handling
try {
    $dbExists = Invoke-Sqlcmd -ServerInstance $ServerInstance -Query "SELECT name FROM sys.databases WHERE name = '$DatabaseName'" -ErrorAction Stop
    
    if ($dbExists) {
        Write-Host "The database $DatabaseName already exists. Deleting it..."
        #Set the database to single user mode to ensure there are no connections interfering with deletion
        Invoke-Sqlcmd -ServerInstance $ServerInstance -Query "ALTER DATABASE $DatabaseName SET SINGLE_USER WITH ROLLBACK IMMEDIATE"


        Invoke-Sqlcmd -ServerInstance $ServerInstance -Query "DROP DATABASE $DatabaseName"
        Write-Host "The database $DatabaseName was deleted." -ForegroundColor Green
    } else {
        Write-Host "The database $DatabaseName does not exist."
    }

} catch {
    #catch any error information for debugging
    Write-Host "An error occured when checking for $DatabaseName" -ForegroundColor Red
    write-host "error details: $_" -ForegroundColor Red
}

# 2. Create a new database named ClientDB
Write-Host "Creating a new database named $DatabaseName..."
Invoke-Sqlcmd -ServerInstance $ServerInstance -Query "CREATE DATABASE $DatabaseName"
Write-Host "The database $DatabaseName was created." -ForegroundColor Green

# 3. Create a table named Client_A_Contacts in the database
$CreateTableQuery = "
USE $DatabaseName;
CREATE TABLE $TableName (
    ClientID INT IDENTITY(1,1) PRIMARY KEY,
    First_Name NVARCHAR(100),
    Last_Name NVARCHAR(100),
    City NVARCHAR(70),
    County NVARCHAR(70),
    Zip_Code NVARCHAR(15),
    Office_Phone VARCHAR(15),
    Mobile_Phone VARCHAR(15)

);
"
Write-Host "Creating the table $TableName in database $DatabaseName..."
Invoke-Sqlcmd -ServerInstance $ServerInstance -Query $CreateTableQuery
Write-Host "The table $TableName was created." -ForegroundColor Green

# 4. Import the data from NewClientData.csv and write the data to the table Client_A_Contacts
Write-Host "Importing data from $CsvPath into table $TableName..."
$CsvData = Import-Csv -Path $CsvPath
foreach ($row in $CsvData) {
    $InsertQuery = "
    USE $DatabaseName;
    INSERT INTO $TableName (First_Name, Last_Name, City, County, Zip_Code, Office_Phone, Mobile_Phone)
    VALUES ('$($row.first_name)', '$($row.last_name)', '$($row.city)', '$($row.county)', '$($row.zip)', '$($row.officePhone)','$($row.mobilePhone)');
    "
    Invoke-Sqlcmd -ServerInstance $ServerInstance -Query $InsertQuery
}
Write-Host "Data from $CsvPath has been successfully inserted into table $TableName." -ForegroundColor Green

# 5. Generate the output file SqlResults.txt
Write-Host "Generating output file SqlResults.txt..."
Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $DatabaseName -Query 'SELECT * FROM dbo.Client_A_Contacts' > .\SqlResults.txt
Write-Host "The output file SqlResults.txt has been generated." -ForegroundColor Green