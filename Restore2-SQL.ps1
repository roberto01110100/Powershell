# Import SQLServer module
if (-not (Get-packageprovider -Name NuGet)) {
Install-packageProvider -Name NuGet -Force -Confirm:$false
}

if (-not (Get-Module -ListAvailable -Name SqlServer)) {
Write-Host "SqlServer module not found. Installing................" -ForegroundColor Yellow
Install-Module -Name SqlServer -Scope CurrentUser -Force -Allowclobber -Confirm:$false
}

Import-Module SqlServer

# Define parameters
$ServerInstance = "Client1"
$DatabaseName = "TestDB"
$TableName = "Client_A_Contacts"
$CsvPath = Join-Path $PSScriptroot "NewClientData.csv"

# Define the SQL Server connection string
$ConnectionString = "Server=$ServerInstance;Integrated Security=True;"

# 1. Check for the existence of the database using .NET objects
try {
    $sqlConnection = New-Object System.Data.SqlClient.SqlConnection($ConnectionString)
    $sqlConnection.Open()

    $sqlCommand = $sqlConnection.CreateCommand()
    $sqlCommand.CommandText = "SELECT name FROM sys.databases WHERE name = '$DatabaseName'"

    $result = $sqlCommand.ExecuteScalar()

    if ($result) {
        Write-Host "The database $DatabaseName already exists. Deleting it..."
        # Set the database to single-user mode to ensure there are no connections interfering with deletion
        $sqlCommand.CommandText = "ALTER DATABASE $DatabaseName SET SINGLE_USER WITH ROLLBACK IMMEDIATE"
        $sqlCommand.ExecuteNonQuery()

        # Drop the database
        $sqlCommand.CommandText = "DROP DATABASE $DatabaseName"
        $sqlCommand.ExecuteNonQuery()
        Write-Host "The database $DatabaseName was deleted." -ForegroundColor Green
    } else {
        Write-Host "The database $DatabaseName does not exist."
    }

    # Close the connection
    $sqlConnection.Close()

} catch {
    Write-Host "An error occurred when checking for $DatabaseName" -ForegroundColor Red
    Write-Host "Error details: $_" -ForegroundColor Red
}

# 2. Create a new database using .NET objects
Write-Host "Creating a new database named $DatabaseName..."
try {
    $sqlConnection.Open()
    $sqlCommand = $sqlConnection.CreateCommand()
    $sqlCommand.CommandText = "CREATE DATABASE $DatabaseName"
    $sqlCommand.ExecuteNonQuery()
    Write-Host "The database $DatabaseName was created." -ForegroundColor Green
    $sqlConnection.Close()
} catch {
    Write-Host "An error occurred when creating the database $DatabaseName" -ForegroundColor Red
    Write-Host "Error details: $_" -ForegroundColor Red
}

# 3. Create a table named Client_A_Contacts in the database
Write-Host "Creating the table $TableName in database $DatabaseName..."
try {
    $sqlConnection.Open()
    $sqlCommand = $sqlConnection.CreateCommand()

    $CreateTableQuery = "USE $DatabaseName;"
    $CreateTableQuery += "CREATE TABLE $TableName ("
    $CreateTableQuery += "ClientID INT IDENTITY(1,1) PRIMARY KEY,"
    $CreateTableQuery += "First_Name NVARCHAR(100),"
    $CreateTableQuery += "Last_Name NVARCHAR(100),"
    $CreateTableQuery += "City NVARCHAR(70),"
    $CreateTableQuery += "County NVARCHAR(70),"
    $CreateTableQuery += "Zip_Code NVARCHAR(15),"
    $CreateTableQuery += "Office_Phone VARCHAR(15),"
    $CreateTableQuery += "Mobile_Phone VARCHAR(15)"
    $CreateTableQuery += ");"

    $sqlCommand.CommandText = $CreateTableQuery
    $sqlCommand.ExecuteNonQuery()
    Write-Host "The table $TableName was created." -ForegroundColor Green
    $sqlConnection.Close()
} catch {
    Write-Host "An error occurred when creating the table $TableName" -ForegroundColor Red
    Write-Host "Error details: $_" -ForegroundColor Red
}

# 4. Import the data from NewClientData.csv and write the data to the table Client_A_Contacts
Write-Host "Importing data from $CsvPath into table $TableName..."
try {
    $CsvData = Import-Csv -Path $CsvPath
    $sqlConnection.Open()

    foreach ($row in $CsvData) {
        $InsertQuery = "USE $DatabaseName;"
        $InsertQuery += "INSERT INTO $TableName (First_Name, Last_Name, City, County, Zip_Code, Office_Phone, Mobile_Phone)"
        $InsertQuery += " VALUES ('$($row.first_name)', '$($row.last_name)', '$($row.city)', '$($row.county)', '$($row.zip)', '$($row.officePhone)', '$($row.mobilePhone)');"

        $sqlCommand = $sqlConnection.CreateCommand()
        $sqlCommand.CommandText = $InsertQuery
        $sqlCommand.ExecuteNonQuery()
    }
    Write-Host "Data from $CsvPath has been successfully inserted into table $TableName." -ForegroundColor Green
    $sqlConnection.Close()
} catch {
    Write-Host "An error occurred while importing data" -ForegroundColor Red
    Write-Host "Error details: $_" -ForegroundColor Red
}

# 5. Generate the output file SqlResults.txt
Write-Host "Generating output file SqlResults.txt..."
try {
    $sqlConnection.Open()
    $sqlCommand = $sqlConnection.CreateCommand()
    $sqlCommand.CommandText = "SELECT * FROM dbo.Client_A_Contacts"

    # Execute the query and output results to file
    $reader = $sqlCommand.ExecuteReader()
    $output = ""
    while ($reader.Read()) {
        $row = ""
        for ($i = 0; $i -lt $reader.FieldCount; $i++) {
            $row += $reader.GetValue($i) + "`t"
        }
        $output += $row + "`n"
    }
    $reader.Close()

    # Write the output to a file
    $output | Out-File "$PSScriptRoot\SqlResults.txt"
    Write-Host "The output file SqlResults.txt has been generated." -ForegroundColor Green
    $sqlConnection.Close()
} catch {
    Write-Host "An error occurred while generating the output file" -ForegroundColor Red
    Write-Host "Error details: $_" -ForegroundColor Red
}
