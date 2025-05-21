# Excel Data Import Backend

A .NET Core Web API backend for importing Excel data into SQL Server database tables.

## Features

- File Upload: Upload Excel (.xlsx) files
- Sheet Preview: View sheet names and data
- Data Verification: Check compatibility between Excel data and database tables
- Data Import: Insert verified data into the database

## API Endpoints

1. **Upload Excel File**
   - Route: POST `/api/upload`
   - Accepts an Excel file (.xlsx)
   - Returns a list of sheet names

2. **Preview Sheet Data**
   - Route: GET `/api/sheet-preview`
   - Parameters: `fileName`, `sheetName`
   - Returns the data of the selected sheet

3. **Verify Sheet Data**
   - Route: POST `/api/verify`
   - Body: sheet name and sheet data
   - Verifies compatibility between Excel and database table

4. **Insert Data**
   - Route: POST `/api/insert`
   - Body: sheet name and sheet data
   - Inserts data into the database table

## Prerequisites

- .NET 8.0 SDK
- SQL Server (Local instance or SQL Express)

## Configuration

The application connects to a local SQL Server database. You can configure the connection string in `appsettings.json`:

```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Server=localhost\\SQLEXPRESS;Database=YourDatabaseName;Trusted_Connection=True;"
  }
}
```

For SQL Authentication:

```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Server=localhost\\SQLEXPRESS;Database=YourDatabaseName;User Id=yourUsername;Password=yourPassword;"
  }
}
```

## Running the Application

1. Update the connection string in `appsettings.json`
2. Run the application:
   ```
   dotnet run
   ```
3. The API will be available at http://localhost:5000 (or https://localhost:5001)

## Technical Details

- Uses EPPlus for Excel file processing
- Uses System.Data.SqlClient for database operations
- Case-insensitive column name comparison
- Dynamic data type conversion and validation 