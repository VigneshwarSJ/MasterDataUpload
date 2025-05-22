# Master Data Upload System

A full-stack solution for uploading, validating, and importing Excel data into SQL Server database tables, with a modern Angular frontend and a .NET Core Web API backend.

---

## Features

- **Excel Upload:** Upload `.xlsx` files and select sheets for import.
- **Sheet Preview:** View sheet names and preview data before import.
- **Data Validation:** Check Excel data against database schema (type, length, required fields).
- **Duplicate Detection:** Detect duplicate values in columns (e.g., name columns).
- **Dynamic Table Mapping:** Sheet name maps to table name; columns map to table columns.
- **Upsert Logic:** If the first column value exists, the row is updated; otherwise, it is inserted.
- **Special Handling:** Custom logic for certain tables (e.g., Counterparties).
- **Error Feedback:** User-friendly error messages for validation and import issues.

---

## Project Structure

```
/FrontEnd      # Angular 16+ SPA for user interaction
/BackEnd       # .NET Core Web API for Excel processing and DB import
```

---

## Backend (BackEnd)

- **Tech:** .NET 8.0, System.Data.SqlClient, EPPlus (for Excel), SQL Server
- **Main Controller:** `ExcelController.cs`
- **Key Services:** `ExcelService`, `DataValidationService`
- **Helpers:** `ExcelHelper`, `DatabaseHelper`
- **Models:** `ExcelFileInfo`, `SheetData`, `InsertRequest`, etc.

### API Endpoints

| Endpoint            | Method | Description                                 |
|---------------------|--------|---------------------------------------------|
| `/api/upload`       | POST   | Upload Excel file                           |
| `/api/sheet-preview`| GET    | Preview sheet data                          |
| `/api/verify`       | POST   | Verify sheet data against DB schema         |
| `/api/insert`       | POST   | Upsert data into mapped table               |
| `/api/verify-sheet` | POST   | Validate each cell in a sheet (type/length) |

### Configuration

Edit `BackEnd/appsettings.json` for your SQL Server connection string.

```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Server=localhost\\SQLEXPRESS;Database=YourDatabaseName;Trusted_Connection=True;"
  }
}
```

### Running the Backend

```sh
cd BackEnd
dotnet run
```
API will be available at `https://localhost:7001` (see console output for port).

---

## Frontend (FrontEnd)

- **Tech:** Angular 16+, TypeScript, Bootstrap (or similar)
- **Main Components:** `welcome`, `login`, `data-modal`
- **Proxy:** Configured in `proxy.conf.json` to forward `/api` calls to backend.

### Running the Frontend

```sh
cd FrontEnd
npm install
ng serve
```
App will be available at [http://localhost:4200](http://localhost:4200).

---

## Usage

1. **Login** (if enabled).
2. **Upload** an Excel file.
3. **Select** a sheet to preview.
4. **Verify** the sheet to check for schema/data issues.
5. **Insert** to upsert data into the mapped database table.

---

## Notes

- The first column in your sheet is used as the unique key for upsert.
- Special logic is applied for the `Counterparties` table (ID lookups, etc.).
- All other tables are handled dynamically.
- Validation and error feedback are provided in the UI.

---

## Development

- **Backend:** .NET 8.0+ required. Update connection string as needed.
- **Frontend:** Node.js 18+ and Angular CLI required.

---

## License

MIT (or your chosen license) 