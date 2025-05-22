using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using BackEnd.Models;
using BackEnd.Services;
using Microsoft.Extensions.Configuration;
using System.Text.Json;
using System.Collections.Generic;
using System.Linq;
using Microsoft.EntityFrameworkCore;

namespace BackEnd.Controllers
{
    [ApiController]
    [Route("api")]
    public class ExcelController : ControllerBase
    {
        private readonly ExcelService _excelService;
        private readonly DataValidationService _dataValidationService;
        private readonly IConfiguration _configuration;
        private readonly ApplicationDbContext _dbContext;

        public ExcelController(
            ExcelService excelService, 
            DataValidationService dataValidationService, 
            IConfiguration configuration,
            ApplicationDbContext dbContext)
        {
            _excelService = excelService;
            _dataValidationService = dataValidationService;
            _configuration = configuration;
            _dbContext = dbContext;
        }

        [HttpPost("upload")]
        public async Task<IActionResult> UploadExcel(IFormFile file)
        {
            try
            {
                var result = await _excelService.UploadExcelFile(file);
                return Ok(new ApiResponse<ExcelFileInfo>
                {
                    Success = true,
                    Message = "File uploaded successfully",
                    Data = result
                });
            }
            catch (Exception ex)
            {
                return BadRequest(new ApiResponse<object>
                {
                    Success = false,
                    Message = ex.Message,
                    Data = new { error = ex.Message }
                });
            }
        }

        [HttpGet("sheet-preview")]
        public IActionResult GetSheetPreview([FromQuery] SheetPreviewRequest request)
        {
            try
            {
                var result = _excelService.GetSheetPreview(request.FileName, request.SheetName);
                return Ok(new ApiResponse<SheetData>
                {
                    Success = true,
                    Message = "Sheet data retrieved successfully",
                    Data = result
                });
            }
            catch (Exception ex)
            {
                return BadRequest(new ApiResponse<object>
                {
                    Success = false,
                    Message = ex.Message,
                    Data = new { error = ex.Message }
                });
            }
        }

        [HttpPost("verify")]
        public IActionResult VerifyData([FromBody] VerifyRequest request)
        {
            try
            {
                var result = _dataValidationService.VerifySheetData(request.SheetName, request.SheetData);
                return Ok(new ApiResponse<VerificationResult>
                {
                    Success = true,
                    Message = result.IsValid ? "Verification passed" : "Verification failed",
                    Data = result
                });
            }
            catch (Exception ex)
            {
                return BadRequest(new ApiResponse<object>
                {
                    Success = false,
                    Message = ex.Message,
                    Data = new { error = ex.Message }
                });
            }
        }

        [HttpPost("insert")]
        public async Task<IActionResult> InsertData([FromBody] InsertRequest request)
        {
            try
            {
                var tableName = request.SheetName;
                var columns = request.Columns;
                var keyColumnIndex = request.KeyColumnIndex ?? 0;
                var keyColumn = columns[keyColumnIndex];

                // First verify if the table exists
                using (var connection = new System.Data.SqlClient.SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
                {
                    await connection.OpenAsync();
                    var cmd = new System.Data.SqlClient.SqlCommand(
                        "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = @TableName", 
                        connection);
                    cmd.Parameters.AddWithValue("@TableName", tableName);
                    var tableExists = (int)await cmd.ExecuteScalarAsync() > 0;

                    if (!tableExists)
                    {
                        return BadRequest(new ApiResponse<object>
                        {
                            Success = false,
                            Message = $"Table '{tableName}' does not exist in the database",
                            Data = null
                        });
                    }

                    // Get table schema
                    var schema = new Dictionary<string, (string DataType, bool IsNullable)>();
                    cmd = new System.Data.SqlClient.SqlCommand(
                        @"SELECT COLUMN_NAME, DATA_TYPE, IS_NULLABLE 
                          FROM INFORMATION_SCHEMA.COLUMNS 
                          WHERE TABLE_NAME = @TableName", 
                        connection);
                    cmd.Parameters.AddWithValue("@TableName", tableName);
                    using (var reader = await cmd.ExecuteReaderAsync())
                    {
                        while (await reader.ReadAsync())
                        {
                            var col = reader["COLUMN_NAME"].ToString();
                            schema[col] = (
                                reader["DATA_TYPE"].ToString(),
                                reader["IS_NULLABLE"].ToString() == "YES"
                            );
                        }
                    }

                    // Validate columns exist in the table
                    foreach (var column in columns)
                    {
                        if (!schema.ContainsKey(column))
                        {
                            return BadRequest(new ApiResponse<object>
                            {
                                Success = false,
                                Message = $"Column '{column}' does not exist in table '{tableName}'",
                                Data = null
                            });
                        }
                    }
                }

                // Special handling for Counterparties table
                if (string.Equals(tableName, "Counterparties", StringComparison.OrdinalIgnoreCase))
                {
                    // Validate required columns for Counterparties
                    var requiredColumns = new[] { 
                        "COunterpartyname", 
                        "Counterpartytype", 
                        "COunterpartyParentId",
                        "AP_PaymentTermId",
                        "AP_PaymentMethodId",
                        "AR_PaymentTermId",
                        "AR_PaymentMethodId",
                        "CustomerProfile",
                        "VendorProfile"
                    };

                    foreach (var requiredCol in requiredColumns)
                    {
                        if (!columns.Any(c => string.Equals(c, requiredCol, StringComparison.OrdinalIgnoreCase)))
                        {
                            return BadRequest(new ApiResponse<object>
                            {
                                Success = false,
                                Message = $"Required column '{requiredCol}' is missing in the Counterparties sheet",
                                Data = null
                            });
                        }
                    }
                }

                // Process the data for all tables
                var processedRows = new List<List<object>>();
                foreach (var row in request.Rows)
                {
                    var processedRow = new List<object>();
                    for (int i = 0; i < columns.Count; i++)
                    {
                        var columnName = columns[i].ToLower();
                        var value = i < row.Count ? GetValue(row[i]) : null;

                        // Special handling for Counterparties table
                        if (string.Equals(tableName, "Counterparties", StringComparison.OrdinalIgnoreCase))
                        {
                            switch (columnName)
                            {
                                case "ap_paymenttermid":
                                case "ar_paymenttermid":
                                    processedRow.Add(await LookupIdWithErrorAsync(
                                        "SELECT PaymentTermId FROM paymentterms WHERE PaymentTermName = @val",
                                        value,
                                        $"Payment Term '{value}' not found in paymentterms table"
                                    ));
                                    break;

                                case "ap_paymentmethodid":
                                case "ar_paymentmethodid":
                                    processedRow.Add(await LookupIdWithErrorAsync(
                                        "SELECT PaymentMethodId FROM paymentmethods WHERE PaymentMethodCode = @val",
                                        value,
                                        $"Payment Method '{value}' not found in paymentmethods table"
                                    ));
                                    break;

                                case "customerprofile":
                                    processedRow.Add(await LookupIdWithErrorAsync(
                                        "SELECT CounterPartyPostingGroupId FROM CounterPartyPostingGroups WHERE CounterPartyPostingGroupName = @val AND CounterPartyPostingGroupType = 'C'",
                                        value,
                                        $"Customer Profile '{value}' not found in CounterPartyPostingGroups table"
                                    ));
                                    break;

                                case "vendorprofile":
                                    processedRow.Add(await LookupIdWithErrorAsync(
                                        "SELECT CounterPartyPostingGroupId FROM CounterPartyPostingGroups WHERE CounterPartyPostingGroupName = @val AND CounterPartyPostingGroupType = 'V'",
                                        value,
                                        $"Vendor Profile '{value}' not found in CounterPartyPostingGroups table"
                                    ));
                                    break;

                                default:
                                    processedRow.Add(value);
                                    break;
                            }
                        }
                        else
                        {
                            // For all other tables, just add the value as is
                            processedRow.Add(value);
                        }
                    }
                    processedRows.Add(processedRow);
                }

                // Build and execute the dynamic SQL
                var keyCol = columns[0];
                var updateCols = columns.Skip(1).ToList();
                using (var connection = new System.Data.SqlClient.SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
                {
                    await connection.OpenAsync();
                    using (var transaction = connection.BeginTransaction())
                    {
                        try
                        {
                            foreach (var row in processedRows)
                            {
                                // Build MERGE statement
                                var mergeSql = $@"
MERGE INTO [{tableName}] AS Target
USING (SELECT @{keyCol} AS [{keyCol}] {string.Concat(updateCols.Select(c => $", @{c} AS [{c}]"))}) AS Source
    ON Target.[{keyCol}] = Source.[{keyCol}]
WHEN MATCHED THEN
    UPDATE SET {string.Join(", ", updateCols.Select(c => $"Target.[{c}] = Source.[{c}]"))}
WHEN NOT MATCHED THEN
    INSERT ({string.Join(", ", columns.Select(c => $"[{c}]"))}) VALUES ({string.Join(", ", columns.Select(c => $"Source.[{c}]"))});";

                                using (var command = new System.Data.SqlClient.SqlCommand(mergeSql, connection, transaction))
                                {
                                    for (int i = 0; i < columns.Count; i++)
                                    {
                                        var value = i < row.Count ? GetValue(row[i]) : DBNull.Value;
                                        command.Parameters.AddWithValue("@" + columns[i], value ?? DBNull.Value);
                                    }
                                    await command.ExecuteNonQueryAsync();
                                }
                            }
                            transaction.Commit();
                        }
                        catch
                        {
                            transaction.Rollback();
                            throw;
                        }
                    }
                }

                return Ok(new ApiResponse<object>
                {
                    Success = true,
                    Message = $"Data inserted successfully into {tableName}",
                    Data = new { rowsInserted = processedRows.Count }
                });
            }
            catch (Exception ex)
            {
                return BadRequest(new ApiResponse<object>
                {
                    Success = false,
                    Message = ex.Message,
                    Data = new { error = ex.Message }
                });
            }
        }

        private object GetValue(object value)
        {
            if (value is System.Text.Json.JsonElement jsonElement)
            {
                return jsonElement.ValueKind switch
                {
                    System.Text.Json.JsonValueKind.String => jsonElement.GetString(),
                    System.Text.Json.JsonValueKind.Number => jsonElement.TryGetInt64(out var l) ? l : jsonElement.GetDouble(),
                    System.Text.Json.JsonValueKind.True => true,
                    System.Text.Json.JsonValueKind.False => false,
                    System.Text.Json.JsonValueKind.Null => null,
                    _ => jsonElement.ToString()
                };
            }
            return value;
        }

        private async Task<object> LookupIdWithErrorAsync(string sql, object value, string notFoundMessage)
        {
            var result = await _dbContext.Database.SqlQueryRaw<int>(sql, new[] { value }).FirstOrDefaultAsync();
            if (result == 0)
            {
                throw new Exception(notFoundMessage);
            }
            return result;
        }

        private async Task<IActionResult> ProcessGenericSheet(InsertRequest request)
        {
            // Original insert logic for non-Counterparties sheets
            // ... (copy the original non-Counterparties logic here)
            return Ok(new { message = "Data inserted/updated successfully!" });
        }

        [HttpPost("verify-sheet")]
        public async Task<IActionResult> VerifySheet([FromBody] SheetValidationRequest request)
        {
            var results = new List<List<CellValidationResult>>();
            var tableName = request.SheetName;
            var columns = request.Columns;

            // Get schema info for the columns in the sheet
            var schema = new Dictionary<string, (string DataType, int? MaxLength, bool IsNullable)>();
            var connectionString = _configuration.GetConnectionString("DefaultConnection");
            using (var connection = new System.Data.SqlClient.SqlConnection(connectionString))
            {
                await connection.OpenAsync();
                var cmd = new System.Data.SqlClient.SqlCommand(
                    @"SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, IS_NULLABLE
                      FROM INFORMATION_SCHEMA.COLUMNS
                      WHERE TABLE_NAME = @TableName", connection);
                cmd.Parameters.AddWithValue("@TableName", tableName);
                using (var reader = await cmd.ExecuteReaderAsync())
                {
                    while (await reader.ReadAsync())
                    {
                        var col = reader["COLUMN_NAME"].ToString();
                        if (!string.IsNullOrWhiteSpace(col) && columns.Contains(col))
                        {
                            schema[col] = (
                                reader["DATA_TYPE"].ToString(),
                                reader["CHARACTER_MAXIMUM_LENGTH"] == DBNull.Value ? null : Convert.ToInt32(reader["CHARACTER_MAXIMUM_LENGTH"]),
                                reader["IS_NULLABLE"].ToString() == "YES"
                            );
                        }
                    }
                }
            }

            // Validate each cell
            foreach (var row in request.Rows)
            {
                var rowResult = new List<CellValidationResult>();
                for (int i = 0; i < columns.Count; i++)
                {
                    var col = columns[i];
                    if (string.IsNullOrWhiteSpace(col))
                    {
                        rowResult.Add(new CellValidationResult { Status = "valid", Message = "" });
                        continue;
                    }
                    var value = row.Count > i ? row[i] : null;
                    if (!schema.ContainsKey(col))
                    {
                        rowResult.Add(new CellValidationResult { Status = "valid", Message = "" });
                        continue;
                    }
                    var (dataType, maxLength, isNullable) = schema[col];

                    // Type check
                    bool typeOk = true;
                    string typeMsg = "";
                    if (value == null || value is System.Text.Json.JsonElement je && je.ValueKind == System.Text.Json.JsonValueKind.Null)
                    {
                        if (!isNullable)
                        {
                            typeOk = false;
                            typeMsg = "Value required";
                        }
                    }
                    else
                    {
                        object nativeValue = value;
                        if (value is System.Text.Json.JsonElement jsonElement)
                        {
                            switch (jsonElement.ValueKind)
                            {
                                case System.Text.Json.JsonValueKind.String:
                                    nativeValue = jsonElement.GetString();
                                    break;
                                case System.Text.Json.JsonValueKind.Number:
                                    nativeValue = jsonElement.TryGetInt64(out var l) ? l : jsonElement.GetDouble();
                                    break;
                                case System.Text.Json.JsonValueKind.True:
                                case System.Text.Json.JsonValueKind.False:
                                    nativeValue = jsonElement.GetBoolean();
                                    break;
                                default:
                                    nativeValue = jsonElement.ToString();
                                    break;
                            }
                        }

                        switch (dataType.ToLower())
                        {
                            case "int":
                            case "bigint":
                            case "smallint":
                            case "tinyint":
                                if (!(nativeValue is int || nativeValue is long || int.TryParse(nativeValue?.ToString(), out _)))
                                {
                                    typeOk = false;
                                    typeMsg = $"Expected integer but got '{nativeValue}'";
                                }
                                break;
                            case "bit":
                                if (!(nativeValue is bool || nativeValue?.ToString() == "0" || nativeValue?.ToString() == "1"))
                                {
                                    typeOk = false;
                                    typeMsg = $"Expected boolean (0/1) but got '{nativeValue}'";
                                }
                                break;
                            case "decimal":
                            case "numeric":
                            case "float":
                            case "real":
                                if (!decimal.TryParse(nativeValue?.ToString(), out _))
                                {
                                    typeOk = false;
                                    typeMsg = $"Expected decimal but got '{nativeValue}'";
                                }
                                break;
                            case "varchar":
                            case "nvarchar":
                            case "char":
                            case "nchar":
                            case "text":
                            case "ntext":
                                // Always ok for type, check length below
                                break;
                        }
                    }

                    // Length check
                    bool lengthOk = true;
                    string lengthMsg = "";
                    if (typeOk && maxLength.HasValue && value != null)
                    {
                        string strVal = value is System.Text.Json.JsonElement je2 && je2.ValueKind == System.Text.Json.JsonValueKind.String
                            ? je2.GetString()
                            : value.ToString();
                        if (strVal != null && strVal.Length > maxLength.Value)
                        {
                            lengthOk = false;
                            lengthMsg = $"Exceeds max length of {maxLength.Value}";
                        }
                    }

                    // Set status
                    if (!typeOk)
                        rowResult.Add(new CellValidationResult { Status = "type", Message = typeMsg });
                    else if (!lengthOk)
                        rowResult.Add(new CellValidationResult { Status = "length", Message = lengthMsg });
                    else
                        rowResult.Add(new CellValidationResult { Status = "valid", Message = "" });
                }
                results.Add(rowResult);
            }

            return Ok(new SheetValidationResponse { Results = results });
        }
    }
} 