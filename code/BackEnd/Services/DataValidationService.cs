using System;
using System.Collections.Generic;
using System.Linq;
using BackEnd.Helpers;
using BackEnd.Models;
using Microsoft.Extensions.Configuration;

namespace BackEnd.Services
{
    public class DataValidationService
    {
        private readonly DatabaseHelper _databaseHelper;

        public DataValidationService(IConfiguration configuration)
        {
            _databaseHelper = new DatabaseHelper(configuration);
        }

        public VerificationResult VerifySheetData(string tableName, SheetData sheetData)
        {
            var result = new VerificationResult
            {
                IsValid = true
            };

            // Check if table exists
            if (!_databaseHelper.TableExists(tableName))
            {
                result.IsValid = false;
                result.Errors.Add($"Table '{tableName}' does not exist in the database");
                return result;
            }

            // Get table columns and data types
            var tableColumns = _databaseHelper.GetTableColumns(tableName);
            var columnDataTypes = _databaseHelper.GetColumnDataTypes(tableName);
            
            // Check if all Excel columns exist in the database table
            foreach (var header in sheetData.Headers)
            {
                if (!tableColumns.Contains(header, StringComparer.OrdinalIgnoreCase))
                {
                    result.IsValid = false;
                    result.Errors.Add($"Column '{header}' does not exist in table '{tableName}'");
                }
            }

            // Check if any required columns in the database are missing from Excel
            foreach (var column in tableColumns)
            {
                if (!sheetData.Headers.Contains(column, StringComparer.OrdinalIgnoreCase))
                {
                    result.Warnings.Add($"Column '{column}' from table '{tableName}' is not present in the Excel data");
                }
            }

            // Check data types for each column
            if (sheetData.Rows.Count > 0)
            {
                for (int colIndex = 0; colIndex < sheetData.Headers.Count; colIndex++)
                {
                    var header = sheetData.Headers[colIndex];
                    if (tableColumns.Contains(header, StringComparer.OrdinalIgnoreCase))
                    {
                        var sqlDataType = columnDataTypes[tableColumns.First(c => 
                            string.Equals(c, header, StringComparison.OrdinalIgnoreCase))];
                        
                        // Check sample values to ensure data type compatibility
                        for (int rowIndex = 0; rowIndex < Math.Min(10, sheetData.Rows.Count); rowIndex++)
                        {
                            var cellValue = sheetData.Rows[rowIndex][colIndex];
                            if (cellValue != null)
                            {
                                var typeMismatch = CheckDataTypeCompatibility(cellValue, sqlDataType);
                                if (typeMismatch != null)
                                {
                                    result.Warnings.Add($"Column '{header}': {typeMismatch}");
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            return result;
        }

        public void InsertData(string tableName, SheetData sheetData)
        {
            // Get the table columns to ensure we only insert into existing columns
            var tableColumns = _databaseHelper.GetTableColumns(tableName);
            
            // Filter headers to include only those that exist in the table
            var validHeaders = sheetData.Headers
                .Where(h => tableColumns.Contains(h, StringComparer.OrdinalIgnoreCase))
                .ToList();
            
            // Map Excel headers to actual table column names with correct casing
            var mappedColumns = validHeaders
                .Select(h => tableColumns.First(c => string.Equals(c, h, StringComparison.OrdinalIgnoreCase)))
                .ToList();
            
            // Prepare data rows for insertion
            var dataRows = new List<List<object>>();
            foreach (var row in sheetData.Rows)
            {
                var dataRow = new List<object>();
                for (int i = 0; i < sheetData.Headers.Count; i++)
                {
                    var header = sheetData.Headers[i];
                    if (validHeaders.Contains(header, StringComparer.OrdinalIgnoreCase))
                    {
                        dataRow.Add(row[i]);
                    }
                }
                dataRows.Add(dataRow);
            }
            
            // Insert the data
            _databaseHelper.InsertData(tableName, mappedColumns, dataRows);
        }

        private string? CheckDataTypeCompatibility(object value, string sqlDataType)
        {
            try
            {
                switch (sqlDataType.ToLower())
                {
                    case "int":
                    case "bigint":
                    case "smallint":
                    case "tinyint":
                        Convert.ToInt64(value);
                        break;
                    case "decimal":
                    case "numeric":
                    case "float":
                    case "real":
                    case "money":
                    case "smallmoney":
                        Convert.ToDecimal(value);
                        break;
                    case "bit":
                        Convert.ToBoolean(value);
                        break;
                    case "datetime":
                    case "date":
                    case "datetime2":
                    case "smalldatetime":
                        Convert.ToDateTime(value);
                        break;
                    case "uniqueidentifier":
                        if (value.ToString() != null)
                        {
                            Guid.Parse(value.ToString()!);
                        }
                        break;
                    // For text types (nvarchar, varchar, etc.), any value is acceptable
                    default:
                        break;
                }
                
                return null; // No error, types are compatible
            }
            catch
            {
                return $"Value '{value}' is not compatible with SQL data type '{sqlDataType}'";
            }
        }
    }
} 