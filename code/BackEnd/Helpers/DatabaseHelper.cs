using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using Microsoft.Extensions.Configuration;
using BackEnd.Models;
using System.Text.Json;

namespace BackEnd.Helpers
{
    public class DatabaseHelper
    {
        private readonly string _connectionString;

        public DatabaseHelper(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection") ?? 
                throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
        }

        public List<string> GetTableColumns(string tableName)
        {
            var columns = new List<string>();

            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                
                var schema = connection.GetSchema("Columns", new[] { null, null, tableName, null });
                
                if (schema.Rows.Count == 0)
                {
                    throw new Exception($"Table '{tableName}' not found in the database");
                }

                foreach (DataRow row in schema.Rows)
                {
                    string? columnName = row["COLUMN_NAME"]?.ToString();
                    if (!string.IsNullOrEmpty(columnName))
                    {
                        columns.Add(columnName);
                    }
                }
            }

            return columns;
        }

        public Dictionary<string, string> GetColumnDataTypes(string tableName)
        {
            var columnTypes = new Dictionary<string, string>();

            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                string query = @"
                    SELECT 
                        c.name AS ColumnName,
                        t.name AS DataType
                    FROM 
                        sys.columns c
                    INNER JOIN 
                        sys.types t ON c.user_type_id = t.user_type_id
                    INNER JOIN 
                        sys.tables tbl ON c.object_id = tbl.object_id
                    WHERE 
                        tbl.name = @TableName
                    ORDER BY 
                        c.column_id";

                using (var command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@TableName", tableName);

                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string? columnName = reader["ColumnName"]?.ToString();
                            string? dataType = reader["DataType"]?.ToString();
                            
                            if (!string.IsNullOrEmpty(columnName) && !string.IsNullOrEmpty(dataType))
                            {
                                columnTypes.Add(columnName, dataType);
                            }
                        }
                    }
                }
            }

            return columnTypes;
        }

        public bool TableExists(string tableName)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                
                string query = "SELECT COUNT(1) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = @TableName";
                
                using (var command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@TableName", tableName);
                    var result = (int)command.ExecuteScalar();
                    return result > 0;
                }
            }
        }

        public void InsertData(string tableName, List<string> columns, List<List<object>> rows)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        string columnList = string.Join(", ", columns);
                        string parameterList = string.Join(", ", columns.Select(c => "@" + c));
                        
                        string insertQuery = $"INSERT INTO {tableName} ({columnList}) VALUES ({parameterList})";

                        foreach (var row in rows)
                        {
                            using (var command = new SqlCommand(insertQuery, connection, transaction))
                            {
                                for (int i = 0; i < columns.Count; i++)
                                {
                                    var value = i < row.Count ? row[i] : DBNull.Value;
                                    
                                    // Handle JsonElement conversion
                                    if (value is JsonElement jsonElement)
                                    {
                                        value = ConvertJsonElementToObject(jsonElement);
                                    }
                                    
                                    command.Parameters.AddWithValue("@" + columns[i], value ?? DBNull.Value);
                                }

                                command.ExecuteNonQuery();
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
        }

        private object ConvertJsonElementToObject(JsonElement element)
        {
            switch (element.ValueKind)
            {
                case JsonValueKind.String:
                    return element.GetString() ?? string.Empty;
                case JsonValueKind.Number:
                    if (element.TryGetInt32(out int intValue))
                        return intValue;
                    if (element.TryGetInt64(out long longValue))
                        return longValue;
                    if (element.TryGetDecimal(out decimal decimalValue))
                        return decimalValue;
                    if (element.TryGetDouble(out double doubleValue))
                        return doubleValue;
                    return 0;
                case JsonValueKind.True:
                    return true;
                case JsonValueKind.False:
                    return false;
                case JsonValueKind.Null:
                    return DBNull.Value;
                case JsonValueKind.Object:
                    return element.ToString() ?? string.Empty;
                case JsonValueKind.Array:
                    return element.ToString() ?? string.Empty;
                default:
                    return DBNull.Value;
            }
        }
    }
} 