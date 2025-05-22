using System;
using System.Collections.Generic;

namespace BackEnd.Models
{
    public class ExcelFileInfo
    {
        public int Id { get; set; }
        public string FileName { get; set; } = string.Empty;
        public DateTime UploadDate { get; set; }
        public string Status { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
        public string SheetName { get; set; } = string.Empty;
        public List<string> Columns { get; set; } = new List<string>();
        public List<List<object>> Rows { get; set; } = new List<List<object>>();
        public Dictionary<string, object> Results { get; set; } = new Dictionary<string, object>();
    }

    public class SheetPreviewRequest
    {
        public required string FileName { get; set; }
        public required string SheetName { get; set; }
    }

    public class SheetData
    {
        public required string SheetName { get; set; }
        public required List<string> Headers { get; set; }
        public required List<List<object>> Rows { get; set; }
    }

    public class VerifyRequest
    {
        public required string SheetName { get; set; }
        public required SheetData SheetData { get; set; }
    }

    public class InsertRequest
    {
        public required string SheetName { get; set; }
        public required List<string> Columns { get; set; }
        public required List<List<object>> Rows { get; set; }
        public int? KeyColumnIndex { get; set; } // Optional key column index for upsert
    }

    public class ApiResponse<T>
    {
        public bool Success { get; set; }
        public required string Message { get; set; }
        public required T Data { get; set; }
    }

    public class VerificationResult
    {
        public bool IsValid { get; set; }
        public List<string> Warnings { get; set; } = new List<string>();
        public List<string> Errors { get; set; } = new List<string>();
    }

    public class SheetValidationRequest
    {
        public string SheetName { get; set; }
        public List<string> Columns { get; set; }
        public List<List<object>> Rows { get; set; }
    }

    public class CellValidationResult
    {
        public string Status { get; set; } // "valid", "length", "type"
        public string Message { get; set; }
    }

    public class SheetValidationResponse
    {
        public List<List<CellValidationResult>> Results { get; set; }
    }
} 