using System.Collections.Generic;

namespace BackEnd.Models
{
    public class ExcelFileInfo
    {
        public required string FileName { get; set; }
        public required List<string> SheetNames { get; set; }
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