using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using BackEnd.Models;

namespace BackEnd.Helpers
{
    public class ExcelHelper
    {
        static ExcelHelper()
        {
            // Set EPPlus license for non-commercial use
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }
        
        public static ExcelFileInfo GetSheetNames(string filePath)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var fileInfo = new ExcelFileInfo
                {
                    FileName = Path.GetFileName(filePath),
                    SheetNames = package.Workbook.Worksheets.Select(s => s.Name).ToList()
                };
                
                return fileInfo;
            }
        }

        public static SheetData GetSheetData(string filePath, string sheetName)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[sheetName];
                if (worksheet == null)
                {
                    throw new Exception($"Sheet {sheetName} not found");
                }

                var dimension = worksheet.Dimension;
                if (dimension == null)
                {
                    // Empty worksheet
                    return new SheetData
                    {
                        SheetName = sheetName,
                        Headers = new List<string>(),
                        Rows = new List<List<object>>()
                    };
                }

                var rows = dimension.Rows;
                var cols = dimension.Columns;

                var headers = new List<string>();
                for (int col = 1; col <= cols; col++)
                {
                    var cellValue = worksheet.Cells[1, col].Value?.ToString() ?? string.Empty;
                    headers.Add(cellValue);
                }

                var dataRows = new List<List<object>>();
                for (int row = 2; row <= rows; row++)
                {
                    var dataRow = new List<object>();
                    for (int col = 1; col <= cols; col++)
                    {
                        dataRow.Add(worksheet.Cells[row, col].Value);
                    }
                    dataRows.Add(dataRow);
                }

                return new SheetData
                {
                    SheetName = sheetName,
                    Headers = headers,
                    Rows = dataRows
                };
            }
        }
    }
} 