using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using BackEnd.Helpers;
using BackEnd.Models;

namespace BackEnd.Services
{
    public class ExcelService
    {
        private readonly string _uploadFolder;
        
        public ExcelService()
        {
            _uploadFolder = Path.Combine(Directory.GetCurrentDirectory(), "Uploads");
            if (!Directory.Exists(_uploadFolder))
            {
                Directory.CreateDirectory(_uploadFolder);
            }
        }
        
        public async Task<ExcelFileInfo> UploadExcelFile(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                throw new ArgumentException("No file was uploaded");
            }
            
            var fileExtension = Path.GetExtension(file.FileName).ToLower();
            if (fileExtension != ".xlsx")
            {
                throw new ArgumentException("Only .xlsx files are supported");
            }
            
            // Generate unique filename to prevent overwriting
            var uniqueFileName = $"{Guid.NewGuid()}_{file.FileName}";
            var filePath = Path.Combine(_uploadFolder, uniqueFileName);
            
            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }
            
            var fileInfo = ExcelHelper.GetSheetNames(filePath);
            return fileInfo;
        }
        
        public SheetData GetSheetPreview(string fileName, string sheetName)
        {
            var filePath = Path.Combine(_uploadFolder, fileName);
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"File {fileName} not found");
            }
            
            return ExcelHelper.GetSheetData(filePath, sheetName);
        }
    }
} 