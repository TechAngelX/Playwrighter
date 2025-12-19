// Services/ExcelService.cs

using OfficeOpenXml;
using Playwrighter.Models;

namespace Playwrighter.Services;

public class ExcelService : IExcelService
{
    public List<string> GetSheetNames(string filePath)
    {
        var sheetNames = new List<string>();
        using var package = new ExcelPackage(new FileInfo(filePath));
        foreach (var worksheet in package.Workbook.Worksheets)
        {
            sheetNames.Add(worksheet.Name);
        }
        return sheetNames;
    }

    public List<StudentRecord> LoadStudentsFromFile(string filePath, string sheetName = "Dept In-tray")
    {
        var students = new List<StudentRecord>();
        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets[sheetName];

        if (worksheet == null)
        {
            worksheet = package.Workbook.Worksheets.FirstOrDefault();
            if (worksheet == null) throw new InvalidOperationException("No worksheets found.");
        }
        
        int studentNoCol = -1;
        int decisionCol = -1;
        int nameCol = -1;
        int programmeCol = -1;
        
        int headerRow = 1;
        int colCount = worksheet.Dimension?.Columns ?? 0;
        
        // Log all column headers for debugging
        Console.WriteLine("=== Excel Column Headers ===");
        for (int col = 1; col <= colCount; col++)
        {
            string header = worksheet.Cells[headerRow, col].Value?.ToString()?.Trim() ?? "";
            Console.WriteLine($"  Column {col}: '{header}'");
        }
        
        // --- 1. First Pass: Strict Matches ---
        for (int col = 1; col <= colCount; col++)
        {
            string header = worksheet.Cells[headerRow, col].Value?.ToString()?.Trim().ToLowerInvariant() ?? "";
            
            if (string.IsNullOrEmpty(header)) continue;

            // Student No
            if (studentNoCol == -1 && (header == "studentno" || header == "student_no" || header == "student number" || header == "student id" || header == "id"))
                studentNoCol = col;
            
            // Decision
            if (decisionCol == -1 && (header == "decision" || header == "status" || header == "offer"))
                decisionCol = col;
            
            // Name
            if (nameCol == -1 && (header == "name" || header == "applicant name" || header == "student name"))
                nameCol = col;

            // Programme - Strict match for "programme"
            if (programmeCol == -1 && header == "programme")
                programmeCol = col;
        }

        // --- 2. Second Pass: Loose Matches (if strict failed) ---
        for (int col = 1; col <= colCount; col++)
        {
            string header = worksheet.Cells[headerRow, col].Value?.ToString()?.Trim().ToLowerInvariant() ?? "";
            if (string.IsNullOrEmpty(header)) continue;

            if (studentNoCol == -1 && header.Contains("student") && header.Contains("no")) studentNoCol = col;
            if (decisionCol == -1 && header.Contains("decision")) decisionCol = col;
            
            // Loose programme match
            if (programmeCol == -1 && (header == "prog" || header == "progcode" || header == "prog code" || header == "progshort" || header == "route"))
                programmeCol = col;
        }
        
        Console.WriteLine($"=== Detected Columns ===");
        Console.WriteLine($"  StudentNo column: {studentNoCol}");
        Console.WriteLine($"  Decision column: {decisionCol}");
        Console.WriteLine($"  Name column: {nameCol}");
        Console.WriteLine($"  Programme column: {programmeCol}");
        
        if (studentNoCol == -1) throw new InvalidOperationException("Could not find StudentNo column.");
        if (decisionCol == -1) throw new InvalidOperationException("Could not find Decision column.");
        
        if (programmeCol == -1)
        {
            Console.WriteLine("WARNING: Programme column not found! Row matching may fail.");
        }

        int rowCount = worksheet.Dimension?.Rows ?? 0;
        
        for (int row = headerRow + 1; row <= rowCount; row++)
        {
            string? studentNo = worksheet.Cells[row, studentNoCol].Value?.ToString()?.Trim();
            if (string.IsNullOrWhiteSpace(studentNo)) continue;

            string programmeValue = programmeCol > 0 ? worksheet.Cells[row, programmeCol].Value?.ToString()?.Trim() ?? "" : "";

            var record = new StudentRecord
            {
                StudentNo = studentNo,
                Decision = worksheet.Cells[row, decisionCol].Value?.ToString()?.Trim() ?? "",
                Name = nameCol > 0 ? worksheet.Cells[row, nameCol].Value?.ToString()?.Trim() ?? "" : "",
                Programme = programmeValue,
                Status = ProcessingStatus.Pending
            };

            // Debug: Log first few records
            if (students.Count < 5)
            {
                Console.WriteLine($"  Loaded: StudentNo={record.StudentNo}, Decision={record.Decision}, Programme='{record.Programme}'");
            }

            students.Add(record);
        }
        
        Console.WriteLine($"=== Total loaded: {students.Count} students ===");
        
        return students;
    }
}
