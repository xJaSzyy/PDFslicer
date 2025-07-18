using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using PDFSlicer.Enums;
using PDFSlicer.Models;

namespace PDFSlicer.Processing;

public static class ExcelParser
{
    public static Dictionary<string, ExcelRecord> Parse(string filePath, int startRow)
    {
        var records = new Dictionary<string, ExcelRecord>();

        using var workbook = new XLWorkbook(filePath);
        var worksheet = workbook.Worksheet(1);
        var lastRow = worksheet.LastRowUsed()!.RowNumber();

        for (var row = startRow; row <= lastRow; row++)
        {
            try
            {
                var record = new ExcelRecord
                {
                    RowNumber = row,
                    RegistrationNumber = worksheet.Cell(row, (int)ExcelColumns.RegistrationNumber).GetString().Trim(),
                    DocumentName = worksheet.Cell(row, (int)ExcelColumns.DocumentName).GetString().Trim(),
                    CertificateNumber = worksheet.Cell(row, (int)ExcelColumns.CertificateNumber).GetString().Trim(),
                    IssueDate = FormatDate(worksheet.Cell(row, (int)ExcelColumns.IssueDate).GetString().Trim()),
                    FullName = worksheet.Cell(row, (int)ExcelColumns.FullName).GetString().Trim(),
                    ProgramName = worksheet.Cell(row, (int)ExcelColumns.ProgramName).GetString().Trim(),
                    Hours = worksheet.Cell(row, (int)ExcelColumns.Hours).GetString().Trim()
                };

                var key = $"{record.FullName}_{record.CertificateNumber}_{record.IssueDate}";
                records[key] = record;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing row {row}: {ex.Message}");
            }
        }

        return records;
    }

    private static string FormatDate(string date)
    {
        if (DateTime.TryParse(date, out var result))
        {
            return result.ToString("dd.MM.yy");
        }

        return date.Length switch
        {
            10 => date.Substring(0, 5) + date.Substring(8, 2), // dd.MM.yyyy
            8 => date, // dd.MM.yy
            5 => date + DateTime.Now.ToString(".yy"), // dd.MM
            _ => "01.01.00"
        };
    }
}