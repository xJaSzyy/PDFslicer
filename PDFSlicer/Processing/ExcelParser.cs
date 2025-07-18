using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using PDFSlicer.Models;

namespace PDFSlicer.Processing;

public static class ExcelParser
{
    public static Dictionary<string, ExcelRecord> Parse(string filePath, int startRow)
    {
        var records = new Dictionary<string, ExcelRecord>();

        using (var workbook = new XLWorkbook(filePath))
        {
            var worksheet = workbook.Worksheet(1);
            var lastRow = worksheet.LastRowUsed().RowNumber();

            for (int row = startRow; row <= lastRow; row++)
            {
                try
                {
                    var record = new ExcelRecord
                    {
                        RowNumber = row,
                        RegistrationNumber = worksheet.Cell(row, 3).GetString().Trim(),  // Столбец c
                        DocumentName = worksheet.Cell(row, 5).GetString().Trim(),       // Столбец e
                        CertificateNumber = worksheet.Cell(row, 6).GetString().Trim(),  // Столбец f
                        IssueDate = FormatDate(worksheet.Cell(row, 7).GetString().Trim()), // Столбец g
                        FullName = worksheet.Cell(row, 8).GetString().Trim(),           // Столбец h
                        ProgramName = worksheet.Cell(row, 9).GetString().Trim(),        // Столбец i
                        Hours = worksheet.Cell(row, 15).GetString().Trim()             // Столбец o
                    };

                    var key = $"{record.FullName}_{record.CertificateNumber}_{record.IssueDate}";
                    records[key] = record;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка обработки строки {row}: {ex.Message}");
                }
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