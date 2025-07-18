using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PDFSlicer.Models;
using PdfSharpCore.Pdf;
using System.IO;

namespace PDFSlicer.Processing;

public static class PdfProcessor
{
    private const string DefaultProgramName = "Program";
    
    public static PdfInfo ParsePdfName(string fileName)
    {
        var parts = fileName.Split('_');
        return new PdfInfo
        {
            CertificateRange = parts.Length > 0 ? parts[0] : "Unknown",
            IssueDate = parts.Length > 1 ? FormatDate(parts[1]) : "01.01.00",
            ProgramName = parts.Length > 2 ? ShortenProgramName(parts[2]) : DefaultProgramName,
            Hours = parts.Length > 3 ? parts[3] : "Hours"
        };
    }

    public static ExcelRecord FindMatchingRecord(Dictionary<string, ExcelRecord> excelData, int pageNumber)
    {
        return excelData.Values.Skip(pageNumber - 1).FirstOrDefault();
    }

    public static string GenerateFileName(string originalPdfName, ExcelRecord record)
    {
        var cleanName = originalPdfName.Replace(".pdf", "");
        var cleanFullName = RemoveInvalidChars(record.FullName.Replace(" ", "_"));

        return $"{cleanName}-{cleanFullName}.pdf";
    }

    public static void SaveSinglePage(PdfPage page, string directory, string fileName)
    {
        using var document = new PdfDocument();
        document.AddPage(page);
        document.Save(Path.Combine(directory, fileName));
    }

    private static string ShortenProgramName(string programName)
    {
        if (string.IsNullOrWhiteSpace(programName))
        {
            return DefaultProgramName;
        }

        var words = programName.Split(' ');
        var shortened = new StringBuilder();

        foreach (var word in words)
        {
            shortened.Append(word.Length > 4 ? word.Substring(0, 4) : word);
            shortened.Append('_');
        }

        return shortened.ToString().TrimEnd('_');
    }

    private static string RemoveInvalidChars(string input)
    {
        var invalidChars = Path.GetInvalidFileNameChars();
        return new string(input.Where(c => !invalidChars.Contains(c)).ToArray());
    }

    private static string FormatDate(string date)
    {
        return date.Length switch
        {
            8 => date, // dd.MM.yy
            5 => $"{date}.{DateTime.Now:yy}", // dd.MM
            _ => "01.01.00"
        };
    }
}