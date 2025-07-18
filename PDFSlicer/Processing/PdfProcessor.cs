using System.Collections.Generic;
using System.Linq;
using PDFSlicer.Models;
using PdfSharpCore.Pdf;
using System.IO;

namespace PDFSlicer.Processing;

public static class PdfProcessor
{
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

    private static string RemoveInvalidChars(string input)
    {
        var invalidChars = Path.GetInvalidFileNameChars();
        return new string(input.Where(c => !invalidChars.Contains(c)).ToArray());
    }
}