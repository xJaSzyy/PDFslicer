using ClosedXML.Excel;
using PdfSharpCore.Pdf.IO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using PdfSharpCore.Pdf;
using PDFSlicer.Models;

namespace PDFSlicer
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnBrowseExcel_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls"
            };
            if (dialog.ShowDialog() == true)
            {
                txtExcelPath.Text = dialog.FileName;
            }
        }

        private void BtnAddPdf_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "PDF Files|*.pdf",
                Multiselect = true
            };
            if (dialog.ShowDialog() == true)
            {
                foreach (var file in dialog.FileNames)
                {
                    lstPdfFiles.Items.Add(file);
                }
            }
        }

        private void BtnClearPdf_Click(object sender, RoutedEventArgs e)
        {
            lstPdfFiles.Items.Clear();
        }

        private void BtnProcess_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(txtExcelPath.Text) || lstPdfFiles.Items.Count == 0)
            {
                MessageBox.Show("Please select Excel file and PDF files");
                return;
            }

            if (!int.TryParse(txtStartRow.Text, out int startRow))
            {
                MessageBox.Show("Invalid start row number");
                return;
            }

            try
            {
                var excelData = ExcelParser.Parse(txtExcelPath.Text, startRow);
                progressBar.Maximum = lstPdfFiles.Items.Count;
                progressBar.Value = 0;

                foreach (var pdfPath in lstPdfFiles.Items)
                {
                    ProcessPdf(pdfPath.ToString(), excelData);
                    progressBar.Value++;
                }

                txtLog.AppendText("Processing completed successfully!\n");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
                txtLog.AppendText($"ERROR: {ex.Message}\n");
            }
        }

        private void ProcessPdf(string pdfPath, Dictionary<string, ExcelRecord> excelData)
        {
            var pdfInfo = PdfProcessor.ParsePdfName(Path.GetFileNameWithoutExtension(pdfPath));
            var outputDirectory = Path.Combine(Path.GetDirectoryName(pdfPath), "Output");
            Directory.CreateDirectory(outputDirectory);

            using (var document = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Import))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    var page = document.Pages[i];
                    var record = PdfProcessor.FindMatchingRecord(pdfInfo, excelData, i + 1);

                    if (record != null)
                    {
                        var newFileName = PdfProcessor.GenerateFileName(Path.GetFileName(pdfPath), record);
                        PdfProcessor.SaveSinglePage(page, outputDirectory, newFileName);
                        txtLog.AppendText($"Created: {newFileName}\n");
                    }
                    else
                    {
                        txtLog.AppendText($"No matching record found for page {i + 1}\n");
                    }
                }
            }
        }
    }

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
                        // Логирование ошибки парсинга строки
                        Console.WriteLine($"Ошибка обработки строки {row}: {ex.Message}");
                    }
                }
            }

            return records;
        }

        private static string FormatDate(string date)
        {
            if (DateTime.TryParse(date, out DateTime result))
            {
                return result.ToString("dd.MM.yy");
            }

            if (date.Length == 10) // dd.MM.yyyy
            {
                return date.Substring(0, 5) + date.Substring(8, 2); // -> dd.MM.yy
            }
            else if (date.Length == 8) // dd.MM.yy
            {
                return date;
            }
            else if (date.Length == 5) // dd.MM
            {
                return date + DateTime.Now.ToString(".yy"); // -> dd.MM.yy
            }
            else
            {
                return "01.01.00"; // По умолчанию
            }
        }
    }

    

    public static class PdfProcessor
    {
        public static PdfInfo ParsePdfName(string fileName)
        {
            var parts = fileName.Split('_');
            return new PdfInfo
            {
                CertificateRange = parts.Length > 0 ? parts[0] : "Unknown",
                IssueDate = parts.Length > 1 ? FormatDate(parts[1]) : "01.01.00",
                ProgramName = parts.Length > 2 ? ShortenProgramName(parts[2]) : "Program",
                Hours = parts.Length > 3 ? parts[3] : "Hours"
            };
        }

        public static ExcelRecord FindMatchingRecord(PdfInfo pdfInfo, Dictionary<string, ExcelRecord> excelData, int pageNumber)
        {
            // Simple matching logic based on page order
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
            using (var document = new PdfDocument())
            {
                document.AddPage(page);
                document.Save(Path.Combine(directory, fileName));
            }
        }

        private static string ShortenProgramName(string programName)
        {
            if (string.IsNullOrWhiteSpace(programName))
                return "Program";

            var words = programName.Split(' ');
            var shortened = new StringBuilder();

            foreach (var word in words)
            {
                if (word.Length > 4)
                    shortened.Append(word.Substring(0, 4));
                else
                    shortened.Append(word);

                shortened.Append('_');
            }

            return shortened.ToString().TrimEnd('_');
        }

        private static string RemoveInvalidChars(string input)
        {
            var invalidChars = System.IO.Path.GetInvalidFileNameChars();
            return new string(input.Where(c => !invalidChars.Contains(c)).ToArray());
        }

        private static string FormatDate(string date)
        {
            if (date.Length == 8) // dd.MM.yy
                return date;
            if (date.Length == 5) // dd.MM
                return $"{date}.{DateTime.Now:yy}";
            return "01.01.00";
        }
    }
}
