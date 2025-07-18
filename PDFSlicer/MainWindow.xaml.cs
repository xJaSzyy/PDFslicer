using PdfSharpCore.Pdf.IO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using PDFSlicer.Models;
using PDFSlicer.Processing;

namespace PDFSlicer;

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

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        foreach (var file in dialog.FileNames)
        {
            lstPdfFiles.Items.Add(file);
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

        if (!int.TryParse(txtStartRow.Text, out var startRow))
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
        var outputDirectory = Path.Combine(Path.GetDirectoryName(pdfPath)!, "Output");
        Directory.CreateDirectory(outputDirectory);

        using var document = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Import);
        for (var i = 0; i < document.PageCount; i++)
        {
            var page = document.Pages[i];
            var record = PdfProcessor.FindMatchingRecord(excelData, i + 1);

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