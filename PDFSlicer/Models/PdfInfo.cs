namespace PDFSlicer.Models;

public class PdfInfo
{
    /// <summary>
    /// Диапазон серии сертификата
    /// </summary>
    public string CertificateRange { get; set; }
    
    /// <summary>
    /// Дата выдачи удостоверения
    /// </summary>
    public string IssueDate { get; set; }
    
    /// <summary>
    /// Наименование программы
    /// </summary>
    public string ProgramName { get; set; }
    
    /// <summary>
    /// Количество часов
    /// </summary>
    public string Hours { get; set; }
}