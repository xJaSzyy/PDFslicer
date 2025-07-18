namespace PDFSlicer.Models;

public class ExcelRecord
{
    /// <summary>
    /// 
    /// </summary>
    public int RowNumber { get; set; }
    
    /// <summary>
    /// Регистрационный номер
    /// </summary>
    public string RegistrationNumber { get; set; } 
    
    /// <summary>
    /// Наименование документа
    /// </summary>
    public string DocumentName { get; set; }         
    
    /// <summary>
    /// Серия и номер удостоверения
    /// </summary>
    public string CertificateNumber { get; set; }    
    
    /// <summary>
    /// Дата выдачи удостоверения
    /// </summary>
    public string IssueDate { get; set; }

    /// <summary>
    /// ФИО
    /// </summary>
    public string FullName { get; set; }
    
    /// <summary>
    /// Наименование программы
    /// </summary>
    public string ProgramName { get; set; }
    
    /// <summary>
    /// Количество часов
    /// </summary>
    public string Hours { get; set; }               
}