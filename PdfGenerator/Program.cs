using System;
using System.Collections.Generic;

class Program
{
    static void Main(string[] args)
    {
        string excelFilePath = "D:\\Downloads\\schoolManagement.xlsx";
        string outputDirectory = "C:\\Users\\chidv\\OneDrive\\Documents";

        IExcelReader excelReader = new ExcelReader();
        IReportGenerator pdfGenerator = new ReportGenerator();

        IEnumerable<Student> students = excelReader.ReadStudents(excelFilePath);
        foreach (var student in students)
        {
            pdfGenerator.GenerateReport(student, outputDirectory);
        }
    }
}
