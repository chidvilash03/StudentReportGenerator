using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;

public class ReportGenerator : IReportGenerator
{
    public void GenerateReport(Student student, string outputDirectory,string excelFilePath)
    {
        var fileName = Path.Combine(outputDirectory, $"{student.RollNumber}_Report.pdf");
        using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            //Open the document
            var doc = new Document(PageSize.A4, 50, 50, 50, 50);
            var writer = PdfWriter.GetInstance(doc, fs);
            doc.Open();

            //Generate header for each report
            ReportHeader.CreateHeader(doc, student);

            var tableGenerator = new TableGenerator();
            tableGenerator.GenerateTable(doc, excelFilePath,student.RollNumber,student.Name);

            // Close the document
            doc.Close();
            writer.Close();
        }
    }
}
