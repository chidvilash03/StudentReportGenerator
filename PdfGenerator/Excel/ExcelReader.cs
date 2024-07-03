using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

public class ExcelReader : IExcelReader
{
    public IEnumerable<Student> ReadStudents(string filePath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var students = new List<Student>();
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                var student = new Student
                {
                    RollNumber = worksheet.Cells[row, 1].Text,
                    Name = worksheet.Cells[row, 2].Text,
                    Class = worksheet.Cells[row, 3].Text,
                };
                students.Add(student);
            }
        }
        return students;
    }
}
