using System;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using iTextSharp.text;
using iTextSharp.text.pdf;

public class TableGenerator
{
    public void GenerateTable(Document doc, string excelFilePath, string studentRollNo, string studentName)
    {
        // Define fonts
        var boldBlackFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.Black);
        var regularBlackFont = FontFactory.GetFont(FontFactory.HELVETICA, 10, BaseColor.Black);

        // Create table with specified number of columns
        PdfPTable table = new PdfPTable(10)
        {
            WidthPercentage = 100,
            SpacingBefore = 10f,
            SpacingAfter = 10f
        };
        table.SetWidths(new float[] { 3f, 2f, 2f, 2f, 2f, 2f, 2f, 2f, 2f, 2f });

        // Add main headers
        AddCell(table, "Scholastic", boldBlackFont, 1, 1);
        AddCell(table, "Semester - I", boldBlackFont, 2, 1);
        AddCell(table, "Semester - II", boldBlackFont, 2, 1);
        AddCell(table, "Total", boldBlackFont, 1, 1);
        AddCell(table, "Overall", boldBlackFont, 4, 1);

        // Add sub-headers
        AddCell(table, "Subject", boldBlackFont);
        AddCell(table, "20", boldBlackFont);
        AddCell(table, "80", boldBlackFont);
        AddCell(table, "20", boldBlackFont);
        AddCell(table, "80", boldBlackFont);
        AddCell(table, "100", boldBlackFont);
        AddCell(table, "SEM-I\n50%", boldBlackFont);
        AddCell(table, "SEM-II\n50%", boldBlackFont);
        AddCell(table, "100%", boldBlackFont);
        AddCell(table, "Grade", boldBlackFont);

        //Fetching data from excel
        using (FileStream file = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            // Loop through each subject sheet
            for (int sheetIndex = 1; sheetIndex < workbook.NumberOfSheets; sheetIndex++)
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                string subjectName = sheet.SheetName;

                // Loop through each student row
                for (int rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    IRow row = sheet.GetRow(rowIndex);
                    if (row == null) continue;

                    string rollNo = row.GetCell(0)?.ToString();
                    string name = row.GetCell(1)?.ToString();

                    if (rollNo == studentRollNo && name == studentName)
                    {
                        string labMarksSem1 = row.GetCell(2).ToString();
                        string theoryMarksSem1 = row.GetCell(3).ToString();
                        string labMarksSem2 = row.GetCell(4).ToString();
                        string theoryMarksSem2 = row.GetCell(5).ToString();

                        int labMarksSem1Int = int.TryParse(labMarksSem1, out int temp1) ? temp1 : 0;
                        int theoryMarksSem1Int = int.TryParse(theoryMarksSem1, out int temp2) ? temp2 : 0;
                        int labMarksSem2Int = int.TryParse(labMarksSem2, out int temp3) ? temp3 : 0;
                        int theoryMarksSem2Int = int.TryParse(theoryMarksSem2, out int temp4) ? temp4 : 0;

                        int totalMarksSem1 = labMarksSem1Int + theoryMarksSem1Int;
                        int totalMarksSem2 = labMarksSem2Int + theoryMarksSem2Int;
                        int overallMarks = (totalMarksSem1 + totalMarksSem2)/2;

                        AddCell(table, subjectName, regularBlackFont);
                        AddCell(table, labMarksSem1, regularBlackFont);
                        AddCell(table, theoryMarksSem1, regularBlackFont);
                        AddCell(table, labMarksSem2, regularBlackFont);
                        AddCell(table, theoryMarksSem2, regularBlackFont);
                        AddCell(table, overallMarks.ToString(), regularBlackFont);
                        AddCell(table, totalMarksSem1.ToString(), regularBlackFont);
                        AddCell(table, totalMarksSem2.ToString(), regularBlackFont);
                        AddCell(table, overallMarks.ToString(), regularBlackFont);
                        AddCell(table, GetGrade(overallMarks), regularBlackFont);
                    }
                }
            }
        }

        doc.Add(table);
    }

    private static void AddCell(PdfPTable table, string text, Font font, int colspan = 1, int rowspan = 1)
    {
        PdfPCell cell = new PdfPCell(new Phrase(text, font))
        {
            Colspan = colspan,
            Rowspan = rowspan,
            HorizontalAlignment = Element.ALIGN_CENTER,
            VerticalAlignment = Element.ALIGN_MIDDLE,
            Padding = 5,
            BorderWidth = 1
        };
        table.AddCell(cell);
    }

    private static string GetGrade(int marks)
    {
        if (marks >= 90) return "A+";
        if (marks >= 80) return "A";
        if (marks >= 70) return "B+";
        if (marks >= 60) return "B";
        if (marks >= 50) return "C";
        return "F";
    }
}
