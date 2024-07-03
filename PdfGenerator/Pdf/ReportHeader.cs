using iTextSharp.text.pdf;
using iTextSharp.text;
using iTextSharp.text.pdf.draw;


public class ReportHeader
{
    public static void CreateHeader(Document doc,Student student)
    {
        // Add college logo
        var logo = Image.GetInstance("D:\\Downloads\\Siemens-logo.png");
        logo.ScaleToFit(80f, 80f);
        logo.Alignment = Image.ALIGN_LEFT;

        // Define fonts
        var collegeNameFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 16, BaseColor.Black);
        var collegeInfoFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, BaseColor.Black);
        var linkFont = FontFactory.GetFont(FontFactory.HELVETICA, 12, BaseColor.Blue);
        var NormalFont = FontFactory.GetFont(FontFactory.HELVETICA, 12, BaseColor.Black);
        var boldBlackFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, BaseColor.Black);

        PdfPTable headerTable = new PdfPTable(2)
        {
            WidthPercentage = 100
        };
        headerTable.SetWidths(new float[] { 1f, 3f });

        PdfPCell logoCell = new PdfPCell(logo)
        {
            Border = Rectangle.NO_BORDER,
            HorizontalAlignment = Element.ALIGN_LEFT,
            VerticalAlignment = Element.ALIGN_MIDDLE,
            Padding = 5
        };
        headerTable.AddCell(logoCell);

        PdfPCell detailsCell = new PdfPCell
        {
            Border = Rectangle.NO_BORDER,
            HorizontalAlignment = Element.ALIGN_RIGHT,
            VerticalAlignment = Element.ALIGN_MIDDLE,
            Padding = 10
        };

        // Create a nested table for the college details
        PdfPTable detailsTable = new PdfPTable(1)
        {
            WidthPercentage = 100
        };

        var collegeName = new Paragraph("SRM Institute of Science and Technology", collegeNameFont)
        {
            Alignment = Element.ALIGN_CENTER
        };
        detailsTable.AddCell(new PdfPCell(collegeName)
        {
            Border = Rectangle.NO_BORDER,
            HorizontalAlignment = Element.ALIGN_CENTER
        });

        var collegeAddress = new Paragraph("Address\nSRM University\nPh: 5465646565 / 5454546546", collegeInfoFont)
        {
            Alignment = Element.ALIGN_CENTER
        };
        detailsTable.AddCell(new PdfPCell(collegeAddress)
        {
            Border = Rectangle.NO_BORDER,
            HorizontalAlignment = Element.ALIGN_CENTER
        });

        // Create a table for email and website to be on the same line
        PdfPTable contactTable = new PdfPTable(2)
        {
            WidthPercentage = 100
        };
        contactTable.SetWidths(new float[] { 1f, 1f });

        // Email cell with mixed color text
        Phrase emailPhrase = new Phrase();
        emailPhrase.Add(new Chunk("Email: ", boldBlackFont));
        emailPhrase.Add(new Chunk("official@srm.com", linkFont));

        PdfPCell emailCell = new PdfPCell(emailPhrase)
        {
            Border = Rectangle.NO_BORDER,
            HorizontalAlignment = Element.ALIGN_RIGHT
        };
        contactTable.AddCell(emailCell);

        // Website cell with mixed color text
        Phrase websitePhrase = new Phrase();
        websitePhrase.Add(new Chunk("Visit us: ", boldBlackFont));
        websitePhrase.Add(new Chunk("www.srm.edu.in", linkFont));

        PdfPCell websiteCell = new PdfPCell(websitePhrase)
        {
            Border = Rectangle.NO_BORDER,
            HorizontalAlignment = Element.ALIGN_LEFT
        };
        contactTable.AddCell(websiteCell);

        detailsTable.AddCell(new PdfPCell(contactTable)
        {
            Border = Rectangle.NO_BORDER,
            HorizontalAlignment = Element.ALIGN_CENTER
        });

        detailsCell.AddElement(detailsTable);
        headerTable.AddCell(detailsCell);

        doc.Add(headerTable);

        // Add report card title in a single cell
        PdfPTable reportCardTable = new PdfPTable(1)
        {
            WidthPercentage = 100,
            SpacingBefore = 10f
        };

        PdfPCell reportCardCell = new PdfPCell(new Phrase("REPORT CARD", collegeNameFont))
        {
            Border = Rectangle.NO_BORDER,
            HorizontalAlignment = Element.ALIGN_CENTER,
            PaddingBottom = 10f
        };
        reportCardTable.AddCell(reportCardCell);

        doc.Add(reportCardTable);

        // Add student details table
        PdfPTable table = new PdfPTable(3)
        {
            WidthPercentage = 100,
            SpacingBefore = 10f
        };
        table.SetWidths(new float[] { 1f, 1f, 1f });

        // First row: Class and Academic Session
        PdfPCell leftCell1 = new PdfPCell(new Phrase("Class: " + student.Class, boldBlackFont))
        {
            Border = Rectangle.NO_BORDER,
            HorizontalAlignment = Element.ALIGN_LEFT,
            PaddingBottom = 5f // Decrease padding to reduce spacing
        };

        PdfPCell middleCell1 = new PdfPCell(new Phrase("Academic Session: 2024-25", boldBlackFont))
        {
            Border = Rectangle.NO_BORDER,
            HorizontalAlignment = Element.ALIGN_CENTER,
            PaddingBottom = 5f // Decrease padding to reduce spacing
        };

        PdfPCell rightCell1 = new PdfPCell(new Phrase("", boldBlackFont))
        {
            Border = Rectangle.NO_BORDER,
            HorizontalAlignment = Element.ALIGN_RIGHT,
            PaddingBottom = 5f // Decrease padding to reduce spacing
        };

        table.AddCell(leftCell1);
        table.AddCell(middleCell1);
        table.AddCell(rightCell1);

        doc.Add(table);

        // Second row: Name and Roll No
        PdfPTable secondRowTable = new PdfPTable(3)
        {
            WidthPercentage = 100,
            SpacingBefore = 10f
        };
        secondRowTable.SetWidths(new float[] { 1f, 1f, 1f });

        PdfPCell leftCell2 = new PdfPCell(new Phrase("Name : " + student.Name, boldBlackFont))
        {
            Border = Rectangle.NO_BORDER,
            HorizontalAlignment = Element.ALIGN_LEFT,
            PaddingBottom = 5f // Decrease padding to reduce spacing
        };

        PdfPCell middleCell2 = new PdfPCell(new Phrase("", boldBlackFont))
        {
            Border = Rectangle.NO_BORDER,
            HorizontalAlignment = Element.ALIGN_CENTER,
            PaddingBottom = 5f // Decrease padding to reduce spacing
        };

        PdfPCell rightCell2 = new PdfPCell(new Phrase("Roll No: " + student.RollNumber, boldBlackFont))
        {
            Border = Rectangle.NO_BORDER,
            HorizontalAlignment = Element.ALIGN_RIGHT,
            PaddingBottom = 5f // Decrease padding to reduce spacing
        };

        secondRowTable.AddCell(leftCell2);
        secondRowTable.AddCell(middleCell2);
        secondRowTable.AddCell(rightCell2);

        doc.Add(secondRowTable);
        var boldLine = new LineSeparator(1f, 100f, BaseColor.Black, Element.ALIGN_CENTER, -2);
        doc.Add(new Chunk(boldLine));
    }
}
