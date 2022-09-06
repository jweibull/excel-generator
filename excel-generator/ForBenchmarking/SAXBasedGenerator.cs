using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelGenerator.ForBenchmarking;

public class SAXBasedGenerator
{
    public void CreatePackage(string filename)
    {
        using (var stream = new MemoryStream())
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                //this list of attributes will be used when writing a start element
                OpenXmlWriter writer;
                document.AddWorkbookPart();
                if (document.WorkbookPart != null)
                {
                    WorksheetPart workSheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();

                    writer = OpenXmlWriter.Create(workSheetPart);
                    writer.WriteStartElement(new Worksheet());
                    writer.WriteStartElement(new SheetData());

                    for (int rowNum = 1; rowNum <= 100000; rowNum++)
                    {
                        //write the row start element with the row index attribute
                        writer.WriteStartElement(new Row() { RowIndex = (uint)rowNum });

                        for (int columnNum = 1; columnNum <= 50; columnNum++)
                        {
                            //write the cell start element with the type and reference attributes
                            writer.WriteStartElement(new Cell() { CellReference = string.Format("{0}{1}", GetColumnName(columnNum), rowNum), DataType = CellValues.String });
                            //write the cell value
                            writer.WriteElement(new CellValue(string.Format("This is Row {0}, Cell {1}", rowNum, columnNum)));
                            // write the end cell element
                            writer.WriteEndElement();
                        }

                        // write the end row element
                        writer.WriteEndElement();
                    }

                    // write the end SheetData element
                    writer.WriteEndElement();
                    // write the end Worksheet element
                    writer.WriteEndElement();
                    writer.Close();

                    writer = OpenXmlWriter.Create(document.WorkbookPart);
                    writer.WriteStartElement(new Workbook());
                    writer.WriteStartElement(new Sheets());

                    writer.WriteElement(new Sheet()
                    {
                        Name = "Large Sheet",
                        SheetId = 1,
                        Id = document.WorkbookPart.GetIdOfPart(workSheetPart)
                    });

                    // End Sheets
                    writer.WriteEndElement();
                    // End Workbook
                    writer.WriteEndElement();

                    writer.Close();

                    document.SaveAs(filename);

                    //document.Save();

                    //document.Close();

                    //stream.Seek(0, SeekOrigin.Begin);

                    //var documentBytes = stream.ToArray();

                    //var base64File = Convert.ToBase64String(documentBytes);

                    //var bytes = Convert.FromBase64String(base64File);

                    //using (var ms = new MemoryStream(bytes))
                    //{
                    //    using (SpreadsheetDocument testDocument = SpreadsheetDocument.Open(ms, false))
                    //    {
                    //        testDocument.SaveAs(filename);
                    //        testDocument.Close();
                    //    }
                    //}
                }
            }
        }
    }



    //A simple helper to get the column name from the column index. This is not well tested!
    private string GetColumnName(int columnIndex)
    {
        int dividend = columnIndex;
        string columnName = string.Empty;
        int modifier;

        while (dividend > 0)
        {
            modifier = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modifier).ToString() + columnName;
            dividend = (dividend - modifier) / 26;
        }

        return columnName;
    }
}
