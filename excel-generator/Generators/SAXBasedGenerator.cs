using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelGenerator.Generators;

public class SAXBasedGenerator
{
    public void LargeExport(string filename)
    {
        using (var stream = new MemoryStream())
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                //this list of attributes will be used when writing a start element
                List<OpenXmlAttribute> attributes;
                OpenXmlWriter writer;
                var namespaceUri = "";
                document.AddWorkbookPart();
                if (document.WorkbookPart != null)
                {
                    WorksheetPart workSheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();

                    writer = OpenXmlWriter.Create(workSheetPart);
                    writer.WriteStartElement(new Worksheet());
                    writer.WriteStartElement(new SheetData());

                    for (int rowNum = 1; rowNum <= 10; rowNum++)
                    {
                        //create a new list of attributes
                        attributes = new List<OpenXmlAttribute>();
                        // add the row index attribute to the list
                        attributes.Add(new OpenXmlAttribute("r", namespaceUri, rowNum.ToString()));
                        //write the row start element with the row index attribute
                        writer.WriteStartElement(new Row(), attributes);

                        for (int columnNum = 1; columnNum <= 12; columnNum++)
                        {
                            //reset the list of attributes
                            attributes = new List<OpenXmlAttribute>();
                            // add data type attribute - in this case inline string (you might want to look at the shared strings table)
                            attributes.Add(new OpenXmlAttribute("t", namespaceUri, "str"));
                            //add the cell reference attribute
                            attributes.Add(new OpenXmlAttribute("r", namespaceUri, string.Format("{0}{1}", GetColumnName(columnNum), rowNum)));

                            //write the cell start element with the type and reference attributes
                            writer.WriteStartElement(new Cell(), attributes);
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

                    document.Save();

                    document.Close();

                    stream.Seek(0, SeekOrigin.Begin);

                    var documentBytes = stream.ToArray();

                    var base64File = Convert.ToBase64String(documentBytes);

                    var bytes = Convert.FromBase64String(base64File);

                    using (var ms = new MemoryStream(bytes))
                    {
                        using (SpreadsheetDocument testDocument = SpreadsheetDocument.Open(ms, false))
                        {
                            testDocument.SaveAs(filename);
                            testDocument.Close();
                        }
                    }
                }
            }
        }
    }



    //A simple helper to get the column name from the column index. This is not well tested!
    private string GetColumnName(int columnIndex)
    {
        int dividend = columnIndex;
        string columnName = String.Empty;
        int modifier;

        while (dividend > 0)
        {
            modifier = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modifier).ToString() + columnName;
            dividend = (int)((dividend - modifier) / 26);
        }

        return columnName;
    }
}
