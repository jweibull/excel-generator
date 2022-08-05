using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using X15ac = DocumentFormat.OpenXml.Office2013.ExcelAc;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using A = DocumentFormat.OpenXml.Drawing;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;

namespace ExcelGenerator.Generators;

public class MinimalBoilerPlateWithSax
{
    public void CreatePackage(string filename)
    {
        using (var stream = new MemoryStream())
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                // TestData
                var data = new string[4] { "Header 1", "A", "B", "B" };
                var fonts = new string[2] { "Calibri", "Calibri Light" };
                var partId = 1;
                string partIdString = string.Empty;
                document.AddWorkbookPart();
                if (document.WorkbookPart != null)
                {
                    // Generate all Shared Strings that will be used in all the sheets
                    var sharedStringsToIndex = GenerateSharedStringsTable(document.WorkbookPart, data);

                    // Generate all Styles needed on every sheet in this workbook
                    partIdString = "rId" + partId++;
                    GenerateStylePart(document.WorkbookPart, partIdString, fonts);

                    // Generate a single sheet 
                    partIdString = "rId" + partId++;
                    GenerateWorkSheetData(document.WorkbookPart, partIdString, data, sharedStringsToIndex);

                    // Create the worksheet and sheets list to end the package
                    using (var writer = OpenXmlWriter.Create(document.WorkbookPart))
                    {
                        writer.WriteStartElement(new Workbook());
                        writer.WriteStartElement(new Sheets());

                        writer.WriteElement(new Sheet()
                        {
                            Name = "Shared string table",
                            SheetId = 1,
                            Id = partIdString
                        });

                        // End Sheets
                        writer.WriteEndElement();
                        // End Workbook
                        writer.WriteEndElement();

                        writer.Close();
                    }
                    //document.Save();

                    document.SaveAs(filename);

                    document.Close();
                }
            }
        }
    }

    private Dictionary<string,string> GenerateSharedStringsTable(WorkbookPart workbookPart, string[] sharedStrings)
    {
        // Run this for all strings in the workbook
        // string[] sharedStrings must contain all the strings in the project

        var sharedStringsToIndex = new Dictionary<string, string>();
        var totalCount = 0;
        totalCount += AddToSharedStringDictionary(sharedStringsToIndex, sharedStrings);
        var uniqueCount = sharedStringsToIndex.Count;

        List<OpenXmlAttribute> attributes;
        
        SharedStringTablePart sharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>();

        using (var writer = OpenXmlWriter.Create(sharedStringTablePart))
        {
            var namespaceUri = string.Empty;

            //write attributes
            // Change this based on real data count
            attributes = new List<OpenXmlAttribute>();
            attributes.Add(new OpenXmlAttribute("count", namespaceUri, totalCount.ToString()));
            attributes.Add(new OpenXmlAttribute("uniqueCount", namespaceUri, uniqueCount.ToString()));
            writer.WriteStartElement(new SharedStringTable(), attributes);

            foreach (var key in sharedStringsToIndex.Keys)
            {
                //write the row start element with the row index attribute
                writer.WriteStartElement(new SharedStringItem());

                //write the text value
                writer.WriteElement(new Text(key));

                // write the end sharedItem element
                writer.WriteEndElement();
            }

            // write the end SharedStringTable element
            writer.WriteEndElement();

            writer.Close();
        }
        return sharedStringsToIndex;
    }

    private int AddToSharedStringDictionary(Dictionary<string, string> sharedStringToIndex, string[] sharedStrings)
    {
        var count = 0;
        foreach (var item in sharedStrings)
        {
            if (sharedStringToIndex.ContainsKey(item))
            {
                count++;
            }
            else
            {
                count++;
                sharedStringToIndex.Add(item, sharedStringToIndex.Count().ToString());
            }
        }
        return count;
    }

    private void GenerateWorkSheetData(WorkbookPart workbookPart, string sheetPartId, string[] data, Dictionary<string, string> sharedStringsToIndex)
    {
        // Actual Cell Values from string table
        WorksheetPart workSheetPart = workbookPart.AddNewPart<WorksheetPart>(sheetPartId);

        var namespaceUri = string.Empty;
        List<OpenXmlAttribute> attributes;

        using (var writer = OpenXmlWriter.Create(workSheetPart))
        {
            writer.WriteStartElement(new Worksheet());
            writer.WriteStartElement(new SheetData());

            for (int rowNum = 1; rowNum <= data.Length; rowNum++)
            {
                //create a new list of attributes
                attributes = new List<OpenXmlAttribute>();
                // add the row index attribute to the list
                attributes.Add(new OpenXmlAttribute("r", namespaceUri, rowNum.ToString()));
                //write the row start element with the row index attribute
                writer.WriteStartElement(new Row(), attributes);

                for (int columnNum = 1; columnNum <= 1; columnNum++)
                {
                    //reset the list of attributes
                    attributes = new List<OpenXmlAttribute>();
                    // add data type attribute - in this case inline string (you might want to look at the shared strings table)
                    attributes.Add(new OpenXmlAttribute("t", namespaceUri, "s"));
                    //add the cell reference attribute
                    attributes.Add(new OpenXmlAttribute("r", namespaceUri, string.Format("{0}{1}", GetColumnName(columnNum), rowNum)));

                    //write the cell start element with the type and reference attributes
                    writer.WriteStartElement(new Cell(), attributes);
                    //write the cell value
                    writer.WriteElement(new CellValue(sharedStringsToIndex[data[rowNum - 1]]));

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
        }
    }

    private void GenerateStylePart(WorkbookPart workbookPart, string sheetPartId, string[] fonts)
    {
        //Hardcoded props
        var fontSize = 11;
        var fontFamily = 2; // Calibri family?
        var theme = 1;
        
        WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>(sheetPartId);

        var namespaceUri = string.Empty;
        List<OpenXmlAttribute> attributes;

        using (var writer = OpenXmlWriter.Create(workbookStylesPart))
        {
            

            writer.WriteStartElement(new Stylesheet());

            #region Fonts
            //write the fonts sections
            //<Fonts>
            //  <Font>...props...</Font>
            //</Fonts>
            attributes = new List<OpenXmlAttribute>();
            attributes.Add(new OpenXmlAttribute("count", namespaceUri, fonts.Length.ToString()));
            writer.WriteStartElement(new Fonts(), attributes);

            foreach (var font in fonts)
            {
                writer.WriteStartElement(new Font());
                
                writer.WriteElement(new FontSize() { Val = fontSize });
                writer.WriteElement(new Color() { Theme = (UInt32)theme });
                writer.WriteElement(new FontName() { Val = font });
                writer.WriteElement(new FontFamily() { Val = fontFamily });
                writer.WriteElement(new FontScheme() { Val = FontSchemeValues.Major });

                //Close the single Font Tag
                writer.WriteEndElement();
            }

            // End Fonts section
            writer.WriteEndElement();
            #endregion

            //End styleSsheet
            writer.WriteEndElement();
            writer.Close();
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
