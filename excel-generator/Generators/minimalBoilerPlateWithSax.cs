using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using x14 = DocumentFormat.OpenXml.Office2010.Excel;
using x15 = DocumentFormat.OpenXml.Office2013.Excel;

namespace ExcelGenerator.Generators;

public class MinimalBoilerPlateWithSax
{
    public Dictionary<string, string> SharedStringsToIndex { get; set; } = new Dictionary<string, string>();
    public int sharedStringsCount { get; set; } = 0;
    public int sharedStringsUniqueCount { get; set; } = 0;

    public void CreatePackage(string filename)
    {
        using (var stream = new MemoryStream())
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                // TestData
                var data = new string[4] { "Header", "A", "B", "B" };
                var fonts = new string[2] { "Calibri", "Calibri Light" };
                var partId = 1;
                string sharedTableId = string.Empty;
                string stylesPartId = string.Empty;
                string sheetPartId = string.Empty;

                document.AddWorkbookPart();
                
                if (document.WorkbookPart != null)
                {
                    // Generate all Shared Strings that will be used in all the sheets
                    sharedStringsCount = 0;
                    AddToSharedStringDictionary(data);
                    sharedStringsUniqueCount = SharedStringsToIndex.Count;

                    // Generate a single sheet 
                    sheetPartId = "rId" + partId++;
                    stylesPartId = "rId" + partId++;
                    sharedTableId = "rId" + partId++;

                    // Generate all Styles needed on every sheet in this workbook
                    WorkbookStylesPart workbookStylesPart = document.WorkbookPart.AddNewPart<WorkbookStylesPart>(stylesPartId);
                    WorksheetPart workSheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>(sheetPartId);
                    TableDefinitionPart sheetTablesPart = workSheetPart.AddNewPart<TableDefinitionPart>(sheetPartId);
                    SharedStringTablePart sharedStringTablePart = document.WorkbookPart.AddNewPart<SharedStringTablePart>(sharedTableId);

                    GenerateStylePart(workbookStylesPart, stylesPartId, fonts);
                    GenerateWorkSheetData(workSheetPart, data, sheetPartId);
                    GenerateTableParts(sheetTablesPart, sheetPartId);
                    GenerateSharedStringsTable(sharedStringTablePart, data, sharedTableId);

                    // Create the worksheet and sheets list to end the package
                    using (var writer = OpenXmlWriter.Create(document.WorkbookPart))
                    {
                        writer.WriteStartElement(new Workbook());
                        writer.WriteStartElement(new Sheets());

                        writer.WriteElement(new Sheet()
                        {
                            Name = "Planilha1",
                            SheetId = 1,
                            Id = sheetPartId
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

    private void GenerateSharedStringsTable(SharedStringTablePart sharedStringTablePart, string[] sharedStrings, string sharedTableId)
    {
        // Run this for all strings in the workbook
        // string[] sharedStrings must contain all the strings in the project

        using (var writer = OpenXmlWriter.Create(sharedStringTablePart))
        {
            // Change this based on real data count
            writer.WriteStartElement(new SharedStringTable() { Count = (UInt32)sharedStringsCount, UniqueCount = (UInt32)sharedStringsUniqueCount });

            foreach (var key in SharedStringsToIndex.Keys)
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
    }

    private void AddToSharedStringDictionary(string[] sharedStrings)
    {
        var count = 0;
        foreach (var item in sharedStrings)
        {
            if (this.SharedStringsToIndex.ContainsKey(item))
            {
                count++;
            }
            else
            {
                count++;
                SharedStringsToIndex.Add(item, SharedStringsToIndex.Count().ToString());
            }
        }
        sharedStringsCount += count;
    }

    private void GenerateWorkSheetData(WorksheetPart workSheetPart, string[] data, string sheetPartId)
    {
        // Actual Cell Values from string table
        using (var writer = OpenXmlWriter.Create(workSheetPart))
        {
            writer.WriteStartElement(new Worksheet());

            writer.WriteStartElement(new Columns() { });
            writer.WriteElement(new Column() { Min = 1, Max = 1, Width=12, CustomWidth=true });
            writer.WriteEndElement();

            writer.WriteStartElement(new SheetData());

            for (int rowNum = 1; rowNum <= data.Length; rowNum++)
            {
                //write the row start element with the row index attribute
                writer.WriteStartElement(new Row() { RowIndex = (UInt32)rowNum });

                for (int columnNum = 1; columnNum <= 1; columnNum++)
                {
                    //write the cell start element with the type and reference attributes
                    writer.WriteStartElement(new Cell() { CellReference = string.Format("{0}{1}", GetColumnName(columnNum), rowNum), DataType = CellValues.SharedString });
                    //write the cell value
                    writer.WriteElement(new CellValue(SharedStringsToIndex[data[rowNum - 1]]));

                    // write the end cell element
                    writer.WriteEndElement();
                }

                // write the end row element
                writer.WriteEndElement();
            }

            // write the end SheetData element
            writer.WriteEndElement();

            writer.WriteStartElement(new TableParts() { Count = 1 });
            writer.WriteElement(new TablePart() { Id = sheetPartId });
            writer.WriteEndElement();

            // write the end Worksheet element
            writer.WriteEndElement();
            writer.Close();
        }
    }


    // TODO create dictionaries to link fonts and fills etc to the CellFormatStyles element inside this method.
    // Everything is linked by a string id that is in fact the index of the array of style element. Ex the font with id "2"
    // will be the third font added in fonts section, while the font with id "0" will be the first you added.
    // Same goes for borders, fills, etc.
    private void GenerateStylePart(WorkbookStylesPart workbookStylesPart, string stylesPartId, string[] fonts)
    {
        using (var writer = OpenXmlWriter.Create(workbookStylesPart))
        {
            writer.WriteStartElement(new Stylesheet());

            #region Fonts

            //Hardcoded props
            var fontSize = 11;
            var fontFamily = 2; // Calibri family?
            var theme = 1;

            //write the fonts sections
            //<Fonts>
            //  <Font>...props...</Font>
            //</Fonts>
            writer.WriteStartElement(new Fonts() { Count = (UInt32)fonts.Length });

            foreach (var font in fonts)
            {
                writer.WriteStartElement(new Font());

                writer.WriteElement(new FontSize() { Val = fontSize });
                writer.WriteElement(new Color() { Theme = (UInt32)theme });
                writer.WriteElement(new FontName() { Val = font });
                writer.WriteElement(new FontFamily() { Val = fontFamily });

                // Why is this here??? What's the diference between major and minor fonts
                writer.WriteElement(new FontScheme() { Val = FontSchemeValues.Major });

                //Close the single Font Tag
                writer.WriteEndElement();

                writer.WriteStartElement(new Font());

                writer.WriteElement(new FontSize() { Val = fontSize });
                writer.WriteElement(new Color() { Theme = (UInt32)theme });
                writer.WriteElement(new FontName() { Val = font });
                writer.WriteElement(new FontFamily() { Val = fontFamily });

                // Why is this here??? What's the diference between major and minor fonts
                writer.WriteElement(new FontScheme() { Val = FontSchemeValues.Minor });

                //Close the single Font Tag
                writer.WriteEndElement();
            }

            // End Fonts section
            writer.WriteEndElement();

            #endregion

            #region Fills

            //Hardcoded Props
            var fills = new PatternValues[2] { PatternValues.None, PatternValues.Gray125 };

            writer.WriteStartElement(new Fills() { Count = (UInt32)fills.Length });

            foreach (var fill in fills)
            {
                writer.WriteStartElement(new Fill());

                writer.WriteElement(new PatternFill() { PatternType = fill });

                //Close the single Font Tag
                writer.WriteEndElement();
            }

            // End Fills section
            writer.WriteEndElement();

            #endregion


            #region Borders

            //Hardcoded Props
            var borderCount = 1;
            // Start Borders section
            writer.WriteStartElement(new Borders() { Count = (UInt32)borderCount });
            //Start border element
            writer.WriteStartElement(new Border());

            writer.WriteElement(new LeftBorder());
            writer.WriteElement(new RightBorder());
            writer.WriteElement(new TopBorder());
            writer.WriteElement(new BottomBorder());
            writer.WriteElement(new DiagonalBorder());

            //Close the boder
            writer.WriteEndElement();
            // End Borders section
            writer.WriteEndElement();

            #endregion

            #region CellStyleXfs (Cell Style Formats)

            // Creates a shared style table to apply to cells using an Id. 

            //Hardcoded Props
            var cellStyleXfsCount = 1;

            //Start CellStyleXfs element
            writer.WriteStartElement(new CellStyleFormats() { Count = (UInt32)cellStyleXfsCount });

            writer.WriteElement(new CellFormat() { NumberFormatId = (UInt32)0, FontId = (UInt32)0, FillId = (UInt32)0, BorderId = (UInt32)0 });

            // End CellStyleXfs section
            writer.WriteEndElement();

            #endregion

            #region CellXfs (CellFormats)

            //Hardcoded Props
            var cellXfsCount = 2;

            //Start CellStyleFormats section
            writer.WriteStartElement(new CellFormats() { Count = (UInt32)cellXfsCount });

            writer.WriteElement(new CellFormat() { NumberFormatId = (UInt32)0, FontId = (UInt32)0, FillId = (UInt32)0, BorderId = (UInt32)0 });
            writer.WriteElement(new CellFormat() { NumberFormatId = (UInt32)0, FontId = (UInt32)1, FillId = (UInt32)0, BorderId = (UInt32)0, ApplyFont = true });

            // End CellStyleFormats section
            writer.WriteEndElement();

            #endregion

            #region CellStyles

            //Hardcoded Props
            var cellStylesCount = 1;

            //Start CellStyleFormats element
            writer.WriteStartElement(new CellStyles() { Count = (UInt32)cellStylesCount });

            writer.WriteElement(new CellStyle() { Name = "Normal", FormatId = (UInt32)0, BuiltinId = (UInt32)0 });

            // End CellStyles section
            writer.WriteEndElement();

            #endregion

            #region Diferential formats

            //Hardcoded Props
            var diferentialFormatsCount = 1;

            // Start diferential formats section
            writer.WriteStartElement(new DifferentialFormats() { Count = (UInt32)diferentialFormatsCount });
            // Start diferential format tag
            writer.WriteStartElement(new DifferentialFormat());
            // Start font tag
            writer.WriteStartElement(new Font());

            writer.WriteElement(new Bold() { Val = false });
            writer.WriteElement(new Italic() { Val = false });
            writer.WriteElement(new Strike() { Val = false });
            writer.WriteElement(new Condense() { Val = false });
            writer.WriteElement(new Extend() { Val = false });
            writer.WriteElement(new Outline() { Val = false });
            writer.WriteElement(new Shadow() { Val = false });
            writer.WriteElement(new Underline() { Val = UnderlineValues.None });
            // Superscript, Subscript and Baseline
            writer.WriteElement(new VerticalTextAlignment() { Val = VerticalAlignmentRunValues.Baseline });
            writer.WriteElement(new FontSize() { Val = 11 });
            writer.WriteElement(new Color() { Theme = (UInt32)1 });
            writer.WriteElement(new FontName() { Val = "Calibri Light" });
            writer.WriteElement(new FontScheme() { Val = FontSchemeValues.Major });

            // End font tag
            writer.WriteEndElement();
            // End diferential format tag
            writer.WriteEndElement();
            // End diferential formats section
            writer.WriteEndElement();

            #endregion

            #region TableStyles

            writer.WriteElement(new TableStyles() { Count = 0, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" });

            #endregion

            #region Style Extensions List

            //Start extensions list
            writer.WriteStartElement(new StylesheetExtensionList());

            var guid = "{" + Guid.NewGuid() + "}";
            writer.WriteStartElement(new StylesheetExtension() { Uri = guid });
            writer.WriteElement(new x14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" });
            writer.WriteEndElement();

            guid = "{" + Guid.NewGuid() + "}";
            writer.WriteStartElement(new StylesheetExtension() { Uri = guid });
            writer.WriteElement(new x15.TimelineStyles() { DefaultTimelineStyle = "TimeSlicerStyleLight1" });
            writer.WriteEndElement();

            // End extensions list
            writer.WriteEndElement();

            #endregion


            //End styleSsheet
            writer.WriteEndElement();
            writer.Close();
        }
    }

    private void GenerateTableParts(TableDefinitionPart sheetTablesPart, string sheetPartId)
    {
        using (var writer = OpenXmlWriter.Create(sheetTablesPart))
        {
            var table = new Table() { Id = (UInt32Value)1U, Name = "Table", DisplayName = "Table", Reference = "A1:A4", TotalsRowShown = false, HeaderRowFormatId = (UInt32Value)0U };
            // Start Table element
            writer.WriteStartElement(table); 
            
            writer.WriteElement(new AutoFilter() { Reference = "A1:A4" });

            writer.WriteStartElement(new TableColumns() { Count = (UInt32Value)1U });
            writer.WriteElement(new TableColumn() { Id = (UInt32Value)1U, Name = "Header" });
            writer.WriteEndElement();

            writer.WriteElement(new TableStyleInfo() { Name = "TableStyleLight3", ShowFirstColumn = false, ShowLastColumn = false, ShowRowStripes = true, ShowColumnStripes = false });

            //End Table
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
