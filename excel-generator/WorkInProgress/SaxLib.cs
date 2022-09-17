using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using x14 = DocumentFormat.OpenXml.Office2010.Excel;
using x15 = DocumentFormat.OpenXml.Office2013.Excel;
using ExcelGenerator.Excel;
using Newtonsoft.Json;
using System.Globalization;
using static ExcelGenerator.ExcelDefs.ExcelModelDefs;

namespace ExcelGenerator.Generators;

public class SaxLib
{
    private readonly Dictionary<string, string> _sharedStringsToIndex = new Dictionary<string, string>();

    private int _sharedStringsCount;

    private int _sharedStringsUniqueCount;

    private readonly Dictionary<string, UInt32> _styleIndexes = new Dictionary<string, UInt32>();

    public void Run()
    {
        string path = Directory.GetCurrentDirectory();
        path = Path.Combine(path, "..", "..", "..", "output");
        for (int i = 0; i < 1; i++)
        {
            var nameCounter = 1;
            var baseFilename = "output";
            var filename = baseFilename;
            while (File.Exists(Path.Combine(path, filename + ".xlsx")))
            {
                filename = baseFilename + nameCounter++;
            }
            filename = Path.Combine(path, filename + ".xlsx");

            CreatePackage(filename);
        }
    }

    public void CreatePackage(string filename)
    {
        var serializer = new JsonSerializer();

        ModelData? modelData;

        using (StreamReader sr = new StreamReader(@"d:\excel.json"))
        using (var jsonTextReader = new JsonTextReader(sr))
        {
            modelData = serializer.Deserialize<ModelData>(jsonTextReader);
        }

        if (modelData == null)
        {
            throw new Exception("Model should not be null");
        }

        using (var stream = new MemoryStream())
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                // TestData
                var partId = 1;
                string sharedTableId = string.Empty;
                string stylesPartId = string.Empty;
                string sheetPartId = string.Empty;

                document.AddWorkbookPart();
                
                if (document.WorkbookPart != null)
                {
                    // Generate all Shared Strings that will be used in all the sheets
                    _sharedStringsCount = 0;
                    AddToSharedStringDictionary(modelData.WorkbookModel.Tables[0].Header.Data);
                    AddToSharedStringDictionary(modelData.WorkbookModel.Tables[0].Columns[0].Data);
                    _sharedStringsUniqueCount = _sharedStringsToIndex.Count;

                    // Generate a single sheet 
                    stylesPartId = "rId" + partId++;
                    sharedTableId = "rId" + partId++;
                    sheetPartId = "rId" + partId++;

                    // Generate all Styles needed on every sheet in this workbook
                    WorkbookStylesPart workbookStylesPart = document.WorkbookPart.AddNewPart<WorkbookStylesPart>(stylesPartId);
                    SharedStringTablePart sharedStringTablePart = document.WorkbookPart.AddNewPart<SharedStringTablePart>(sharedTableId);
                    GenerateStylePart(workbookStylesPart, stylesPartId, modelData);
                    GenerateSharedStringsTable(sharedStringTablePart, sharedTableId);

                    WorksheetPart workSheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>(sheetPartId);
                    TableDefinitionPart sheetTablesPart = workSheetPart.AddNewPart<TableDefinitionPart>(sheetPartId);

                    GenerateWorkSheetData(workSheetPart, modelData, sheetPartId, sheetTablesPart);
                    

                    // Create the worksheet and sheets list to end the package
                    using (var writer = OpenXmlWriter.Create(document.WorkbookPart))
                    {
                        writer.WriteStartElement(new Workbook());
                        writer.WriteStartElement(new Sheets());

                        writer.WriteElement(new Sheet()
                        {
                            Name = modelData.WorkbookModel.Tables[0].Name,
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

    private void GenerateSharedStringsTable(SharedStringTablePart sharedStringTablePart, string sharedTableId)
    {
        // Run this for all strings in the workbook
        // string[] sharedStrings must contain all the strings in the project

        using (var writer = OpenXmlWriter.Create(sharedStringTablePart))
        {
            // Change this based on real data count
            writer.WriteStartElement(new SharedStringTable() { Count = (UInt32)_sharedStringsCount, UniqueCount = (UInt32)_sharedStringsUniqueCount });

            foreach (var key in _sharedStringsToIndex.Keys)
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
            if (this._sharedStringsToIndex.ContainsKey(item))
            {
                count++;
            }
            else
            {
                count++;
                _sharedStringsToIndex.Add(item, _sharedStringsToIndex.Count().ToString());
            }
        }
        _sharedStringsCount += count;
    }

    private void GenerateWorkSheetData(WorksheetPart workSheetPart, ModelData modelData, string sheetPartId, TableDefinitionPart sheetTablesPart)
    {
        // Actual Cell Values from string table
        using (var writer = OpenXmlWriter.Create(workSheetPart))
        {
            var headers = modelData.WorkbookModel.Tables[0].Header;
            var allColumns = modelData.WorkbookModel.Tables[0].Columns;
            var numColumns = allColumns.Count();
            //+1 is for the headers
            int numRows = allColumns.OrderBy(x => x.Data.Count()).Select(x => x.Data.Count()).LastOrDefault(0) + 1;

            GenerateTableParts(sheetTablesPart, sheetPartId, modelData.WorkbookModel.Tables[0].Header, modelData.WorkbookModel.Tables[0].Theme, numRows);

            writer.WriteStartElement(new Worksheet());


            //Alinhar com o Table generation
            for (int columnNum = 1; columnNum <= numColumns; columnNum++)
            {
                writer.WriteStartElement(new Columns() { });
                writer.WriteElement(new Column() { Min = (UInt32)columnNum, Max = (UInt32)columnNum, Width = allColumns[columnNum - 1].MaxWidth, CustomWidth = true });
                writer.WriteEndElement();
            }

            writer.WriteStartElement(new SheetData());

            Row row = new Row();
            Cell cell = new Cell();
            CellValue cellValue = new CellValue();
            
            //Add header row
            row.RowIndex = 1U;
            writer.WriteStartElement(row);

            for (int columnNum = 1; columnNum <= numColumns; columnNum++)
            {
                //write the cell start element with the type and reference attributes
                cell.CellReference = string.Format("{0}{1}", GetColumnName(columnNum), 1U);

                cell.DataType = CellValues.SharedString;
                cell.StyleIndex = _styleIndexes[headers.StyleKey];
                writer.WriteStartElement(cell);
                //write the cell value
                cellValue.Text = _sharedStringsToIndex[headers.Data[columnNum - 1]];
                writer.WriteElement(cellValue);

                // write the end cell element
                writer.WriteEndElement();
            }

            writer.WriteEndElement();

            for (int rowNum = 2; rowNum <= numRows; rowNum++)
            {
                //write the row start element with the row index attribute
                row.RowIndex = (UInt32)rowNum;
                writer.WriteStartElement(row);

                for (int columnNum = 1; columnNum <= numColumns; columnNum++)
                {
                    var currentColumn = allColumns[columnNum - 1];
                    //write the cell start element with the type and reference attributes
                    cell.CellReference = string.Format("{0}{1}", GetColumnName(columnNum), rowNum);
                    
                    cell.DataType = CellValues.SharedString;
                    cell.StyleIndex = _styleIndexes[currentColumn.StyleKey];
                    writer.WriteStartElement(cell);
                    //write the cell value
                    cellValue.Text = _sharedStringsToIndex[allColumns[columnNum - 1].Data[rowNum - 2]];
                    writer.WriteElement(cellValue);

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

    private void GenerateTableParts(TableDefinitionPart sheetTablesPart, string sheetPartId, ExcelHeaderModel headers, ExcelThemes.Theme theme, int numRows)
    {
        var numColumns = headers.Data.Count();

        using (var writer = OpenXmlWriter.Create(sheetTablesPart))
        {
            var reference = "A1:" + GetColumnName(numColumns) + numRows.ToString();

            var table = new Table()
            {
                Id = (UInt32)1U,
                Name = "Table",
                DisplayName = "Table",
                Reference = reference,
                TotalsRowShown = false,
            };
            // Start Table element
            writer.WriteStartElement(table);

            writer.WriteElement(new AutoFilter() { Reference = reference });

            writer.WriteStartElement(new TableColumns() { Count = (UInt32)numColumns });

            for (int columnNum = 1; columnNum <= numColumns; columnNum++)
            {
                writer.WriteElement(new TableColumn()
                {
                    Id = (UInt32)columnNum,
                    Name = headers.Data[columnNum - 1]
                });
            }

            writer.WriteEndElement();

            writer.WriteElement(new TableStyleInfo()
            {
                Name = ExcelThemes.GetTheme(theme),
                ShowFirstColumn = false,
                ShowLastColumn = false,
                ShowRowStripes = true,
                ShowColumnStripes = false
            });

            //End Table
            writer.WriteEndElement();
            writer.Close();
        }
    }

    private string AddFontToDictionary(
        Dictionary<string, ExcelFontDetail> fonts, 
        ExcelFonts.FontType font, 
        int fontSize)
    {
        string key;
        ExcelFontDetail fontDetail;

        key = font.ToString() + fontSize.ToString();
        if (!fonts.ContainsKey(key))
        {
            fontDetail = ExcelFontDetail.GetFontStyles(font, (UInt32)fonts.Count, fontSize);
            fonts.Add(key, fontDetail);
        }

        return key;
    }
    
    private UInt32 AddNumFormatToDictionary(Dictionary<string, ExcelNumFormat> numFormats, string dataFormat)
    {
        UInt32 numFormatId;
        
        if (numFormats.ContainsKey(dataFormat))
        {
            numFormatId = numFormats[dataFormat].FormatId;
        }
        else
        {
            numFormatId = StyleContants.StartIndex + (UInt32)numFormats.Count;
            ExcelNumFormat numFormat = new ExcelNumFormat(dataFormat, numFormatId);
            numFormats.Add(dataFormat, numFormat);
        }
        
        return numFormatId;
    }

    private void AddStyleFormatToDictionary(
        Dictionary<string, ExcelStyleFormat> styleFormats,
        string styleKey,
        UInt32 fontIdx, 
        UInt32 numFormatIdx, 
        UInt32 cellStyleIdx, 
        UInt32 fillIdx, 
        UInt32 borderIdx,
        bool applyNumFormat,
        bool applyFont)
    {
        ExcelStyleFormat styleFormat;

        if (!styleFormats.ContainsKey(styleKey))
        {
            styleFormat = new ExcelStyleFormat(fontIdx, numFormatIdx, cellStyleIdx, fillIdx, borderIdx, (UInt32)styleFormats.Count + 1);
            styleFormat.ApplyFont = applyFont;
            styleFormat.ApplyNumFormat = applyNumFormat;
            styleFormats.Add(styleKey, styleFormat);
            _styleIndexes.Add(styleKey, styleFormat.StyleIndex);
        }
    }

    // TODO create dictionaries to link fonts and fills etc to the CellFormatStyles element inside this method.
    // Everything is linked by a string id that is in fact the index of the array of style element. Ex the font with id "2"
    // will be the third font added in fonts section, while the font with id "0" will be the first you added.
    // Same goes for borders, fills, etc.
    private void GenerateStylePart(WorkbookStylesPart workbookStylesPart, string stylesPartId, ModelData modelData )
    {
        #region Fonts, NumFormats, CellXfs and CellStyles

        // TODO Write all fonts and styles here!!!

        var fonts = new Dictionary<string, ExcelFontDetail>();
        var numFormats = new Dictionary<string, ExcelNumFormat>();
        var hyperlinkFormats = new Dictionary<string, UInt32>();
        var styleFormats = new Dictionary<string, ExcelStyleFormat>();
        string styleKey;
        string key;

        //Run all tables looking for styles
        foreach (var table in modelData.WorkbookModel.Tables)
        {
            key = AddFontToDictionary(fonts, table.Header.Style.Font, table.Header.Style.FontSize);
                
            styleKey = key + ExcelDataTypes.DataType.Text.ToString();
                
            AddStyleFormatToDictionary(styleFormats, styleKey, (UInt32)fonts[key].FontIndex, 0U, 0U, 0U, 0U, false, true);
                
            table.Header.AddStyleKey(styleKey);

            foreach (var column in table.Columns)
            {
                key = AddFontToDictionary(fonts, column.Style.Font, column.Style.FontSize);
                // Columns can have diferent types, formats and fonts
                styleKey = key + ExcelDataTypes.DataType.Text.ToString();
                if (column.DataType == ExcelDataTypes.DataType.Text)
                {
                    AddStyleFormatToDictionary(styleFormats, styleKey, fonts[key].FontIndex, 0U, 0U, 0U, 0U, false, true);
                    column.AddStyleKey(styleKey);
                }
                else
                {
                    // Add numFormat or CellStyle to xml and get index to add to the style class
                    styleKey = key + column.DataType.ToString();
                    
                    if (column.DataType == ExcelDataTypes.DataType.Number)
                    {
                        if (!string.IsNullOrEmpty(column.DataFormat))
                        {
                            var numFormatId = AddNumFormatToDictionary(numFormats, column.DataFormat);
                            styleKey = styleKey + numFormatId.ToString();
                            AddStyleFormatToDictionary(styleFormats, styleKey, fonts[key].FontIndex, numFormatId, 0U, 0U, 0U, true, true);
                            column.AddStyleKey(styleKey);
                        }   
                    }
                    else if (column.DataType == ExcelDataTypes.DataType.HyperLink)
                    {
                        styleKey = key + column.DataType.ToString();
                        hyperlinkFormats.TryAdd(key, fonts[key].FontIndex);
                        AddStyleFormatToDictionary(styleFormats, styleKey, fonts[key].FontIndex, 0U, 1U, 0U, 0U, false, false);
                        column.AddStyleKey(styleKey);
                    }
                    else if (column.DataType == ExcelDataTypes.DataType.DateTime)
                    {
                        //DateTime dt = DateTime.Now;
                        //double x = dt.ToOADate();
                        string sysDateFormat;
                        if (string.IsNullOrEmpty(column.DataFormat))
                        {
                            sysDateFormat = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;
                        }
                        else
                        {
                            sysDateFormat = column.DataFormat;
                        }
                        var numFormatId = AddNumFormatToDictionary(numFormats, sysDateFormat);
                        styleKey = styleKey + numFormatId.ToString();
                        AddStyleFormatToDictionary(styleFormats, styleKey, fonts[key].FontIndex, numFormatId, 0U, 0U, 0U, true, true);
                        column.AddStyleKey(styleKey);
                    }
                }
            }
        }

        #endregion

        //TODO Add Chart Fonts;

        // Future Watermark Details
        if (modelData.WorkbookModel.Watermark != null)
        {
            key = AddFontToDictionary(fonts, modelData.WorkbookModel.Watermark.Font, modelData.WorkbookModel.Watermark.FontSize);

            styleKey = key + ExcelDataTypes.DataType.Text.ToString();

            AddStyleFormatToDictionary(styleFormats, key, (UInt32)fonts[key].FontIndex, 0U, 0U, 0U, 0U, false, true);

            modelData.WorkbookModel.Watermark.AddStyleKey(styleKey);
        }

        using (var writer = OpenXmlWriter.Create(workbookStylesPart))
        {
            writer.WriteStartElement(new Stylesheet());

            #region NumFormats

            //TODO add numformat fields here for numbering and Dates

            #endregion

            #region Fonts

            //write the fonts sections
            //<Fonts>
            //  <Font>...props...</Font>
            //</Fonts>
            //writer.WriteStartElement(new Fonts() { Count = (UInt32)hardCodedFonts.Length });
            writer.WriteStartElement(new Fonts() { Count = (UInt32)fonts.Count });

            foreach (var font in fonts.Values)
            {
                writer.WriteStartElement(new Font());

                writer.WriteElement(new FontSize() { Val = font.FontSize });
                writer.WriteElement(new Color() { Theme = font.Theme });
                writer.WriteElement(new FontName() { Val = font.FontName });
                writer.WriteElement(new FontFamily() { Val = font.FontFamily });

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

            var cellStyleXfsCount = hyperlinkFormats.Count + 1;
            //Start CellStyleXfs element
            writer.WriteStartElement(new CellStyleFormats() { Count = (UInt32)cellStyleXfsCount });
            //Hardcoded base CellFormat
            writer.WriteElement(new CellFormat() { NumberFormatId = (UInt32)0, FontId = (UInt32)0, FillId = (UInt32)0, BorderId = (UInt32)0 });

            foreach(var fontIndex in hyperlinkFormats.Values)
            {
                writer.WriteElement(new CellFormat()
                {
                    NumberFormatId = (UInt32)0,
                    FontId = fontIndex,
                    FillId = (UInt32)0,
                    BorderId = (UInt32)0,
                    ApplyNumberFormat = false,
                    ApplyFill = false,
                    ApplyBorder = false,
                    ApplyAlignment = false,
                    ApplyProtection = false
                });
            }
            
            // End CellStyleXfs section
            writer.WriteEndElement();

            #endregion

            #region CellXfs (CellFormats)
            
            // Add all alignment and apply numberformat features
            
            var cellXfsCount = styleFormats.Count() + 1;

            //Start CellStyleFormats section
            writer.WriteStartElement(new CellFormats() { Count = (UInt32)cellXfsCount });

            writer.WriteElement(new CellFormat() { NumberFormatId = (UInt32)0, FontId = (UInt32)0, FillId = (UInt32)0, BorderId = (UInt32)0 });

            foreach (var styleFormat in styleFormats.Values)
            {
                writer.WriteStartElement(new CellFormat() 
                { 
                    NumberFormatId = styleFormat.NumFormatIndex,
                    FontId = styleFormat.FontIndex, 
                    FillId = (UInt32)0, 
                    BorderId = (UInt32)0, 
                    ApplyFont = styleFormat.ApplyFont,
                    ApplyNumberFormat = styleFormat.ApplyNumFormat,
                    ApplyAlignment = true,
                    
                });
                writer.WriteElement(new Alignment() 
                {
                    Horizontal = HorizontalAlignmentValues.Left,
                    Vertical = VerticalAlignmentValues.Center, 
                    WrapText = true 
                });
                //End CellXf
                writer.WriteEndElement();
            }
            // End CellStyleFormats section
            writer.WriteEndElement();

            #endregion

            #region CellStyles

            var cellStylesCount = hyperlinkFormats.Count() + 1;

            //Start CellStyleFormats element
            writer.WriteStartElement(new CellStyles() { Count = (UInt32)cellStylesCount });

            writer.WriteElement(new CellStyle() { Name = "Normal", FormatId = (UInt32)0, BuiltinId = (UInt32)0 });
            if (hyperlinkFormats.Count() > 0)
            {
                writer.WriteElement(new CellStyle() { Name = "Hyperlink", FormatId = (UInt32)1, BuiltinId = (UInt32)8 });
            }
            // End CellStyles section
            writer.WriteEndElement();

            #endregion

            #region Diferential formats

            //Hardcoded Props
            var diferentialFormatsCount = 1;

            // Start diferential formats section Although empty it is a needed part of the Stylesheet
            writer.WriteStartElement(new DifferentialFormats() { Count = (UInt32)diferentialFormatsCount });
            // Start diferential format tag
            writer.WriteStartElement(new DifferentialFormat());
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
