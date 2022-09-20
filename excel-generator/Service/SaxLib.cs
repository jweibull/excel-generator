using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using x14 = DocumentFormat.OpenXml.Office2010.Excel;
using x15 = DocumentFormat.OpenXml.Office2013.Excel;
using ExcelGenerator.Excel;
using Newtonsoft.Json;
using System.Globalization;
using static ExcelGenerator.ExcelDefs.ExcelModelDefs;
using ExcelGenerator.ExcelDefs;


namespace ExcelGenerator.Generators;

public class SaxLib
{
    private readonly DataParser _parser;
    private readonly StyleParser _styleParser;
    
    public SaxLib()
    {
        _parser = new DataParser();
        _styleParser = new StyleParser();
    }

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
                document.AddWorkbookPart();

                if (document.WorkbookPart == null)
                {
                    throw new Exception("Error creating workbook part");
                }

                // Generate all Shared Strings that will be used in all the sheets
                _parser.PrepareData(modelData.WorkbookModel);

                // Generate all Styles needed on every sheet in this workbook
                var stylesPartId = "sPrId1";
                var sharedTableId = "sTrId1";
                WorkbookStylesPart workbookStylesPart = document.WorkbookPart.AddNewPart<WorkbookStylesPart>(stylesPartId);
                SharedStringTablePart sharedStringTablePart = document.WorkbookPart.AddNewPart<SharedStringTablePart>(sharedTableId);
                GenerateStylePart(workbookStylesPart, modelData.WorkbookModel);
                GenerateSharedStringsTable(sharedStringTablePart);

                var partId = 1;
                var linksId = 1;
                List<string> sheetPartIds = new List<string>();
                var numSheets = modelData.WorkbookModel.Tables.Count();
                for (int sheetNum = 1; sheetNum <= numSheets; sheetNum++)
                {
                    var sheetPartId = "rId" + partId++;
                    sheetPartIds.Add(sheetPartId);
                    WorksheetPart workSheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>(sheetPartId);
                    TableDefinitionPart sheetTablesPart = workSheetPart.AddNewPart<TableDefinitionPart>(sheetPartId);

                    var sheetModel = modelData.WorkbookModel.Tables[sheetNum - 1];

                    var allColumns = sheetModel.Columns;

                    var linkColumns = allColumns.Where(x => x.DataType == ExcelDataTypes.DataType.HyperLink).ToList();
                    foreach (var linkColumn in linkColumns)
                    {
                        linksId = GenerateHyperlinkParts(workSheetPart, linkColumn, linksId);
                    }

                    int numRows = allColumns.Select(x => x.Data.Count()).Max() + 1;

                    GenerateWorkSheetData(workSheetPart, sheetModel, allColumns, numRows, sheetPartId);
                    GenerateTableParts(sheetTablesPart, (UInt32)sheetNum, sheetModel.Header, sheetModel.Theme, numRows);
                }

                // Create the worksheet and sheets list to end the package
                FinishDocument(document.WorkbookPart, modelData.WorkbookModel, numSheets, sheetPartIds);
                
                //document.Save();
                document.SaveAs(filename);
                document.Close();
            }
        }
    }

    private void FinishDocument(WorkbookPart workbookPart, ExcelWorkbookModel workbookModel, int numSheets, List<string> sheetPartIds)
    {
        using (var writer = OpenXmlWriter.Create(workbookPart))
        {
            writer.WriteStartElement(new Workbook());
            writer.WriteStartElement(new Sheets());

            for (int sheetNum = 1; sheetNum <= numSheets; sheetNum++)
            {
                writer.WriteElement(new Sheet()
                {
                    Name = workbookModel.Tables[sheetNum - 1].Name,
                    SheetId = (UInt32)sheetNum,
                    Id = sheetPartIds[sheetNum - 1]
                });
            }

            // End Sheets
            writer.WriteEndElement();
            // End Workbook
            writer.WriteEndElement();

            writer.Close();
        }
    }

    private int GenerateHyperlinkParts(WorksheetPart workSheetPart, ExcelColumnModel linkColumn, int partIdSequencer)
    {
        string url;
        foreach (var link in linkColumn.HyperLinkData)
        {
            var id = "lId" + partIdSequencer++;
            url = link.Hyperlink.StartsWith("http://") || link.Hyperlink.StartsWith("https://") ? link.Hyperlink : @"http://" + link.Hyperlink;
            workSheetPart.AddHyperlinkRelationship(new Uri(url, UriKind.Absolute), true, id);
            link.LinkId = id;
        }
        return partIdSequencer;
    }

    private void GenerateSharedStringsTable(SharedStringTablePart sharedStringTablePart)
    {
        // Run this for all strings in the workbook
        // string[] sharedStrings must contain all the strings in the project

        using (var writer = OpenXmlWriter.Create(sharedStringTablePart))
        {
            // Change this based on real data count
            writer.WriteStartElement(new SharedStringTable() { Count = (UInt32)_parser.SharedStringsCount, UniqueCount = (UInt32)_parser.SharedStringsUniqueCount });

            foreach (var key in _parser.SharedStringsToIndex.Keys)
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

    private double FitColumn(string header, ExcelStyleClasses headerStyle, ExcelColumnModel column, bool isMultilined, int maxWidth)
    {
        var hOffset = 2;
        var cOffset = 0;
        var numSamples = 50;

        var hFontFactor = ExcelModelDefs.ExcelFonts.GetFontSizeFactor(headerStyle.Font);
        var cFontFactor = ExcelModelDefs.ExcelFonts.GetFontSizeFactor(column.Style.Font);

        if (column.DataType == ExcelDataTypes.DataType.DateTime)
        {
            cOffset = 5;
        }

        if (headerStyle.Bold)
        {
            hFontFactor -= 0.5;
        }

        if (column.Style.Bold == true)
        {
            cFontFactor -= 0.5D;
        }

        double headerWidth = (header.Length + hOffset) * (72D / 96D) * (headerStyle.FontSize / hFontFactor) * ((double)headerStyle.FontSize / (double)column.Style.FontSize);
        double columnWidth = (column.GetMaxDataLength(isMultilined, numSamples) + cOffset) * (72D / 96D) * (column.Style.FontSize / cFontFactor) * ((double)column.Style.FontSize / (double)headerStyle.FontSize);
        
        var width = headerWidth >= columnWidth ? headerWidth : columnWidth;
        
        if (maxWidth > 13)
        {
            var higherFontSize = headerStyle.FontSize > column.Style.FontSize ? headerStyle.FontSize : column.Style.FontSize;
            var correctedMaxWidth = maxWidth * (72D / 96D) * (higherFontSize / 9D);
            width = width > correctedMaxWidth ? correctedMaxWidth : width;
        }
        
        return width;
    }

    private void GenerateWorkSheetData(WorksheetPart workSheetPart, ExcelTableSheetModel sheetModel, ExcelColumnModel[] allColumns, int numRows, string sheetPartId)
    {
        // Actual Cell Values from string table
        using (var writer = OpenXmlWriter.Create(workSheetPart))
        {
            var headers = sheetModel.Header;
            var numColumns = allColumns.Count();
                        
            writer.WriteStartElement(new Worksheet());

            //Alinhar com o Table generation
            writer.WriteStartElement(new Columns());
            for (int columnNum = 1; columnNum <= numColumns; columnNum++)
            {
                var width = FitColumn(headers.Data[columnNum - 1], headers.Style, allColumns[columnNum - 1], sheetModel.IsMultilined, allColumns[columnNum - 1].MaxWidth);
                writer.WriteElement(new Column() { Min = (UInt32)columnNum, Max = (UInt32)columnNum, Width = width, CustomWidth = true });
            }
            
            writer.WriteEndElement();

            writer.WriteStartElement(new SheetData());

            Row row = new Row();
            Cell cell = new Cell();
            CellValue cellValue = new CellValue();
            var stringIndexes = _parser.SharedStringsToIndex;
            var dates = _parser.OleADates;
            //Add header row
            row.RowIndex = 1U;
            writer.WriteStartElement(row);

            for (int columnNum = 1; columnNum <= numColumns; columnNum++)
            {
                cell.CellReference = string.Format("{0}{1}", GetColumnName(columnNum), 1U);

                cell.DataType = CellValues.SharedString;
                cell.StyleIndex = _styleParser.StyleIndexes[headers.StyleKey];
                writer.WriteStartElement(cell);
                cellValue.Text = stringIndexes[headers.Data[columnNum - 1]];
                writer.WriteElement(cellValue);

                writer.WriteEndElement();
            }

            writer.WriteEndElement();

            // Add the rest of the data
            for (int rowNum = 2; rowNum <= numRows; rowNum++)
            {
                //write the row start element with the row index attribute
                row.RowIndex = (UInt32)rowNum;
                writer.WriteStartElement(row);

                for (int columnNum = 1; columnNum <= numColumns; columnNum++)
                {
                    var currentColumn = allColumns[columnNum - 1];
                    if (allColumns.Length > (columnNum - 1) && allColumns[columnNum - 1].Data.Length > (rowNum -2))
                    {
                        cell.CellReference = string.Format("{0}{1}", GetColumnName(columnNum), rowNum);
                        cell.StyleIndex = _styleParser.StyleIndexes[currentColumn.StyleKey];

                        if (currentColumn.DataType == ExcelDataTypes.DataType.Text || currentColumn.DataType == ExcelDataTypes.DataType.HyperLink)
                        {
                            cell.DataType = CellValues.SharedString;
                            cellValue.Text = stringIndexes[allColumns[columnNum - 1].Data[rowNum - 2]];
                        }
                        else if (currentColumn.DataType == ExcelDataTypes.DataType.DateTime)
                        {
                            cell.DataType = CellValues.Number;
                            if (dates.ContainsKey(allColumns[columnNum - 1].Data[rowNum - 2]))
                            {
                                cellValue.Text = dates[allColumns[columnNum - 1].Data[rowNum - 2]].ToString();
                            }
                            else
                            {
                                cellValue.Text = String.Empty;
                            }
                        }
                        else
                        {
                            cell.DataType = CellValues.Number;
                            cellValue.Text = allColumns[columnNum - 1].Data[rowNum - 2];
                        }

                        writer.WriteStartElement(cell);
                        writer.WriteElement(cellValue);
                        writer.WriteEndElement();
                    }
                }
                writer.WriteEndElement();
            }

            // write the end SheetData element
            writer.WriteEndElement();

            //HyperlinksInfo
            if (allColumns.Any(x => x.DataType == ExcelDataTypes.DataType.HyperLink))
            {
                writer.WriteStartElement(new Hyperlinks());
                for (int columnNum = 1; columnNum <= numColumns; columnNum++)
                {
                    if (allColumns[columnNum - 1].DataType == ExcelDataTypes.DataType.HyperLink)
                    {
                        var linkColumn = allColumns[columnNum - 1];
                        var hyperlink = new Hyperlink();
                        for (int rowNum = 2; rowNum <= linkColumn.HyperLinkData.Length + 1; rowNum++)
                        {
                            hyperlink.Reference = string.Format("{0}{1}", GetColumnName(columnNum), rowNum);
                            hyperlink.Id = linkColumn.HyperLinkData[rowNum - 2].LinkId;
                            writer.WriteElement(hyperlink);
                        }
                    }
                }
                writer.WriteEndElement();
            }

            // Table Info
            writer.WriteStartElement(new TableParts() { Count = 1 });
            writer.WriteElement(new TablePart() { Id = sheetPartId });
            writer.WriteEndElement();

            // write the end Worksheet element
            writer.WriteEndElement();
            writer.Close();
        }
    }

    private void GenerateTableParts(TableDefinitionPart sheetTablesPart, UInt32 tableId, ExcelHeaderModel headers, ExcelThemes.Theme theme, int numRows)
    {
        var numColumns = headers.Data.Count();

        using (var writer = OpenXmlWriter.Create(sheetTablesPart))
        {
            var reference = "A1:" + GetColumnName(numColumns) + numRows.ToString();

            var table = new Table()
            {
                Id = tableId,
                Name = "Table" + tableId.ToString(),
                DisplayName = "Table" + tableId.ToString(),
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

    // Everything is linked by a string id that is in fact the index of the array of style element. Ex the font with id "2"
    // will be the third font added in fonts section, while the font with id "0" will be the first you added.
    // Same goes for borders, fills, etc.
    private void GenerateStylePart(WorkbookStylesPart workbookStylesPart, ExcelWorkbookModel workbookModel)
    {
        _styleParser.ParseStyles(workbookModel);

        using (var writer = OpenXmlWriter.Create(workbookStylesPart))
        {
            writer.WriteStartElement(new Stylesheet());

            #region NumFormats

            writer.WriteStartElement(new NumberingFormats() { Count = (UInt32)_styleParser.NumFormats.Count });
            
            foreach (var format in _styleParser.NumFormats.Values)
            {
                writer.WriteElement(new NumberingFormat() { NumberFormatId = format.FormatId, FormatCode = format.FormatCode });  
            }

            writer.WriteEndElement();

            #endregion

            #region Fonts

                //write the fonts sections
                //<Fonts>
                //  <Font>...props...</Font>
                //</Fonts>
                //writer.WriteStartElement(new Fonts() { Count = (UInt32)hardCodedFonts.Length });
                writer.WriteStartElement(new Fonts() { Count = (UInt32)_styleParser.Fonts.Count });

            foreach (var font in _styleParser.Fonts.Values)
            {
                writer.WriteStartElement(new Font());
                if (font.Bold)
                {
                    writer.WriteElement(new Bold());
                }
                if (font.Italic)
                {
                    writer.WriteElement(new Italic());
                }
                if (font.Underline)
                {
                    writer.WriteElement(new Underline());
                }
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

            var cellStyleXfsCount = _styleParser.HyperlinkFormats.Count + 1;
            //Start CellStyleXfs element
            writer.WriteStartElement(new CellStyleFormats() { Count = (UInt32)cellStyleXfsCount });
            //Hardcoded base CellFormat
            writer.WriteElement(new CellFormat() { NumberFormatId = (UInt32)0, FontId = (UInt32)0, FillId = (UInt32)0, BorderId = (UInt32)0 });

            foreach(var fontIndex in _styleParser.HyperlinkFormats.Values)
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
            
            var cellXfsCount = _styleParser.StyleFormats.Count() + 1;

            //Start CellStyleFormats section
            writer.WriteStartElement(new CellFormats() { Count = (UInt32)cellXfsCount });

            writer.WriteElement(new CellFormat() { NumberFormatId = (UInt32)0, FontId = (UInt32)0, FillId = (UInt32)0, BorderId = (UInt32)0 });

            foreach (var styleFormat in _styleParser.StyleFormats.Values)
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

            var cellStylesCount = _styleParser.HyperlinkFormats.Count() + 1;

            //Start CellStyleFormats element
            writer.WriteStartElement(new CellStyles() { Count = (UInt32)cellStylesCount });

            writer.WriteElement(new CellStyle() { Name = "Normal", FormatId = (UInt32)0, BuiltinId = (UInt32)0 });
            if (_styleParser.HyperlinkFormats.Count() > 0)
            {
                writer.WriteElement(new CellStyle() { Name = "Hyperlink", FormatId = (UInt32)1, BuiltinId = (UInt32)8 });
            }
            // End CellStyles section
            writer.WriteEndElement();

            #endregion

            #region Diferential formats

            //Hardcoded Props
            var diferentialFormatsCount = 0;
            // Start diferential formats section Although empty it is a needed part of the Stylesheet
            writer.WriteStartElement(new DifferentialFormats() { Count = (UInt32)diferentialFormatsCount });
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
