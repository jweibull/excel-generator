using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using x14 = DocumentFormat.OpenXml.Office2010.Excel;
using x15 = DocumentFormat.OpenXml.Office2013.Excel;
using ExcelGenerator.Excel;
using Newtonsoft.Json;
using System.Globalization;
using static ExcelGenerator.ExcelDefs.ExcelModelDefs;
using System.Text.RegularExpressions;


namespace ExcelGenerator.Generators;

public class SaxLib
{
    private readonly Dictionary<string, string> _sharedStringsToIndex = new Dictionary<string, string>();

    private int _sharedStringsCount;

    private int _sharedStringsUniqueCount;

    private readonly Dictionary<string, UInt32> _styleIndexes = new Dictionary<string, UInt32>();

    private readonly Dictionary<string, double> _oleADates = new Dictionary<string, double>();

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
                PrepareData(modelData.WorkbookModel);

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

    private void PrepareData(ExcelWorkbookModel workbookModel)
    {
        _sharedStringsCount = 0;
        foreach (var table in workbookModel.Tables)
        {
            AddToSharedStringDictionary(table.Header.Data);
            foreach (var column in table.Columns)
            {
                
                if (column.DataType == ExcelDataTypes.DataType.AutoDetect)
                {
                    // Check for either Dates or Hyperlinks on data colunms
                    PrepareAutodetectData(column, table.IsMultilined);
                }
                else
                {
                    // If not autodetect prepare regular types
                    PrepareDeclaredTypeData(column, table.IsMultilined);
                }
                                
            }
        }
        _sharedStringsUniqueCount = _sharedStringsToIndex.Count;
    }

    private void PrepareDeclaredTypeData(ExcelColumnModel column, bool isMultilined)
    {
        if (column.DataType == ExcelDataTypes.DataType.Text)
        {
            AddToSharedStringDictionary(column.Data);
        }
        else if (column.DataType == ExcelDataTypes.DataType.HyperLink)
        {
            var linkSample = column.Data.FirstOrDefault(x => !string.IsNullOrEmpty(x.Trim()) && x.Contains("href"));
            if (linkSample != null)
            {
                if (isMultilined)
                {
                    PrepareMultilinedHrefHyperlinks(column);
                }
                else
                {
                    PrepareHrefHyperlinks(column);
                }
            }
            else
            {
                if (isMultilined)
                {
                    PrepareMultilinedRegularHyperlinks(column);
                }
                else
                {
                    PrepareRegularHyperlinks(column);
                }
            }
            AddToSharedStringDictionary(column.Data);
        }
        else if (column.DataType == ExcelDataTypes.DataType.DateTime)
        {
            if (!AddToDatetimeToDictionary(column.Data, column.DataFormat))
            {
                column.DataType = ExcelDataTypes.DataType.Text;
                AddToSharedStringDictionary(column.Data);
            }
        }
    }

    private void PrepareAutodetectData(ExcelColumnModel column, bool isMultilined)
    {
        var linkSample = column.Data.FirstOrDefault(x => !string.IsNullOrEmpty(x.Trim()) && (x.Contains("href") || x.StartsWith("http://") || x.StartsWith("https://")));
        if (linkSample != null)
        {
            if (isMultilined)
            {
                PrepareMultilinedAutodetectedHyperlinks(column, linkSample);
            }
            else
            {
                PrepareAutodetectedHyperlinks(column, linkSample);
            }
            AddToSharedStringDictionary(column.Data);
        }
        else if (DateTime.TryParseExact(
            column.Data.FirstOrDefault(x => !string.IsNullOrEmpty(x.Trim())),
            CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern.ToString(),
            CultureInfo.InvariantCulture,
            DateTimeStyles.None,
            out var date))
        {
            column.DataType = ExcelDataTypes.DataType.DateTime;
            if (!AddToDatetimeToDictionary(column.Data, column.DataFormat))
            {
                column.DataType = ExcelDataTypes.DataType.Text;
                AddToSharedStringDictionary(column.Data);
            }
        }
        else
        {
            column.DataType = ExcelDataTypes.DataType.Text;
            AddToSharedStringDictionary(column.Data);
        }
    }

    private void PrepareMultilinedAutodetectedHyperlinks(ExcelColumnModel column, string linkSample)
    {
        if (linkSample.Contains("href"))
        {
            PrepareMultilinedHrefHyperlinks(column);
        }
        else
        {
            PrepareMultilinedRegularHyperlinks(column);
        }
    }

    private void PrepareMultilinedRegularHyperlinks(ExcelColumnModel column)
    {
        column.DataType = ExcelDataTypes.DataType.Text;
        var data = column.Data;
        for (int itemIndex = 0; itemIndex < data.Length; itemIndex++)
        {
            data[itemIndex] = Regex.Replace(data[itemIndex], "<br>", Environment.NewLine, RegexOptions.IgnoreCase);
        }
    }

    private void PrepareMultilinedHrefHyperlinks(ExcelColumnModel column)
    {
        column.DataType = ExcelDataTypes.DataType.Text;
        var data = column.Data;
        for (int itemIndex = 0; itemIndex < data.Length; itemIndex++)
        {
            string hyperlink = data[itemIndex];
            hyperlink = Regex.Replace(hyperlink, "<br>", Environment.NewLine, RegexOptions.IgnoreCase);
            var matches = Regex.Matches(hyperlink, @"<a.*?href=[\'""]?([^\'"" >]+).*?<\/a>", RegexOptions.IgnoreCase);

            foreach (Match match in matches)
            {
                hyperlink = hyperlink.Replace(match.Value, match.Groups[1].Value);
            }
            data[itemIndex] = hyperlink;
        }
    }

    private void PrepareAutodetectedHyperlinks(ExcelColumnModel column, string linkSample)
    {
        column.DataType = ExcelDataTypes.DataType.HyperLink;
        if (linkSample.Contains("href"))
        {
            PrepareHrefHyperlinks(column);
        }
        else
        {
            PrepareRegularHyperlinks(column);
        }
    }

    private void PrepareRegularHyperlinks(ExcelColumnModel column)
    {
        var data = column.Data;
        var hyperlinks = new List<ExcelHyperlink>();
        for (int itemIndex = 0; itemIndex < data.Length; itemIndex++)
        {
            if (!string.IsNullOrEmpty(data[itemIndex].Trim()))
            {
                data[itemIndex] = Regex.Replace(data[itemIndex], "<br>", Environment.NewLine, RegexOptions.IgnoreCase);
                hyperlinks.Add(new ExcelHyperlink() { Hyperlink = data[itemIndex] });
            }
        }
        column.AddHyperLinkData(hyperlinks.ToArray());
    }

    private void PrepareHrefHyperlinks(ExcelColumnModel column)
    {
        var data = column.Data;
        var hyperlinks = new List<ExcelHyperlink>();
        
        for (int itemIndex = 0; itemIndex < data.Length; itemIndex++)
        {
            if (!string.IsNullOrEmpty(data[itemIndex].Trim()))
            {
                string hyperlink = data[itemIndex];
                hyperlink = Regex.Replace(hyperlink, "<br>", Environment.NewLine, RegexOptions.IgnoreCase);

                string text = Regex.Replace(hyperlink, "(<[a|A][^>]*>|)", "");

                var matches = Regex.Matches(hyperlink, @"<a.*?href=[\'""]?([^\'"" >]+).*?<\/a>", RegexOptions.IgnoreCase);

                foreach (Match match in matches)
                {
                    hyperlink = hyperlink.Replace(match.Value, match.Groups[1].Value);
                }
                data[itemIndex] = text;
                hyperlinks.Add(new ExcelHyperlink() { Hyperlink = hyperlink });
            }
        }
        column.AddHyperLinkData(hyperlinks.ToArray());
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

    private bool AddToDatetimeToDictionary(string[] dates, string dataFormat)
    {
        var baseDataFormat = string.IsNullOrEmpty(dataFormat) ? CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern.ToString() : dataFormat;
        var index = 0;
        var isDate = true;
        DateTime date;
        while (index < dates.Length && isDate)
        {
            if (!_oleADates.ContainsKey(dates[index]))
            {
                if (string.IsNullOrEmpty(dataFormat))
                {
                    isDate = DateTime.TryParseExact(dates[index], baseDataFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out date);
                }
                else
                {
                    isDate = DateTime.TryParseExact(dates[index], baseDataFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out date);
                }
                
                if (isDate)
                {
                    _oleADates.Add(dates[index], date.ToOADate());
                } 
            }
            index++;
        }
        
        return isDate;
    }

    private void AddToSharedStringDictionary(string[] sharedStrings)
    {
        var count = 0;
        for (int itemIndex = 0; itemIndex < sharedStrings.Length; itemIndex++)
        {
            sharedStrings[itemIndex] = Regex.Replace(sharedStrings[itemIndex], "<br>", Environment.NewLine, RegexOptions.IgnoreCase);
            if (_sharedStringsToIndex.ContainsKey(sharedStrings[itemIndex]))
            {
                count++;
            }
            else
            {
                count++;
                _sharedStringsToIndex.Add(sharedStrings[itemIndex], _sharedStringsToIndex.Count().ToString());
            }
        }
        _sharedStringsCount += count;
    }

    private double FitColumn(string header, int headerFontSize, ExcelColumnModel column, bool isMultilined, int maxWidth)
    {
        var offset = 1;
        var numSamples = 50;
        double headerWidth = (header.Length + offset) * (72D / 96D) * (headerFontSize / 9D) * ((double)headerFontSize / (double)column.Style.FontSize);
        double columnWidth = (column.GetMaxDataLength(isMultilined, numSamples) + offset) * (72D / 96D) * (column.Style.FontSize / 9D) * ((double)column.Style.FontSize / (double)headerFontSize);
        var width = headerWidth >= columnWidth ? headerWidth : columnWidth;
        if (maxWidth > 13)
        {
            var higherFontSize = headerFontSize > column.Style.FontSize ? headerFontSize : column.Style.FontSize;
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
                var width = FitColumn(headers.Data[columnNum - 1], headers.Style.FontSize, allColumns[columnNum - 1], sheetModel.IsMultilined, allColumns[columnNum - 1].MaxWidth);
                writer.WriteElement(new Column() { Min = (UInt32)columnNum, Max = (UInt32)columnNum, Width = width, CustomWidth = true });
            }
            
            writer.WriteEndElement();

            writer.WriteStartElement(new SheetData());

            Row row = new Row();
            Cell cell = new Cell();
            CellValue cellValue = new CellValue();
            
            //Add header row
            row.RowIndex = 1U;
            writer.WriteStartElement(row);

            for (int columnNum = 1; columnNum <= numColumns; columnNum++)
            {
                cell.CellReference = string.Format("{0}{1}", GetColumnName(columnNum), 1U);

                cell.DataType = CellValues.SharedString;
                cell.StyleIndex = _styleIndexes[headers.StyleKey];
                writer.WriteStartElement(cell);
                cellValue.Text = _sharedStringsToIndex[headers.Data[columnNum - 1]];
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
                        cell.StyleIndex = _styleIndexes[currentColumn.StyleKey];

                        if (currentColumn.DataType == ExcelDataTypes.DataType.Text || currentColumn.DataType == ExcelDataTypes.DataType.HyperLink)
                        {
                            cell.DataType = CellValues.SharedString;
                            cellValue.Text = _sharedStringsToIndex[allColumns[columnNum - 1].Data[rowNum - 2]];
                        }
                        else if (currentColumn.DataType == ExcelDataTypes.DataType.DateTime)
                        {
                            cell.DataType = CellValues.Number;
                            cellValue.Text = _oleADates[allColumns[columnNum - 1].Data[rowNum - 2]].ToString();
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

    private string AddFontToDictionary(
        Dictionary<string, ExcelFontDetail> fonts, 
        ExcelFonts.FontType font, 
        int fontSize,
        int colorTheme)
    {
        string key;
        ExcelFontDetail fontDetail;

        key = font.ToString() + fontSize.ToString() + colorTheme.ToString();
        if (!fonts.ContainsKey(key))
        {
            fontDetail = ExcelFontDetail.GetFontStyles(font, (UInt32)fonts.Count, fontSize, colorTheme);
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

    // Everything is linked by a string id that is in fact the index of the array of style element. Ex the font with id "2"
    // will be the third font added in fonts section, while the font with id "0" will be the first you added.
    // Same goes for borders, fills, etc.
    private void GenerateStylePart(WorkbookStylesPart workbookStylesPart, ExcelWorkbookModel workbookModel)
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
        foreach (var table in workbookModel.Tables)
        {
            key = AddFontToDictionary(fonts, table.Header.Style.Font, table.Header.Style.FontSize, 1);
                
            styleKey = key + ExcelDataTypes.DataType.Text.ToString();
                
            AddStyleFormatToDictionary(styleFormats, styleKey, (UInt32)fonts[key].FontIndex, 0U, 0U, 0U, 0U, false, true);
                
            table.Header.AddStyleKey(styleKey);

            foreach (var column in table.Columns)
            {
                key = column.DataType == ExcelDataTypes.DataType.HyperLink ? 
                    AddFontToDictionary(fonts, column.Style.Font, column.Style.FontSize, 10) :
                    AddFontToDictionary(fonts, column.Style.Font, column.Style.FontSize, 1);
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
                        }
                        else
                        {
                            AddStyleFormatToDictionary(styleFormats, styleKey, fonts[key].FontIndex, 0U, 0U, 0U, 0U, true, true);
                        }
                        column.AddStyleKey(styleKey);
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
                        var sysDateFormat = string.IsNullOrEmpty(column.DataFormat) ? CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern.ToString() : column.DataFormat;
                        var numFormatId = AddNumFormatToDictionary(numFormats, sysDateFormat);
                        styleKey = styleKey + numFormatId.ToString();
                        AddStyleFormatToDictionary(styleFormats, styleKey, fonts[key].FontIndex, numFormatId, 0U, 0U, 0U, true, true);
                        column.AddStyleKey(styleKey);
                    }
                }
            }
        }

        

        //TODO Add Chart Fonts;

        // Future Watermark Details
        if (workbookModel.Watermark != null)
        {
            key = AddFontToDictionary(fonts, workbookModel.Watermark.Font, workbookModel.Watermark.FontSize, 1);

            styleKey = key + ExcelDataTypes.DataType.Text.ToString();

            AddStyleFormatToDictionary(styleFormats, key, (UInt32)fonts[key].FontIndex, 0U, 0U, 0U, 0U, false, true);

            workbookModel.Watermark.AddStyleKey(styleKey);
        }

        #endregion

        using (var writer = OpenXmlWriter.Create(workbookStylesPart))
        {
            writer.WriteStartElement(new Stylesheet());

            #region NumFormats

            writer.WriteStartElement(new NumberingFormats() { Count = (UInt32)numFormats.Count });
            
            foreach (var format in numFormats.Values)
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
