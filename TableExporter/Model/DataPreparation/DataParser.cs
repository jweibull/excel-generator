using TableExporter.DataPreparation;

namespace TableExporter;

/// <summary>
/// Helper class that parses data into dictionaries that can be stored on excel files as indexes.
/// </summary>
internal class DataParser
{
    private ExcelDate _excelDate;
    private ExcelSheetlinkParser _sheetlinkParser;
    private ExcelSharedString _sharedString;
    
    internal DataParser()
    {
        _excelDate = new ExcelDate();
        _sheetlinkParser = new ExcelSheetlinkParser();
        _sharedString = new ExcelSharedString();
    }

    internal SharedStringCount GetSharedStringCount()
    {
        return _sharedString.GetSharedStringCount();
    }

    internal string[] GetSharedStrings()
    {
        return _sharedString.SharedStringsToIndex.Keys.ToArray();
    }

    internal string GetValue(ExcelModelDefs.ExcelDataTypes type, string key)
    {
        switch (type)
        {
            case ExcelModelDefs.ExcelDataTypes.Text:
            case ExcelModelDefs.ExcelDataTypes.Hyperlink:
            case ExcelModelDefs.ExcelDataTypes.Sheetlink:
                return _sharedString.GetValue(key);

            case ExcelModelDefs.ExcelDataTypes.DateTime:
                return _excelDate.GetValue(key);

            default:
                return key;
        }
    }

    internal void SanitizeData(ExcelWorkbookModel workbookModel)
    {
        foreach (var table in workbookModel.Tables)
        {
            for (var headerIndex = 0; headerIndex < table.Header.Data.Length; headerIndex++)
            {
                if (!string.IsNullOrEmpty(table.Header.Data[headerIndex]))
                {
                    table.Header.Data[headerIndex] = StringSanitizer.RemoveSpecialCharacters(table.Header.Data[headerIndex]);
                }
            }
            foreach (var column in table.Columns)
            {
                for (var columnIndex = 0; columnIndex < column.Data.Length; columnIndex++)
                {
                    if (!string.IsNullOrEmpty(column.Data[columnIndex]))
                    {
                        column.Data[columnIndex] = StringSanitizer.RemoveSpecialCharacters(column.Data[columnIndex]);
                    }
                }
            }
        }
    }

    internal void PrepareData(ExcelWorkbookModel workbookModel)
    {
        foreach (var table in workbookModel.Tables)
        {
            if (table.Header.Data.Length != table.Columns.Count())
            {
                throw new Exception("Length of Headers and number of columns must match");
            }

            table.Header.Data = PrepareHeaderData(table.Header.Data);
            _sharedString.AddToSharedStringDictionary(table.Header.Data, false, String.Empty);
            
            foreach (var column in table.Columns)
            {
                SetupColumn(table, column, workbookModel.GlobalColumnBehavior);

                switch (column.DataType)
                {
                    case ExcelModelDefs.ExcelDataTypes.DateTime:
                        _excelDate.AddToDatetimeToDictionary(column.Data, column.DataFormat);
                        break;

                    case ExcelModelDefs.ExcelDataTypes.Hyperlink:
                        ExcelHyperlinkParser.PrepareHyperlinks(column, workbookModel.GlobalColumnBehavior.Hyperlink.IsHtml, column.IsMultilined, column.NewLineSeparator);
                        _sharedString.AddToSharedStringDictionary(column.Data, column.IsMultilined, column.NewLineSeparator);
                        break;

                    case ExcelModelDefs.ExcelDataTypes.Sheetlink:
                        _sheetlinkParser.PrepareSheetlinks(column, workbookModel.AllSheets, column.IsMultilined);
                        _sharedString.AddToSharedStringDictionary(column.Data, column.IsMultilined, column.NewLineSeparator);
                        break;

                    case ExcelModelDefs.ExcelDataTypes.Number:
                    case ExcelModelDefs.ExcelDataTypes.Text:
                    default:
                        _sharedString.AddToSharedStringDictionary(column.Data, column.IsMultilined, column.NewLineSeparator);
                        break;
                }
            }
        }
    }

    private string[] PrepareHeaderData(string[] headerData)
    {
        Dictionary<string, int> stringCounts = new Dictionary<string, int>();
        int emptyCount = 1; // Counter for empty strings

        for (int i = 0; i < headerData.Length; i++)
        {
            string trimmedString = headerData[i].Trim();

            if (string.IsNullOrEmpty(trimmedString))
            {
                // If the string is empty, replace it with "Column" followed by a number
                headerData[i] = $"Column {emptyCount++}";
            }
            else
            {
                // If the string is not empty, append the count to make it unique
                int count = stringCounts.ContainsKey(trimmedString) ? stringCounts[trimmedString] + 1 : 1;
                stringCounts[trimmedString] = count;
                if (count > 1)
                {
                    headerData[i] = $"{trimmedString}{count}";
                }
            }
        }

        return headerData;
    }

    private void SetupColumn(ExcelTableSheetModel table, ExcelColumnModel column, ExcelGlobalBehavior globalBehavior)
    {
        if (column.HasSubtotal)
        {
            table.SetStartRow(2);
        }

        if (!String.IsNullOrEmpty(globalBehavior.NewLineSeparator) && !(column.DataType == ExcelModelDefs.ExcelDataTypes.Sheetlink))
        {
            column.IsMultilined = true;
            column.NewLineSeparator = globalBehavior.NewLineSeparator;
        }
        else if (!String.IsNullOrEmpty(column.NewLineSeparator) && !(column.DataType == ExcelModelDefs.ExcelDataTypes.Sheetlink))
        {
            column.IsMultilined = true;
        }

        if (column.DataType == ExcelModelDefs.ExcelDataTypes.DateTime && String.IsNullOrEmpty(column.DataFormat))
        {
            if (!String.IsNullOrEmpty(globalBehavior.Date.Format))
            {
                column.DataFormat = globalBehavior.Date.Format;
            }
            else
            {
                throw new Exception("No Date Format found.");
            }
        }

        if (column.DataType == ExcelModelDefs.ExcelDataTypes.AutoDetect)
        {
            DetermineDataType(column, globalBehavior, column.IsMultilined);
        }
    }

    private ExcelModelDefs.ExcelDataTypes DetermineDataType(ExcelColumnModel column, ExcelGlobalBehavior behavior, bool isMultilined)
    {
        if (ExcelHyperlinkParser.IsHyperlink(column, behavior.Hyperlink.IsHtml))
        {
            column.DataType = ExcelModelDefs.ExcelDataTypes.Hyperlink;
            return column.DataType;
        }

        // Multilined columns should not detect dates
        if (!isMultilined)
        {
            if (_excelDate.IsDate(column, behavior.Date.Format))
            {
                column.DataType = ExcelModelDefs.ExcelDataTypes.DateTime;
                column.DataFormat = behavior.Date.Format;
                return column.DataType;
            }
        }

        column.DataType = ExcelModelDefs.ExcelDataTypes.Text;

        return column.DataType;
    }
}
