using System.Linq;

namespace rbkApiModules.Utilities.Excel;

/// <summary>
/// Helper class that parses data into dictionaries that can be stored on excel files as indexes.
/// </summary>
internal class DataParser
{
    private ExcelDate _excelDate;
    private ExcelHyperlinkParser _hyperlinkParser;
    private ExcelSheetlinkParser _sheetlinkParser;
    private ExcelSharedString _sharedString;
    
    internal DataParser()
    {
        _excelDate = new ExcelDate();
        _hyperlinkParser = new ExcelHyperlinkParser();
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

    internal string GetValue(ExcelModelDefs.ExcelDataTypes.DataType type, string key)
    {
        switch (type)
        {
            case ExcelModelDefs.ExcelDataTypes.DataType.Text:
            case ExcelModelDefs.ExcelDataTypes.DataType.Hyperlink:
            case ExcelModelDefs.ExcelDataTypes.DataType.Sheetlink:
                return _sharedString.GetValue(key);

            case ExcelModelDefs.ExcelDataTypes.DataType.DateTime:
                return _excelDate.GetValue(key);

            default:
                return key;
        }
    }

    internal void PrepareData(ExcelWorkbookModel workbookModel)
    {
        foreach (var table in workbookModel.Tables)
        {
            if (table.Header.Data.Length != table.Columns.Length)
            {
                throw new Exception("Length of Headers and number of columns must match");
            }

            _sharedString.AddToSharedStringDictionary(table.Header.Data, false, String.Empty);
            
            foreach (var column in table.Columns)
            {
                SetupColumn(table, column, workbookModel.GlobalColumnBehavior);

                switch (column.DataType)
                {
                    case ExcelModelDefs.ExcelDataTypes.DataType.DateTime:
                        _excelDate.AddToDatetimeToDictionary(column.Data, column.DataFormat);
                        break;

                    case ExcelModelDefs.ExcelDataTypes.DataType.Hyperlink:
                        _hyperlinkParser.PrepareHyperlinks(column, workbookModel.GlobalColumnBehavior.Hyperlink.IsHtml, column.IsMultilined, column.NewLineString);
                        _sharedString.AddToSharedStringDictionary(column.Data, column.IsMultilined, column.NewLineString);
                        break;

                    case ExcelModelDefs.ExcelDataTypes.DataType.Sheetlink:
                        _sheetlinkParser.PrepareHyperlinks(column, workbookModel.AllSheets.ToArray(), column.IsMultilined);
                        _sharedString.AddToSharedStringDictionary(column.Data, column.IsMultilined, column.NewLineString);
                        break;

                    case ExcelModelDefs.ExcelDataTypes.DataType.Number:
                    case ExcelModelDefs.ExcelDataTypes.DataType.Text:
                    default:
                        _sharedString.AddToSharedStringDictionary(column.Data, column.IsMultilined, column.NewLineString);
                        break;
                }
            }
        }
    }

    private void SetupColumn(ExcelTableSheetModel table, ExcelColumnModel column, ExcelGlobalBehavior globalBehavior)
    {
        if (column.HasSubtotal)
        {
            table.SetStartRow(2);
        }

        if (!String.IsNullOrEmpty(column.NewLineString))
        {
            column.IsMultilined = true;
        }

        if (column.DataType == ExcelModelDefs.ExcelDataTypes.DataType.DateTime && String.IsNullOrEmpty(column.DataFormat))
        {
            if (!String.IsNullOrEmpty(globalBehavior.Date.Format))
            {
                column.DataFormat = globalBehavior.Date.Format;
            }
            else
            {
                throw new Exception("No Date Format found");
            }
        }

        if (column.DataType == ExcelModelDefs.ExcelDataTypes.DataType.AutoDetect)
        {
            DetermineDataType(column, globalBehavior, column.IsMultilined);
        }
    }

    private ExcelModelDefs.ExcelDataTypes.DataType DetermineDataType(ExcelColumnModel column, ExcelGlobalBehavior behavior, bool isMultilined)
    {
        if (_hyperlinkParser.IsHyperlink(column, behavior.Hyperlink.IsHtml))
        {
            column.DataType = ExcelModelDefs.ExcelDataTypes.DataType.Hyperlink;
            return column.DataType;
        }

        // Multilined columns should not detect dates
        if (!isMultilined)
        {
            if (_excelDate.IsDate(column, behavior.Date.Format))
            {
                column.DataType = ExcelModelDefs.ExcelDataTypes.DataType.DateTime;
                column.DataFormat = behavior.Date.Format;
                return column.DataType;
            }
        }

        column.DataType = ExcelModelDefs.ExcelDataTypes.DataType.Text;

        return column.DataType;
    }
}
