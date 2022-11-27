using rbkApiModules.Utilities.Excel.Configurations;
using rbkApiModules.Utilities.Excel.InputModel;
using rbkApiModules.Utilities.Excel.PersistModel;

namespace rbkApiModules.Utilities.Excel.DataPreparation;

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

    internal void PrepareData(ExcelWorkbookModel workbookModel)
    {
        foreach (var table in workbookModel.Tables)
        {
            if (table.Header.Data.Length != table.Columns.Count())
            {
                throw new Exception("Length of Headers and number of columns must match");
            }

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
                        _hyperlinkParser.PrepareHyperlinks(column, workbookModel.GlobalColumnBehavior.Hyperlink.IsHtml, column.IsMultilined, column.NewLineString);
                        _sharedString.AddToSharedStringDictionary(column.Data, column.IsMultilined, column.NewLineString);
                        break;

                    case ExcelModelDefs.ExcelDataTypes.Sheetlink:
                        _sheetlinkParser.PrepareHyperlinks(column, workbookModel.AllSheets, column.IsMultilined);
                        _sharedString.AddToSharedStringDictionary(column.Data, column.IsMultilined, column.NewLineString);
                        break;

                    case ExcelModelDefs.ExcelDataTypes.Number:
                    case ExcelModelDefs.ExcelDataTypes.Text:
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
        else if (!String.IsNullOrEmpty(globalBehavior.Text.NewLineString))
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
        if (_hyperlinkParser.IsHyperlink(column, behavior.Hyperlink.IsHtml))
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
