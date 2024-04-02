namespace TableExporter;

/// <summary>
/// Helper class that parses data into dictionaries that can be stored on excel files as indexes.
/// </summary>
internal class ExcelSheetlinkParser
{
    internal void PrepareSheetlinks(ExcelColumnModel column, ExcelBaseSheetModel[] allSheets, bool isMultilined)
    {
        
        if (isMultilined)
        {
            throw new Exception("Multilined columns cannot have links between tabs");
        }
    
        PrepareSheetlinks(column, allSheets);
    }
    
    private void PrepareSheetlinks(ExcelColumnModel column, ExcelBaseSheetModel[] allSheets)
    {
        var data = column.Data;
        var sheetlinks = new List<ExcelSheetlink>();
        for (int itemIndex = 0; itemIndex < data.Length; itemIndex++)
        {
            if (int.TryParse(data[itemIndex], out var tabIndex))
            {
                data[itemIndex] = allSheets[tabIndex].Name;
                sheetlinks.Add(new ExcelSheetlink() { Sheetlink = $"'{data[itemIndex]}'!A1" });
            }
            else
            {
                sheetlinks.Add(new ExcelSheetlink());
            }
        }
        column.AddSheetLinkData(sheetlinks.ToArray());
    }
}
