namespace TableExporter.PersistModel;

/// <summary>
/// Class describing Hyperlink Data. Needed to write and link all components inside a XML file.
/// </summary>
public class ExcelSheetlink
{
    public ExcelSheetlink()
    {
    }

    public ExcelSheetlink(string sheetlink)
    {
        Sheetlink = sheetlink;
    }

    /// <summary>
    /// The actual link, the webpath
    /// </summary>
    public string Sheetlink { get; set; } = String.Empty;
}
