namespace rbkApiModules.Utilities.Excel.PersistModel;

/// <summary>
/// Class describing Hyperlink Data. Needed to write and link all components inside a XML file.
/// </summary>
public class ExcelHyperlink
{
    public ExcelHyperlink()
    {
    }

    public ExcelHyperlink(string hyperlink)
    {
        Hyperlink = hyperlink;
    }

    /// <summary>
    /// The actual link, the webpath
    /// </summary>
    public string Hyperlink { get; set; } = String.Empty;

    /// <summary>
    /// Index of the hyperlink that has to be created at worksheet creation time and needed later on the sheet data section
    /// </summary>
    public string LinkId { get; set; } = String.Empty;
}
