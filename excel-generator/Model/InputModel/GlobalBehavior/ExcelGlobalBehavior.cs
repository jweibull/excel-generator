namespace rbkApiModules.Utilities.Excel.InputModel;

/// <summary>
/// Class describing the rules needed when auto detecting a data type on a column
/// </summary>
public class ExcelGlobalBehavior
{
    /// <summary>
    /// Global behaviors for dates.
    /// This behavior Will be overriden by column specific configurations.
    /// </summary>
    public ExcelDateGlobal Date { get; set; } = new ExcelDateGlobal();

    /// <summary>
    /// Global behaviors for hyperlinks. 
    /// This behavior Will be overriden by column specific configurations.
    /// </summary>
    public ExcelHyperlinkGlobal Hyperlink { get; set; } = new ExcelHyperlinkGlobal();

    /// <summary>
    /// If a cell has multiple lines, then NewLineString must define the string which separates the lines: "\n", <br>, etc.
    /// If this is empty, then the cell doesn't have multiple lines
    /// Even if this is not set, this behavior may be overriden by column specific configurations.
    /// </summary>
    public ExcelTextGlobal Text { get; set; } = new ExcelTextGlobal();
}

