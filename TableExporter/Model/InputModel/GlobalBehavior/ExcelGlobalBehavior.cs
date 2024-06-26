﻿namespace TableExporter;

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
    /// Globally sets wether columns have more than one line and which is the separator, for example the <br> tag or \n
    /// </summary>
    public string NewLineSeparator { get; set; } = String.Empty;
}

