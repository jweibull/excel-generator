﻿

using static ExcelGenerator.ExcelDefs.ExcelModelDefs;

namespace ExcelGenerator.Excel;

/// <summary>
/// Class representing a full header row, data and styling.
/// </summary>
public class ExcelHeaderModel
{
    /// <summary>
    /// The data list to be displayed at the header row
    /// </summary>
    public string[] Data { get; set; }
    /// <summary>
    /// Styles to be applied to the header row
    /// </summary>
    public ExcelStyleClasses Style { get; set; }

    /// <summary>
    /// Reserved Quick access Key Built from font, fontSize, data type and format
    /// </summary>
    public string StyleKey { get; private set; } = "";

    public void AddStyleKey(string key)
    {
        StyleKey = key;
    }
}