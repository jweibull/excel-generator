using static ExcelGenerator.ExcelDefs.ExcelModelDefs;

namespace ExcelGenerator.Excel;

/// <summary>
/// Class describing a Font as it is written inside the XML file.
/// </summary>
public class ExcelNumFormat
{ 
    public ExcelNumFormat(string fontName, uint fontSize, uint fontFamily, uint theme, int fontIndex)
    {
        FontName = fontName;
        FontSize = fontSize;
        FontFamily = fontFamily;
        Theme = theme;
        FontIndex = fontIndex;
    }

    /// <summary>
    /// FontName as in Excel font list
    /// </summary>
    public string FontName { get; set; } = "Calibri";

    /// <summary>
    /// Font Size in Points as in Excel font size dropdown
    /// </summary>
    public UInt32 FontSize { get; set; } = 11;

    /// <summary>
    /// Font family number as defined by Excel Microsoft Arial, Calibri, Times are family 2
    /// </summary>
    public UInt32 FontFamily { get; set; } = 2;

    /// <summary>
    /// Theme for custom microsoft themes (Not table style theme) This is a diferent theme
    /// </summary>
    public UInt32 Theme { get; set; } = 1;

    /// <summary>
    /// Index of the Font inside the XML stylestable file
    /// </summary>
    public int FontIndex { get; set; } = -1;
}
