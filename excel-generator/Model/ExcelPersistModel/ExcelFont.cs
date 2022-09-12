using static ExcelGenerator.ExcelDefs.ExcelModelDefs;

namespace ExcelGenerator.Excel;

/// <summary>
/// Class describing a Font as it is written inside the XML file.
/// </summary>
public class ExcelFontDetail
{ 
    public ExcelFontDetail(string fontName, uint fontSize, uint fontFamily, uint theme, int fontIndex)
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

    public static ExcelFontDetail GetFontStyles(ExcelFonts.FontType font, int fontIndex, int? fontSize)
    {
        var size = fontSize == null ? 11 : fontSize.Value;
        
        switch ((int)font)
        {
            case 0: return new ExcelFontDetail("Arial", (UInt32)size, 2, 1, fontIndex);
            case 1: return new ExcelFontDetail("Arial Bold", (UInt32)size, 2, 1, fontIndex);
            case 2: return new ExcelFontDetail("Arial Narrow", (UInt32)size, 2, 1, fontIndex);
            case 3: return new ExcelFontDetail("Calibri", (UInt32)size, 2, 1, fontIndex);
            case 4: return new ExcelFontDetail("Calibri Light", (UInt32)size, 2, 1, fontIndex);
            case 5: return new ExcelFontDetail("Courrier New", (UInt32)size, 1, 1, fontIndex);
            case 6: return new ExcelFontDetail("Times New Roman", (UInt32)size, 2, 1, fontIndex);
            case 7: return new ExcelFontDetail("Georgia", (UInt32)size, 3, 1, fontIndex);
            default:
                return new ExcelFontDetail("Calibri", (UInt32)size, 2, 1, fontIndex);
        }
    }

}
