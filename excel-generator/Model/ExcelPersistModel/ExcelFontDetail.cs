using static ExcelGenerator.ExcelDefs.ExcelModelDefs;

namespace ExcelGenerator.Excel;

/// <summary>
/// Class describing a Font as it is written inside the XML file.
/// </summary>
public class ExcelFontDetail
{ 
    public ExcelFontDetail(string fontName, UInt32 fontSize, Int32 fontFamily, UInt32 theme, UInt32 fontIndex)
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
    public Int32 FontFamily { get; set; } = 2;

    /// <summary>
    /// Theme for custom microsoft themes (Not table style theme) This is a diferent theme
    /// </summary>
    public UInt32 Theme { get; set; } = 1;

    /// <summary>
    /// Index of the Font inside the XML stylestable file
    /// </summary>
    public UInt32 FontIndex { get; set; }

    public static ExcelFontDetail GetFontStyles(ExcelFonts.FontType font, UInt32 fontIndex, int fontSize, int theme)
    {
        switch ((int)font)
        {
            case 0: return new ExcelFontDetail("Arial", (UInt32)fontSize, 2, (UInt32)theme, fontIndex);
            case 1: return new ExcelFontDetail("Arial Bold", (UInt32)fontSize, 2, (UInt32)theme, fontIndex);
            case 2: return new ExcelFontDetail("Arial Narrow", (UInt32)fontSize, 2, (UInt32)theme, fontIndex);
            case 3: return new ExcelFontDetail("Calibri", (UInt32)fontSize, 2, (UInt32)theme, fontIndex);
            case 4: return new ExcelFontDetail("Calibri Light", (UInt32)fontSize, 2, (UInt32)theme, fontIndex);
            case 5: return new ExcelFontDetail("Courrier New", (UInt32)fontSize, 1, (UInt32)theme, fontIndex);
            case 6: return new ExcelFontDetail("Times New Roman", (UInt32)fontSize, 2, (UInt32)theme, fontIndex);
            case 7: return new ExcelFontDetail("Georgia", (UInt32)fontSize, 3, (UInt32)theme, fontIndex);
            default:
                return new ExcelFontDetail("Calibri", (UInt32)fontSize, 2, (UInt32)theme, fontIndex);
        }
    }

}
