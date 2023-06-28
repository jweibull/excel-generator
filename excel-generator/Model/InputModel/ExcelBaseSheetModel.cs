using rbkApiModules.Utilities.Excel.Configurations;

namespace rbkApiModules.Utilities.Excel.InputModel;

/// <summary>
/// Base sheet definition with tab name and color and the type of data to be displayed. Ex: Table Data, Charts, etc.
/// </summary>
public class ExcelBaseSheetModel
{
    /// <summary>
    /// Spreasheet tab name
    /// </summary>
    public string Name { get; set; } = String.Empty;
    /// <summary>
    /// Sets the spreadsheet's tab background color. By default, it will not set a background color.
    /// Expects an 8 characters Hexadecimal ARGB string pattern without the # e.g. "FFFF0000" for solid Red.
    /// "FF00FF00" for solid green and "FF0000FF" for solid blue.
    /// </summary>
    public string TabColor { get; set; } = String.Empty;
    /// <summary>
    /// Data type to be exhibited on the sheet tab. Ex: Table Data, Plots, etc.
    /// </summary>
    public ExcelModelDefs.ExcelSheetTypes SheetType { get; set; } = ExcelModelDefs.ExcelSheetTypes.Table;
    /// <summary>
    /// Tab order in which this sheet should be created 
    /// </summary>
    public int TabIndex { get; set; }
}
