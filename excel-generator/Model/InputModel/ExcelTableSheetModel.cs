using TableExporter.Configurations;

namespace TableExporter.InputModel;

/// <summary>
/// Class representing a single spreadsheet, holding table data, inside an excel workbook.
/// </summary>
public class ExcelTableSheetModel: ExcelBaseSheetModel
{
    /// <summary>
    /// The header data and styling container 
    /// </summary>
    public ExcelHeaderModel Header { get; set; } = new ExcelHeaderModel();

    /// <summary>
    /// A list of all columns and their data/styling
    /// </summary>
    public List<ExcelColumnModel> Columns { get; set; } = new List<ExcelColumnModel>();

    /// <summary>
    /// If diferent from "None", applies a theme from excel's standard theme list to this spreadsheet
    /// </summary>
    public ExcelModelDefs.ExcelThemes Theme { get; set; } = ExcelModelDefs.ExcelThemes.None;


    #region Helper fields and methods

    /// <summary>
    /// Internal work variable that defines the start row if a column contains a subtotal row
    /// </summary>
    public int StartRow { get; private set; } = 1;

    internal void SetStartRow(int startRow)
    {
        StartRow = startRow;
    }

    #endregion
}
