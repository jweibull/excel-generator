using rbkApiModules.Utilities.Excel.InputModel;

namespace rbkApiModules.Utilities.Excel;

/// <summary>
/// The Main Workbook Model container. Holds all data and metadata.
/// </summary>
public class ModelData
{
    public ExcelWorkbookModel WorkbookModel { get; set; } = new ExcelWorkbookModel();
}
