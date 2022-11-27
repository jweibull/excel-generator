using System.Collections.Generic;
using System.Linq;

namespace rbkApiModules.Utilities.Excel.InputModel;

/// <summary>
/// The Main Workbook Model container. Holds all data and metadata.
/// </summary>
public class ExcelWorkbookModel
{
    /// <summary>
    /// Name of the excel file containing all the spreadsheets to be outputed
    /// </summary>
    public string FileName { get; set; } = "ExcelFile.xlsx";
    
    /// <summary>
    /// Authoring metadata such as Author name, Date created, company and comments.
    /// </summary>
    public ExcelWorkbookMetadata AuthoringMetadata { get; set; } = new ExcelWorkbookMetadata();

    /// <summary>
    /// This class must contain all rules needed for finding specific data types when autodetect is true for a column.
    /// </summary>
    public ExcelGlobalBehavior GlobalColumnBehavior { get; set; } = new ExcelGlobalBehavior();

    /// <summary>
    /// The data to generate a watermark image
    /// </summary>
    public Watermark? Watermark { get; set; } = null;
    
    /// <summary>
    /// List of all spreadsheets for this workbook, with tabular data and styling.
    /// </summary>
    public List<ExcelTableSheetModel> Tables { get; set; } = new List<ExcelTableSheetModel>(); 

    /// <summary>
    /// List of all plot sheets for this workbook, with plot, their data and styling.
    /// </summary>
    public List<ExcelChartSheetModel> Charts { get; set; } = new List<ExcelChartSheetModel>();
    
    public ExcelBaseSheetModel[] AllSheets
    {
        get
        {
            var allSheets = new List<ExcelBaseSheetModel>();

            if (Tables != null)
            {
                allSheets.AddRange(Tables);
            }
            if (Charts != null)
            {
                allSheets.AddRange(Charts);
            }

            return allSheets.OrderBy(x => x.TabIndex).ToArray();
        }
    }
}
