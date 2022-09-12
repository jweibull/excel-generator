using static ExcelGenerator.ExcelDefs.ExcelModelDefs;

namespace ExcelGenerator.Excel;

/// <summary>
/// Class combining Fonts, Data Type and formats to be applied to Cells or Columns.
/// Each styling is added in dxfs section in the stylesheet xml file combining fills, fonts, data formating, etc.
/// For each diferent combination a TAG is created in the file
/// </summary>
public class ExcelStyleFormat
{
    public ExcelStyleFormat(ExcelDataTypes.DataType dataType, string dataFormat, int styleIndex, int cellStyleIndex, ExcelFontDetail fontDetail)
    {
        DataType = dataType;
        StyleIndex = styleIndex;
        Format = dataFormat;
        CellStyleIndex = cellStyleIndex;
        FontDetail = fontDetail;
    }

    /// <summary>
    /// Datatype for the CellXfs style section
    /// </summary>
    public ExcelDataTypes.DataType DataType { get; set; } = ExcelDataTypes.DataType.Text;

    /// <summary>
    /// Possible format string for number or date data types
    /// </summary>
    public string Format { get; set; }

    /// <summary>
    /// Index of the CellFxs List inside the XML stylestable file
    /// </summary>
    public int StyleIndex{ get; set; } = -1;

    /// <summary>
    /// Index of the CellStyle List inside the XML stylestable file. 
    /// Regularly will be 1 for Hyperlink and 0 for everything else. 
    /// </summary>
    public int CellStyleIndex { get; set; } = -1;

    /// <summary>
    /// Font added 
    /// </summary>
    public ExcelFontDetail FontDetail { get; set; }
}
