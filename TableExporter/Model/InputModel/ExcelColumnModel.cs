﻿namespace TableExporter;

/// <summary>
/// Class with the data models and styling for a column data, to be displayed under a header title.
/// </summary>
public class ExcelColumnModel
{
    /// <summary>
    /// List of all data to be displayed on one column
    /// </summary>
    public string[] Data { get; set; } = new string[0];

    /// <summary>
    /// Styles to be applied to this column's data
    /// </summary>
    public ExcelStyleClasses Style { get; set; } = new ExcelStyleClasses();

    /// <summary>
    /// The data sets Data Type. Ex: "0" for Text, "1" for Number, "2" for DateTime, etc.
    /// </summary>
    public ExcelModelDefs.ExcelDataTypes DataType { get; set; } = ExcelModelDefs.ExcelDataTypes.Text;

    /// <summary>
    /// Data formating for a specific type. Ex: For Number Type: DataFormat = "0.00" Will format the number with 2 decimal precision.
    /// For the DateTime type: DataFormat = "dd/MM/yyyy" will format dates for the brazilian standard.
    /// </summary>
    public string DataFormat { get; set; } = String.Empty;

    /// <summary>
    /// Excel Max Column Width in POINT units (Not Pixels). If this property is set to a positive value, it will limit the Column to the MaxWidth.
    /// Useful for large texts that often ocupy large portions the monitor's width.
    /// </summary>
    public int MaxWidth { get; set; } = -1;

    /// <summary>
    /// If this is set as true a subtotal line will be added to the first line to compute a filtered sum of the column values
    /// </summary>
    public bool HasSubtotal { get; set; } = false;

    /// <summary>
    /// If the table is multilined, this will invalidate using clicable hyperlinks, since there can only be one hyperlink per cell
    /// </summary>
    public bool IsMultilined { get; set; } = false;

    /// <summary>
    /// If a cell has multiple lines, then NewLineString must define the string which separates the lines: "\n", <br>, etc.
    /// If this is empty, then the cell doesn't have multiple lines
    /// </summary>
    public string NewLineSeparator { get; set; } = String.Empty;

    #region Build Helper Section

    /// <summary>
    /// Reserved Quick access Key Built from font, fontSize, data type and format
    /// </summary>
    internal string StyleKey { get; private set; } = String.Empty;

    /// <summary>
    /// Builder helper Field to distinguish the hyperlink Ids related to this column
    /// </summary>
    internal ExcelHyperlink[] HyperLinkData { get; private set; } = new ExcelHyperlink[0];

    internal void AddHyperLinkData(ExcelHyperlink[] hyperLinkData)
    {
        HyperLinkData = hyperLinkData;
    }

    /// <summary>
    /// Builder helper Field to distinguish the sheetlink Ids related to this column
    /// </summary>
    internal ExcelSheetlink[] SheetlinkData { get; private set; } = new ExcelSheetlink[0];

    internal void AddSheetLinkData(ExcelSheetlink[] sheetlinkData)
    {
        SheetlinkData = sheetlinkData;
    }

    /// <summary>
    /// Reserved MaxDataLength in characters for a given column sampled from n
    /// </summary>
    internal int GetMaxDataLength(bool isMultilined, int numSamples) 
    {
        numSamples = numSamples <= Data.Length ? numSamples : Data.Length;
        var samples = Data.Take(numSamples).ToList();
        
        //If string has multiple lines only return the size the line with longest size
        if (isMultilined)
        {
            var maxSize = 0;
            
            for (int i = 0; i < samples.Count(); i++)
            {
                if (samples[i].Contains(Environment.NewLine))
                {
                    var splitted = samples[i].Split(Environment.NewLine);
                    maxSize = Math.Max(maxSize, splitted.Select(x => x.Length).Max());
                }
                else
                {
                    maxSize = Math.Max(maxSize, samples[i].Length);
                }
            }
            return maxSize;
        }

        // if not multiline, return the max length of the sample list
        return samples.Select(x => x.Length).Max();
    }

    internal void AddStyleKey(string key)
    {
        StyleKey = key;
    }

    #endregion
}
