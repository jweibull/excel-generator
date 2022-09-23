using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Text.RegularExpressions;
using static ExcelGenerator.ExcelDefs.ExcelModelDefs;

namespace ExcelGenerator.Excel;

/// <summary>
/// Helper class that parses data into dictionaries that can be stored on excel files as indexes.
/// </summary>
internal class ExcelDate
{
    private readonly Dictionary<string, string> _oleADates;
        
    internal ExcelDate()
    {
        _oleADates = new Dictionary<string, string>();
    }

    internal string DateFormat { get; set; } = string.Empty;

    internal Dictionary<string, string> OleADates { get { return _oleADates; } }

    internal bool IsDate(ExcelColumnModel column, string format)
    {
        return false;
    }
    
    internal string GetValue(string key)
    {
        if (_oleADates.TryGetValue(key, out var oleADate))
        {
            return oleADate;
        }

        return string.Empty;
    }

    internal void AddToDatetimeToDictionary(string[] dates, string dataFormat)
    {
        var index = 0;
        DateTime date;
        while (index < dates.Length)
        {
            if (!_oleADates.ContainsKey(dates[index]))
            {
                if (DateTime.TryParseExact(dates[index], dataFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                {
                    _oleADates.Add(dates[index], date.ToOADate().ToString());
                }
            }
            index++;
        }
    }

    
}
