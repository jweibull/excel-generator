using System.Globalization;
using System.Text.RegularExpressions;
using static ExcelGenerator.ExcelDefs.ExcelModelDefs;

namespace ExcelGenerator.Excel;

/// <summary>
/// Helper class that parses data into dictionaries that can be stored on excel files as indexes.
/// </summary>
internal class DataParser
{
    private readonly Dictionary<string, double> _oleADates;
    private readonly Dictionary<string, string> _sharedStringsToIndex;
    private int _sharedStringsCount;
    private int _sharedStringsUniqueCount;

    internal DataParser()
    {
        _sharedStringsToIndex = new Dictionary<string, string>();
        _oleADates = new Dictionary<string, double>();
    }

    internal int SharedStringsCount { get { return _sharedStringsCount; } }

    internal int SharedStringsUniqueCount { get { return _sharedStringsUniqueCount; } }

    internal Dictionary<string, double> OleADates { get { return _oleADates; } }

    internal Dictionary <string, string> SharedStringsToIndex { get { return _sharedStringsToIndex; } }

    internal void PrepareData(ExcelWorkbookModel workbookModel)
    {
        _sharedStringsCount = 0;
        foreach (var table in workbookModel.Tables)
        {
            if (table.Header.Data.Length != table.Columns.Length)
            {
                throw new Exception("Length of Headers and columns must match");
            }
            AddToSharedStringDictionary(table.Header.Data);
            foreach (var column in table.Columns)
            {

                if (column.DataType == ExcelDataTypes.DataType.AutoDetect)
                {
                    // Check for either Dates or Hyperlinks on data colunms
                    PrepareAutodetectData(column, table.IsMultilined);
                }
                else
                {
                    // If not autodetect prepare regular types
                    PrepareDeclaredTypeData(column, table.IsMultilined);
                }

            }
        }
        _sharedStringsUniqueCount = _sharedStringsToIndex.Count;
    }

    private void AddToDatetimeToDictionary(string[] dates, string dataFormat)
    {
        var index = 0;
        DateTime date;
        while (index < dates.Length)
        {
            if (!_oleADates.ContainsKey(dates[index]))
            {
                if (DateTime.TryParseExact(dates[index], dataFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                {
                    _oleADates.Add(dates[index], date.ToOADate());
                }
            }
            index++;
        }
    }

    private void AddToSharedStringDictionary(string[] sharedStrings)
    {
        var count = 0;
        for (int itemIndex = 0; itemIndex < sharedStrings.Length; itemIndex++)
        {
            sharedStrings[itemIndex] = Regex.Replace(sharedStrings[itemIndex], "<br>", Environment.NewLine, RegexOptions.IgnoreCase);
            if (_sharedStringsToIndex.ContainsKey(sharedStrings[itemIndex]))
            {
                count++;
            }
            else
            {
                count++;
                _sharedStringsToIndex.Add(sharedStrings[itemIndex], _sharedStringsToIndex.Count().ToString());
            }
        }
        _sharedStringsCount += count;
    }

    private void PrepareDeclaredTypeData(ExcelColumnModel column, bool isMultilined)
    {
        if (column.DataType == ExcelDataTypes.DataType.Text)
        {
            AddToSharedStringDictionary(column.Data);
        }
        else if (column.DataType == ExcelDataTypes.DataType.HyperLink)
        {
            var linkSample = column.Data.FirstOrDefault(x => !string.IsNullOrEmpty(x.Trim()) && x.Contains("href"));
            if (linkSample != null)
            {
                if (isMultilined)
                {
                    PrepareMultilinedHrefHyperlinks(column);
                }
                else
                {
                    PrepareHrefHyperlinks(column);
                }
            }
            else
            {
                if (isMultilined)
                {
                    PrepareMultilinedRegularHyperlinks(column);
                }
                else
                {
                    PrepareRegularHyperlinks(column);
                }
            }
            AddToSharedStringDictionary(column.Data);
        }
        else if (column.DataType == ExcelDataTypes.DataType.DateTime)
        {
            if (string.IsNullOrEmpty(column.DataFormat.Trim()))
            {
                throw new Exception("Data format should not be empty when using the DateTime type");
            }
            AddToDatetimeToDictionary(column.Data, column.DataFormat);
        }
    }

    private void PrepareAutodetectData(ExcelColumnModel column, bool isMultilined)
    {
        var linkSample = column.Data.FirstOrDefault(x => !string.IsNullOrEmpty(x.Trim()) && (x.Contains("href") || x.StartsWith("http://") || x.StartsWith("https://")));
        if (linkSample != null)
        {
            if (isMultilined)
            {
                PrepareMultilinedAutodetectedHyperlinks(column, linkSample);
            }
            else
            {
                PrepareAutodetectedHyperlinks(column, linkSample);
            }
            AddToSharedStringDictionary(column.Data);
        }
        else if (DateTime.TryParseExact(
            column.Data.FirstOrDefault(x => !string.IsNullOrEmpty(x.Trim())),
            CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern.ToString(),
            CultureInfo.InvariantCulture,
            DateTimeStyles.None,
            out var date))
        {
            column.DataType = ExcelDataTypes.DataType.DateTime;
            AddToDatetimeToDictionary(column.Data, column.DataFormat);
        }
        else
        {
            column.DataType = ExcelDataTypes.DataType.Text;
            AddToSharedStringDictionary(column.Data);
        }
    }

    private void PrepareMultilinedAutodetectedHyperlinks(ExcelColumnModel column, string linkSample)
    {
        if (linkSample.Contains("href"))
        {
            PrepareMultilinedHrefHyperlinks(column);
        }
        else
        {
            PrepareMultilinedRegularHyperlinks(column);
        }
    }

    private void PrepareMultilinedRegularHyperlinks(ExcelColumnModel column)
    {
        column.DataType = ExcelDataTypes.DataType.Text;
        var data = column.Data;
        for (int itemIndex = 0; itemIndex < data.Length; itemIndex++)
        {
            data[itemIndex] = Regex.Replace(data[itemIndex], "<br>", Environment.NewLine, RegexOptions.IgnoreCase);
        }
    }

    private void PrepareMultilinedHrefHyperlinks(ExcelColumnModel column)
    {
        column.DataType = ExcelDataTypes.DataType.Text;
        var data = column.Data;
        for (int itemIndex = 0; itemIndex < data.Length; itemIndex++)
        {
            string hyperlink = data[itemIndex];
            hyperlink = Regex.Replace(hyperlink, "<br>", Environment.NewLine, RegexOptions.IgnoreCase);
            var matches = Regex.Matches(hyperlink, @"<a.*?href=[\'""]?([^\'"" >]+).*?<\/a>", RegexOptions.IgnoreCase);

            foreach (Match match in matches)
            {
                hyperlink = hyperlink.Replace(match.Value, match.Groups[1].Value);
            }
            data[itemIndex] = hyperlink;
        }
    }

    private void PrepareAutodetectedHyperlinks(ExcelColumnModel column, string linkSample)
    {
        column.DataType = ExcelDataTypes.DataType.HyperLink;
        if (linkSample.Contains("href"))
        {
            PrepareHrefHyperlinks(column);
        }
        else
        {
            PrepareRegularHyperlinks(column);
        }
    }

    private void PrepareRegularHyperlinks(ExcelColumnModel column)
    {
        var data = column.Data;
        var hyperlinks = new List<ExcelHyperlink>();
        for (int itemIndex = 0; itemIndex < data.Length; itemIndex++)
        {
            if (!string.IsNullOrEmpty(data[itemIndex].Trim()))
            {
                data[itemIndex] = Regex.Replace(data[itemIndex], "<br>", Environment.NewLine, RegexOptions.IgnoreCase);
                hyperlinks.Add(new ExcelHyperlink() { Hyperlink = data[itemIndex] });
            }
        }
        column.AddHyperLinkData(hyperlinks.ToArray());
    }

    private void PrepareHrefHyperlinks(ExcelColumnModel column)
    {
        var data = column.Data;
        var hyperlinks = new List<ExcelHyperlink>();

        for (int itemIndex = 0; itemIndex < data.Length; itemIndex++)
        {
            if (!string.IsNullOrEmpty(data[itemIndex].Trim()))
            {
                string hyperlink = data[itemIndex];
                hyperlink = Regex.Replace(hyperlink, "<br>", Environment.NewLine, RegexOptions.IgnoreCase);

                string text = Regex.Replace(hyperlink, "(<[a|A][^>]*>|)", "");

                var matches = Regex.Matches(hyperlink, @"<a.*?href=[\'""]?([^\'"" >]+).*?<\/a>", RegexOptions.IgnoreCase);

                foreach (Match match in matches)
                {
                    hyperlink = hyperlink.Replace(match.Value, match.Groups[1].Value);
                }
                data[itemIndex] = text;
                hyperlinks.Add(new ExcelHyperlink() { Hyperlink = hyperlink });
            }
        }
        column.AddHyperLinkData(hyperlinks.ToArray());
    }
}
