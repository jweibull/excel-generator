using System.Collections.Specialized;
using System.Net;
using System.Web;

namespace TableExporter.DataPreparation;

/// <summary>
/// Helper class that parses data into dictionaries that can be stored on excel files as indexes.
/// </summary>
internal static class ExcelHyperlinkParser
{
    internal static void PrepareHyperlinks(ExcelColumnModel column, bool isHtml, bool isMultilined, string newLineString)
    {
        if (isHtml)
        {
            if (isMultilined)
            {
                PrepareMultilinedHrefHyperlinks(column, newLineString);
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
                PrepareMultilinedRegularHyperlinks(column, newLineString);
            }
            else
            {
                PrepareRegularHyperlinks(column);
            }
        }
    }

    internal static bool IsHyperlink(ExcelColumnModel column, bool isHtml)
    {
        string linkSample = null;
        if (isHtml == true)
        {
            linkSample = column.Data.FirstOrDefault(x => !String.IsNullOrEmpty(x) && x.Contains("href") && x.Contains("http"));
        }
        else
        {
            linkSample = column.Data.FirstOrDefault(x => !String.IsNullOrEmpty(x) && x.Contains("http"));
        }

        if (linkSample != null)
        {
            return true;
        }

        return false;
    }

    private static void PrepareMultilinedRegularHyperlinks(ExcelColumnModel column, string newLineString)
    {
        column.DataType = ExcelModelDefs.ExcelDataTypes.Text;
        var data = column.Data;

        if (!String.IsNullOrEmpty(newLineString))
        {
            for (int itemIndex = 0; itemIndex < data.Length; itemIndex++)
            {
                data[itemIndex] = Regex.Replace(data[itemIndex], newLineString, Environment.NewLine, RegexOptions.IgnoreCase);

                var hyperlinks = data[itemIndex].Split(Environment.NewLine);

                foreach (var hyperlink in hyperlinks)
                {
                    var encodedUrl = EncodeHyperlink(hyperlink.Trim());

                    data[itemIndex] = data[itemIndex].Replace(hyperlink, encodedUrl);
                }
            }
        }
    }

    private static void PrepareMultilinedHrefHyperlinks(ExcelColumnModel column, string newLineString)
    {
        column.DataType = ExcelModelDefs.ExcelDataTypes.Text;

        var data = column.Data;

        var hasNewLineSeparator = !String.IsNullOrEmpty(newLineString);

        for (int itemIndex = 0; itemIndex < data.Length; itemIndex++)
        {
            var hyperlink = data[itemIndex];

            if (hasNewLineSeparator)
            {
                hyperlink = Regex.Replace(hyperlink, newLineString, Environment.NewLine, RegexOptions.IgnoreCase);
            }

            var matches = Regex.Matches(hyperlink, @"<a.*?href=[\'""]([^\'""]+).*?<\/a>", RegexOptions.IgnoreCase);

            foreach (Match match in matches.Cast<Match>())
            {
                var encodedUrl = EncodeHyperlink(match.Groups[1].Value);

                hyperlink = hyperlink.Replace(match.Value, encodedUrl);
            }

            data[itemIndex] = hyperlink;
        }
    }

    private static void PrepareRegularHyperlinks(ExcelColumnModel column)
    {
        var data = column.Data;
        var hyperlinks = new List<ExcelHyperlink>();
        for (int itemIndex = 0; itemIndex < data.Length; itemIndex++)
        {
            var excelHyperlink = new ExcelHyperlink() { Hyperlink = String.Empty };

            if (!String.IsNullOrEmpty(data[itemIndex]))
            {
                excelHyperlink.Hyperlink = EncodeHyperlink(data[itemIndex]);
            }

            hyperlinks.Add(excelHyperlink);
        }

        column.AddHyperLinkData(hyperlinks.ToArray());
    }

    private static void PrepareHrefHyperlinks(ExcelColumnModel column)
    {
        var data = column.Data;
        var hyperlinks = new List<ExcelHyperlink>();

        for (int itemIndex = 0; itemIndex < data.Length; itemIndex++)
        {
            var excelHyperlink = new ExcelHyperlink() { Hyperlink = String.Empty };
            if (!String.IsNullOrEmpty(data[itemIndex]))
            {
                var hyperlink = data[itemIndex];

                string text = Regex.Replace(hyperlink, "(</?[a|A][^>]*>|)", "");

                var matches = Regex.Matches(hyperlink, @"<a.*?href=[\'""]([^\'""]+).*?<\/a>", RegexOptions.IgnoreCase);

                foreach (Match match in matches.Cast<Match>())
                {
                    var encodedUrl = EncodeHyperlink(match.Groups[1].Value);

                    hyperlink = hyperlink.Replace(match.Value, encodedUrl);
                }

                data[itemIndex] = text;
                excelHyperlink.Hyperlink = hyperlink;
            }

            hyperlinks.Add(excelHyperlink);
        }

        column.AddHyperLinkData(hyperlinks.ToArray());
    }

    /// <summary>
    /// This method cleans the URI/Hyperlink by encoding spaces and other invalid URI characters and returning a sanitized URI string.
    /// </summary>
    private static string EncodeHyperlink(string hyperlink)
    {
        // Parse the URL and Query String
        var uri = new Uri(hyperlink);

        var queryParts = HttpUtility.ParseQueryString(uri.Query);

        // Encode the query string parts
        NameValueCollection encodedQueryParts = new NameValueCollection();
        foreach (string key in queryParts)
        {
            encodedQueryParts.Add(key, WebUtility.UrlEncode(queryParts[key]));
        }

        // Reconstruct the encoded query string
        StringBuilder encodedQueryString = new('?');
        foreach (string key in encodedQueryParts)
        {
            //Only after the first QueryString has been inserted
            if (encodedQueryString.Length > 1)
            {
                encodedQueryString.Append('&');
            }
            // Append the encoded key-value pair to the query string
            encodedQueryString.Append($"{key}={encodedQueryParts[key]}");
        }

        // Reconstruct and return the full URL
        return $"{uri.Scheme}://{uri.Host}{uri.AbsolutePath}{encodedQueryString}";
    }
}