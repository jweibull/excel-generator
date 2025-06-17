using TableExporterApp;

public static class DataWithLinks
{
    // Sheet 1
    public static readonly string[] Sheet1HeaderData = new string[] { "Some data", "sheet links" };

    public static readonly string[] Sheet1Column1Data = new string[] { "Hyperlinks Sheet", "Number Sheet" };
    public static readonly string[] Sheet1Column2Data = new string[] { "1", "2" };

    // Sheet 2
    public static readonly string[] Sheet2HeaderData = new string[] { "MaxW20", "Regular Hyperlink", "Multiline Hyperlinks" };

    public static readonly string[] Sheet2Column1Data = new string[] { "string muito muito muito muito longa", "b", "Estou na Linha 1<br>Pulei pra linha 2", "d", "e",
                "string muito muito muito muito longa", "b", "c", "d", "e", "string muito muito muito muito longa", "b", "c", "d", "e",
                "string muito muito muito muito longa", "b", "c", "d", "e"
        };
    public static readonly string[] Sheet2Column2Data = new string[] { "<a href=\"http://npaa7587.petrobras.biz/WebFacil3/resultBrTecEng.aspx?strCodDoc=ADV  -54236009-SP- \">ADV  -54236009-SP- </a>",
                "<a href=\"http://www.google.com\">Google</a>", "<a href=\"http://www.yahoo.com\">Yahoo</a>", "", "<a href=\"http://www.microsoft.com\">Microsoft</a>",
                "<a href=\"http://www.apple.com\">Apple</a>", "<a href=\"http://www.github.com\">Github</a>", "<a href=\"http://www.google.com\">Google</a>",
                "<a href=\"http://www.yahoo.com\">Yahoo</a>", "<a href=\"http://www.microsoft.com\">Microsoft</a>", "<a href=\"http://www.apple.com\">Apple</a>",
                "<a href=\"http://www.github.com\">Github</a>", "<a href=\"http://www.google.com\">Google</a>", "<a href=\"http://www.yahoo.com\">Yahoo</a>",
                "<a href=\"http://www.microsoft.com\">Microsoft</a>", "<a href=\"http://www.apple.com\">Apple</a>", "<a href=\"http://www.github.com\">Github</a>",
                "<a href=\"http://www.google.com\">Google</a>", "<a href=\"http://www.yahoo.com\">Yahoo</a>", "<a href=\"http://www.microsoft.com\">Microsoft</a>"
        };
    public static readonly string[] Sheet2Column3Data = new string[] { "<a href=\"http://npaa7587.petrobras.biz/WebFacil3/resultBrTecEng.aspx?strCodDoc=ADV  -54236009-SP- \">ADV  -54236009-SP- </a><br><a href=\"http://npaa7587.petrobras.biz/WebFacil3/resultBrTecEng.aspx?strCodDoc=BDV  -52416011-SP- \">BDV  -52416011-SP- </a>",
                "<a href=\"http://www.google.com\">Google</a>", "<a href=\"http://www.yahoo.com\">Yahoo</a>", "", "<a href=\"http://www.microsoft.com\">Microsoft</a>",
                "<a href=\"http://www.apple.com\">Apple</a>", "<a href=\"http://www.github.com\">Github</a>", "<a href=\"http://www.google.com\">Google</a>",
                "<a href=\"http://www.yahoo.com\">Yahoo</a>", "<a href=\"http://www.microsoft.com\">Microsoft</a>", "<a href=\"http://www.apple.com\">Apple</a>",
                "<a href=\"http://www.github.com\">Github</a>", "<a href=\"http://www.google.com\">Google</a>", "<a href=\"http://www.yahoo.com\">Yahoo</a>",
                "<a href=\"http://www.microsoft.com\">Microsoft</a>", "<a href=\"http://www.apple.com\">Apple</a>", "<a href=\"http://www.github.com\">Github</a>",
                "<a href=\"http://www.google.com\">Google</a>", "<a href=\"http://www.yahoo.com\">Yahoo</a>", "<a href=\"http://www.microsoft.com\">Microsoft</a>"
        };
    public static readonly string[] Sheet2Column4Data = new string[] { "<a href=\"http://npaa7587.petrobras.biz/WebFacil3/resultBrTecEng.aspx?strCodDoc=ADV  -54236009-SP- \">ADV  -54236009-SP- </a>",
                "<a href=\"http://www.google.com\">Google</a>", "<a href=\"http://www.yahoo.com\">Yahoo</a>", "", "<a href=\"http://www.microsoft.com\">Microsoft</a>",
                "<a href=\"http://www.apple.com\">Apple</a>", "<a href=\"http://www.github.com\">Github</a>", "<a href=\"http://www.google.com\">Google</a>",
                "<a href=\"http://www.yahoo.com\">Yahoo</a>", "<a href=\"http://www.microsoft.com\">Microsoft</a>", "<a href=\"http://www.apple.com\">Apple</a>",
                "<a href=\"http://www.github.com\">Github</a>", "<a href=\"http://www.google.com\">Google</a>", "<a href=\"http://www.yahoo.com\">Yahoo</a>",
                "<a href=\"http://www.microsoft.com\">Microsoft</a>", "<a href=\"http://www.apple.com\">Apple</a>", "<a href=\"http://www.github.com\">Github</a>",
                "<a href=\"http://www.google.com\">Google</a>", "<a href=\"http://www.yahoo.com\">Yahoo</a>", "<a href=\"http://www.microsoft.com\">Microsoft</a>"
        };


    // Sheet 3
    public static readonly string[] Sheet3HeaderData = new string[] { "Numbers", "Auto Sum Column", "Dates", "Dates Auto" };

    public static readonly string[] Sheet3Column1Data = new string[] { "100", "2000", "30", "40000", "500.54", "6", "7", "8.563", "9", "10", "20000.34", "0.237" };

    public static readonly string[] Sheet3Column2Data = new string[] { "10.10", "11.11", "9.76", "8.566", "30.45", "5.87", "6.34", "8.563", "9.23", "", "10.10", "11.11" };

    public static readonly string[] Sheet3Column3Data = new string[] { "12/10/1977", "13/10/1977", "14/10/1977", "15/10/1977", "16/10/1977", "", "13/10/1977", "14/10/1977", "15/10/1977", "16/10/1977",
                "12/11/1977", "13/11/1977"
        };
    public static readonly string[] Sheet3Column4Data = new string[] { "12/10/1977", "16/10/1977", "18/10/1977", "10/10/1977", "10/10/1977", "25/01/1965", "11/10/1977", "11/10/1977", "11/10/1977",
                "11/10/1977", "11/11/1977", "11/11/1977"
        };
}