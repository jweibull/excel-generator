using Newtonsoft.Json;

namespace rbkApiModules.Utilities.Excel;

public class Runner
{
    public void Run()
    {
        var saxLib = new SaxLib();

        var serializer = new JsonSerializer();

        ModelData? modelData;

        string inputPath = Directory.GetCurrentDirectory();
        inputPath = Path.Combine(inputPath, "..", "..", "..", "input", "excel.json");

        using (StreamReader sr = new StreamReader(inputPath))
        using (var jsonTextReader = new JsonTextReader(sr))
        {
            modelData = serializer.Deserialize<ModelData>(jsonTextReader);
        }

        if (modelData == null)
        {
            throw new Exception("Data input was null.");
        }

        var filename = GetNextFilename();

        var stream = saxLib.CreatePackage(modelData.WorkbookModel);

        using (var fileStream = File.Create(filename))
        {
            stream.Seek(0, SeekOrigin.Begin);
            stream.CopyTo(fileStream);
        }
        
    }

    public void RunMockData()
    {
        var sheet1HeaderData = new string[60];
        for (int i = 0; i < sheet1HeaderData.Length; i++) 
        {
            sheet1HeaderData[i] = "Lorem Ipsum";
        }

        var columnData = new string[200000];
        for (int i = 0;i < columnData.Length; i++)
        {
            columnData[i] = "Neque porro quisquam est qui dolorem ipsum quia dolor sit amet, consectetur, adipisci velit...";
        }

        var filename = GetNextFilename();

        var excel = ExcelWorkbookBuilder.StartWorkbook(Path.GetFileName(filename).Replace(Path.GetExtension(filename), ""))
            .WithAuthor("Excel Lib")
            .WithTitle("Excel Lib Test")
            .WithCompany("Excel Lib")
            .WithComments("Very nice comment goes here")
            .WithGlobalDateFormat("dd/MM/yyyy")
            .WithGlobalHtmlTagHyperlinks()
            .AddTableSheet("Summary")
                .WithTheme(Configurations.ExcelModelDefs.ExcelThemes.TableStyleLight1)
                .WithTabColor("FF222222")
                .AddHeader(sheet1HeaderData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.CalibriLight, 14, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .Workbook
            .Build();
    }

    public void RunFluent()
    {
        // Sheet 1
        var sheet1HeaderData = new string[] { "Some data", "sheet links" };

        var sheet1Column1Data = new string[] { "Hyperlinks Sheet", "Number Sheet" };
        var sheet1Column2Data = new string[] { "1", "2" };

        // Sheet 2
        var sheet2HeaderData = new string[] { "MaxW20", "Regular Hyperlink", "Multiline Hyperlinks" };

        var sheet2Column1Data = new string[] { "string muito muito muito muito longa", "b", "Estou na Linha 1<br>Pulei pra linha 2", "d", "e",
                "string muito muito muito muito longa", "b", "c", "d", "e", "string muito muito muito muito longa", "b", "c", "d", "e",
                "string muito muito muito muito longa", "b", "c", "d", "e"
        };
        var sheet2Column2Data = new string[] { "<a href=\"http://npaa7587.petrobras.biz/WebFacil3/resultBrTecEng.aspx?strCodDoc=ADV  -54236009-SP- \">ADV  -54236009-SP- </a>",
                "<a href=\"http://www.google.com\">Google</a>", "<a href=\"http://www.yahoo.com\">Yahoo</a>", "", "<a href=\"http://www.microsoft.com\">Microsoft</a>",
                "<a href=\"http://www.apple.com\">Apple</a>", "<a href=\"http://www.github.com\">Github</a>", "<a href=\"http://www.google.com\">Google</a>",
                "<a href=\"http://www.yahoo.com\">Yahoo</a>", "<a href=\"http://www.microsoft.com\">Microsoft</a>", "<a href=\"http://www.apple.com\">Apple</a>",
                "<a href=\"http://www.github.com\">Github</a>", "<a href=\"http://www.google.com\">Google</a>", "<a href=\"http://www.yahoo.com\">Yahoo</a>",
                "<a href=\"http://www.microsoft.com\">Microsoft</a>", "<a href=\"http://www.apple.com\">Apple</a>", "<a href=\"http://www.github.com\">Github</a>",
                "<a href=\"http://www.google.com\">Google</a>", "<a href=\"http://www.yahoo.com\">Yahoo</a>", "<a href=\"http://www.microsoft.com\">Microsoft</a>"
        };
        var sheet2Column3Data = new string[] { "<a href=\"http://npaa7587.petrobras.biz/WebFacil3/resultBrTecEng.aspx?strCodDoc=ADV  -54236009-SP- \">ADV  -54236009-SP- </a><br><a href=\"http://npaa7587.petrobras.biz/WebFacil3/resultBrTecEng.aspx?strCodDoc=BDV  -52416011-SP- \">BDV  -52416011-SP- </a>",
                "<a href=\"http://www.google.com\">Google</a>", "<a href=\"http://www.yahoo.com\">Yahoo</a>", "", "<a href=\"http://www.microsoft.com\">Microsoft</a>",
                "<a href=\"http://www.apple.com\">Apple</a>", "<a href=\"http://www.github.com\">Github</a>", "<a href=\"http://www.google.com\">Google</a>",
                "<a href=\"http://www.yahoo.com\">Yahoo</a>", "<a href=\"http://www.microsoft.com\">Microsoft</a>", "<a href=\"http://www.apple.com\">Apple</a>",
                "<a href=\"http://www.github.com\">Github</a>", "<a href=\"http://www.google.com\">Google</a>", "<a href=\"http://www.yahoo.com\">Yahoo</a>",
                "<a href=\"http://www.microsoft.com\">Microsoft</a>", "<a href=\"http://www.apple.com\">Apple</a>", "<a href=\"http://www.github.com\">Github</a>",
                "<a href=\"http://www.google.com\">Google</a>", "<a href=\"http://www.yahoo.com\">Yahoo</a>", "<a href=\"http://www.microsoft.com\">Microsoft</a>"
        };
        var sheet2Column4Data = new string[] { "<a href=\"http://npaa7587.petrobras.biz/WebFacil3/resultBrTecEng.aspx?strCodDoc=ADV  -54236009-SP- \">ADV  -54236009-SP- </a>",
                "<a href=\"http://www.google.com\">Google</a>", "<a href=\"http://www.yahoo.com\">Yahoo</a>", "", "<a href=\"http://www.microsoft.com\">Microsoft</a>",
                "<a href=\"http://www.apple.com\">Apple</a>", "<a href=\"http://www.github.com\">Github</a>", "<a href=\"http://www.google.com\">Google</a>",
                "<a href=\"http://www.yahoo.com\">Yahoo</a>", "<a href=\"http://www.microsoft.com\">Microsoft</a>", "<a href=\"http://www.apple.com\">Apple</a>",
                "<a href=\"http://www.github.com\">Github</a>", "<a href=\"http://www.google.com\">Google</a>", "<a href=\"http://www.yahoo.com\">Yahoo</a>",
                "<a href=\"http://www.microsoft.com\">Microsoft</a>", "<a href=\"http://www.apple.com\">Apple</a>", "<a href=\"http://www.github.com\">Github</a>",
                "<a href=\"http://www.google.com\">Google</a>", "<a href=\"http://www.yahoo.com\">Yahoo</a>", "<a href=\"http://www.microsoft.com\">Microsoft</a>"
        };


        // Sheet 3
        var sheet3HeaderData = new string[] { "Numbers", "Auto Sum Column", "Dates", "Dates Auto" };

        var sheet3Column1Data = new string[] { "100", "2000", "30", "40000", "500.54", "6", "7", "8.563", "9", "10", "20000.34", "0.237" };

        var sheet3Column2Data = new string[] { "10.10", "11.11", "9.76", "8.566", "30.45", "5.87", "6.34", "8.563", "9.23", "", "10.10", "11.11" };

        var sheet3Column3Data = new string[] { "12/10/1977", "13/10/1977", "14/10/1977", "15/10/1977", "16/10/1977", "", "13/10/1977", "14/10/1977", "15/10/1977", "16/10/1977",
                "12/11/1977", "13/11/1977"
        };
        var sheet3Column4Data = new string[] { "12/10/1977", "16/10/1977", "18/10/1977", "10/10/1977", "10/10/1977", "25/01/1965", "11/10/1977", "11/10/1977", "11/10/1977",
                "11/10/1977", "11/11/1977", "11/11/1977"
        };


        var filename = GetNextFilename();

        var excel = ExcelWorkbookBuilder.StartWorkbook(Path.GetFileName(filename).Replace(Path.GetExtension(filename), ""))
            .WithAuthor("Excel Lib")
            .WithTitle("Excel Lib Test")
            .WithCompany("Excel Lib")
            .WithComments("Very nice comment goes here")
            .WithGlobalDateFormat("dd/MM/yyyy")
            .WithGlobalHtmlTagHyperlinks()
            .AddTableSheet("Summary")
                .WithTheme(Configurations.ExcelModelDefs.ExcelThemes.TableStyleLight1)
                .WithTabColor("FF222222")
                .AddHeader(sheet1HeaderData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.CalibriLight, 14, true, false, false)
                    .TableSheet
                .AddColumn(sheet1Column1Data)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(sheet1Column2Data)
                    .WithDataType(Configurations.ExcelModelDefs.ExcelDataTypes.Sheetlink)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.TimesNewRoman, 11, false, false, false)
                    .TableSheet
                .Workbook
            .AddTableSheet("Custom Spreadsheet Name")
                .WithTheme(Configurations.ExcelModelDefs.ExcelThemes.TableStyleLight17)
                .WithTabColor("FFEB8638")
                .AddHeader(sheet2HeaderData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.CalibriLight, 14, true, false, false)
                    .WithRowHeight(30)
                    .TableSheet
                .AddColumn(sheet2Column1Data)
                    .WithDataType(Configurations.ExcelModelDefs.ExcelDataTypes.Hyperlink)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, true, true)
                    .WithMaxWidth(20)
                    .WithNewLineString("<br>")
                    .TableSheet
                .AddColumn(sheet2Column2Data)
                    .WithDataType(Configurations.ExcelModelDefs.ExcelDataTypes.AutoDetect)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.TimesNewRoman, 11, false, false, false)
                    .TableSheet
                .AddColumn(sheet2Column3Data)
                    .WithDataType(Configurations.ExcelModelDefs.ExcelDataTypes.AutoDetect)
                    .WithNewLineString("<br>")
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Arial, 11, false, false, false)
                    .TableSheet
                .Workbook
            .AddTableSheet("Spreadsheet 2")
                .WithTheme(Configurations.ExcelModelDefs.ExcelThemes.TableStyleLight7)
                .WithTabColor("FF5A8F28")
                .AddHeader(sheet3HeaderData)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Arial, 14, true, false, false, "FF000000")
                    .WithRowHeight(25)
                    .TableSheet
                .AddColumn(sheet3Column1Data)
                    .WithDataType(Configurations.ExcelModelDefs.ExcelDataTypes.Number)
                    .WithDataFormat("#,##0.00")
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, true, true)
                    .WithMaxWidth(20)
                    .TableSheet
                .AddColumn(sheet3Column2Data)
                    .WithDataType(Configurations.ExcelModelDefs.ExcelDataTypes.Number)
                    .WithDataFormat("R$ #,##0.00")
                    .AddSubtotal()
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.TimesNewRoman, 11, false, false, false)
                    .TableSheet
                .AddColumn(sheet3Column3Data)
                    .WithDataType(Configurations.ExcelModelDefs.ExcelDataTypes.DateTime)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, false, false, false)
                    .TableSheet
                .AddColumn(sheet3Column4Data)
                    .WithDataType(Configurations.ExcelModelDefs.ExcelDataTypes.AutoDetect)
                    .WithFont(Configurations.ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, false, false, false)
                    .TableSheet
                .Workbook
            .Build();


        //Escrever em disco
        using (var fileStream = File.Create(filename))
        {
            excel.Seek(0, SeekOrigin.Begin);
            excel.CopyTo(fileStream);
        }
    }

    private string GetNextFilename()
    {
        string outputPath = Directory.GetCurrentDirectory();
        outputPath = Path.Combine(outputPath, "..", "..", "..", "output");
                
        var nameCounter = 1;
        var baseFilename = "output";
        var filename = baseFilename;

        while (File.Exists(Path.Combine(outputPath, filename + ".xlsx")))
        {
            filename = baseFilename + nameCounter++;
        }
        
        filename = Path.Combine(outputPath, filename + ".xlsx");
        
        return filename;
    }
}
