using Newtonsoft.Json;
using TableExporter;

namespace TableExporterApp;

public class Runner
{
    public void Run()
    {

        var tableExporter = new TableExporterService();

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

        var stream = tableExporter.GenerateSpreadsheetAsBase64(modelData.WorkbookModel);

        using (var fileStream = File.Create(filename))
        {
            stream.Seek(0, SeekOrigin.Begin);
            stream.CopyTo(fileStream);
        }
        
    }

    public void RunFluent()
    {
        var filename = GetNextFilename();

        var excel = ExcelWorkbookBuilder.StartWorkbook(Path.GetFileName(filename).Replace(Path.GetExtension(filename), ""))
            .WithAuthor("Excel Lib")
            .WithTitle("Excel Lib Test")
            .WithCompany("Excel Lib")
            .WithComments("Very nice comment goes here")
            .WithGlobalDateFormat("dd/MM/yyyy")
            .WithGlobalHtmlTagHyperlinks()
            .WithGlobalNewLineSeparator("")
            .AddTableSheet("Summary")
                .WithTheme(ExcelModelDefs.ExcelThemes.TableStyleLight1)
                .WithTabColor("FF222222")
                .AddHeader(DataWithLinks.Sheet1HeaderData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.CalibriLight, 14, true, false, false)
                    .TableSheet
                .AddColumn(DataWithLinks.Sheet1Column1Data)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(DataWithLinks.Sheet1Column2Data)
                    .WithDataType(ExcelModelDefs.ExcelDataTypes.Sheetlink)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.TimesNewRoman, 11, false, false, false)
                    .TableSheet
                .Workbook
            .AddTableSheet("Custom Spreadsheet Name")
                .WithTheme(ExcelModelDefs.ExcelThemes.TableStyleLight17)
                .WithTabColor("FFEB8638")
                .AddHeader(DataWithLinks.Sheet2HeaderData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.CalibriLight, 14, true, false, false)
                    .WithRowHeight(30)
                    .TableSheet
                .AddColumn(DataWithLinks.Sheet2Column1Data)
                    .WithDataType(ExcelModelDefs.ExcelDataTypes.Hyperlink)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, true, true)
                    .WithMaxWidth(20)
                    .WithNewLineSeparator("<br>")
                    .TableSheet
                .AddColumn(DataWithLinks.Sheet2Column2Data)
                    .WithDataType(ExcelModelDefs.ExcelDataTypes.AutoDetect)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.TimesNewRoman, 11, false, false, false)
                    .TableSheet
                .AddColumn(DataWithLinks.Sheet2Column3Data)
                    .WithDataType(ExcelModelDefs.ExcelDataTypes.AutoDetect)
                    .WithNewLineSeparator("<br>")
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Arial, 11, false, false, false)
                    .TableSheet
                .Workbook
            .AddTableSheet("Spreadsheet 2")
                .WithTheme(ExcelModelDefs.ExcelThemes.TableStyleLight7)
                .WithTabColor("FF5A8F28")
                .AddHeader(DataWithLinks.Sheet3HeaderData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Arial, 14, true, false, false, "FF000000")
                    .WithRowHeight(25)
                    .TableSheet
                .AddColumn(DataWithLinks.Sheet3Column1Data)
                    .WithDataType(ExcelModelDefs.ExcelDataTypes.Number)
                    .WithDataFormat("#,##0.00")
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, true, true)
                    .WithMaxWidth(20)
                    .TableSheet
                .AddColumn(DataWithLinks.Sheet3Column2Data)
                    .WithDataType(ExcelModelDefs.ExcelDataTypes.Number)
                    .WithDataFormat("R$ #,##0.00")
                    .AddSubtotal()
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.TimesNewRoman, 11, false, false, false)
                    .TableSheet
                .AddColumn(DataWithLinks.Sheet3Column3Data)
                    .WithDataType(ExcelModelDefs.ExcelDataTypes.DateTime)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, false, false, false)
                    .TableSheet
                .AddColumn(DataWithLinks.Sheet3Column4Data)
                    .WithDataType(ExcelModelDefs.ExcelDataTypes.AutoDetect)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, false, false, false)
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

    public void RunEnviron()
    {
        var filename = GetNextFilename();
        // Sheet 1
        var sheet1HeaderData = EnvironData.Headers;

        Console.WriteLine(EnvironData.Headers.Count());

        var excel = ExcelWorkbookBuilder.StartWorkbook(Path.GetFileName(filename).Replace(Path.GetExtension(filename), ""))
            .DisableAuthoringMetadata()
            .WithTitle("Excel Lib Test")
            .WithGlobalDateFormat("dd/MM/yyyy")
            .WithGlobalHtmlTagHyperlinks()
            .WithGlobalNewLineSeparator("")
            .AddTableSheet("Cronos")
                .AddHeader(sheet1HeaderData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.TimesNewRoman, 14, true, false, false)
                    .TableSheet
                .AddColumn(EnvironData.Grupo)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.TimesNewRoman, 11, true, false, false)
                    .TableSheet
                .AddColumn(EnvironData.Subgrupo)
                    .WithDataType(ExcelModelDefs.ExcelDataTypes.Sheetlink)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.TimesNewRoman, 11, false, false, false)
                    .TableSheet
                .AddColumn(EnvironData.Percentual_atual)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.TimesNewRoman, 11, false, false, false)
                    .TableSheet
                .AddColumn(EnvironData.Percentual_corrosao_simulado)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.TimesNewRoman, 11, false, false, false)
                    .TableSheet
                .AddColumn(EnvironData.Percentual_corrosao_pos_pintura)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.TimesNewRoman, 11, false, false, false)
                    .TableSheet
                .AddColumn(EnvironData.Plano)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.TimesNewRoman, 11, false, false, false)
                    .TableSheet
                .AddColumn(EnvironData.Area_a_pintar)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.TimesNewRoman, 11, false, false, false)
                    .TableSheet
                .AddColumn(EnvironData.Nomecor)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.TimesNewRoman, 11, false, false, false)
                    .TableSheet
                .AddColumn(EnvironData.Cor)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.TimesNewRoman, 11, false, false, false)
                    .TableSheet
                .AddColumn(EnvironData.RTIs_quitadas)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.TimesNewRoman, 11, false, false, false)
                    .TableSheet
                .AddColumn(EnvironData.RTIs_quitadas_A)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.TimesNewRoman, 11, false, false, false)
                    .TableSheet
                .AddColumn(EnvironData.RTIs_quitadas_B)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.TimesNewRoman, 11, false, false, false)
                    .TableSheet
                .AddColumn(EnvironData.RTIs_quitadas_C)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.TimesNewRoman, 11, false, false, false)
                    .TableSheet
                .AddColumn(EnvironData.RTIs_quitadas_D)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.TimesNewRoman, 11, false, false, false)
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

    public void RunMockData()
    {
        var sheet1HeaderData = new string[80];
        for (int i = 0; i < sheet1HeaderData.Length; i++)
        {
            sheet1HeaderData[i] = $"Lorem Ipsum {i}";
        }

        var columnData = new string[200000];
        for (int i = 0; i < columnData.Length; i++)
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
                .WithTheme(ExcelModelDefs.ExcelThemes.TableStyleLight1)
                .WithTabColor("FF222222")
                .AddHeader(sheet1HeaderData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.CalibriLight, 14, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .AddColumn(columnData)
                    .WithFont(ExcelModelDefs.ExcelFonts.FontType.Calibri, 11, true, false, false)
                    .TableSheet
                .Workbook
            .Build();

        // Escrever em disco
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
        Directory.CreateDirectory(outputPath);
        
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
