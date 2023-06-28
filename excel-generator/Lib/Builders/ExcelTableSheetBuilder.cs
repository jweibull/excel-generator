using TableExporter.Configurations;
using TableExporter.InputModel;

namespace TableExporter;

public class ExcelTableSheetBuilder
{
    public ExcelWorkbookBuilder Workbook => _builder as ExcelWorkbookBuilder;

    private readonly ExcelTableSheetModel _sheet;
    private readonly ExcelWorkbookBuilder _builder;

    private ExcelTableSheetBuilder(ExcelWorkbookBuilder builder, ExcelTableSheetModel sheet)
    {
        _builder = builder;
        _sheet = sheet;
        _sheet.SheetType = ExcelModelDefs.ExcelSheetTypes.Table;
    }

    public static ExcelTableSheetBuilder AddSheet(ExcelWorkbookBuilder builder, ExcelTableSheetModel sheet)
    {
        return new ExcelTableSheetBuilder(builder, sheet);
    }

    public ExcelTableSheetBuilder WithTheme(ExcelModelDefs.ExcelThemes theme)
    {
        _sheet.Theme = theme;
        return this;
    }

    public ExcelTableSheetBuilder WithTabColor(string color)
    {
        _sheet.TabColor = color;
        return this;
    }

    public ExcelHeaderBuilder AddHeader(string[] data)
    {
        var header = new ExcelHeaderModel()
        {
            Data = data
        };

        _sheet.Header = header;

        return ExcelHeaderBuilder.AddHeader(this, header);
    }

    public ExcelColumnBuilder AddColumn(string[] data)
    {
        var column = new ExcelColumnModel()
        {
            Data = data
        };

        _sheet.Columns.Add(column);

        return ExcelColumnBuilder.AddColumn(this, column);
    }
}
