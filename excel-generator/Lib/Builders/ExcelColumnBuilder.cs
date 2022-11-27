using rbkApiModules.Utilities.Excel.Configurations;
using rbkApiModules.Utilities.Excel.InputModel;

namespace rbkApiModules.Utilities.Excel;

public class ExcelColumnBuilder
{
    public ExcelTableSheetBuilder TableSheet => _builder;

    private readonly ExcelColumnModel _column;
    private readonly ExcelTableSheetBuilder _builder;

    private ExcelColumnBuilder(ExcelTableSheetBuilder builder, ExcelColumnModel column)
    {
        _builder = builder;
        _column = column;
    }

    public static ExcelColumnBuilder AddColumn(ExcelTableSheetBuilder builder, ExcelColumnModel column)
    {
        return new ExcelColumnBuilder(builder, column);
    }

    public ExcelColumnBuilder WithDataType(ExcelModelDefs.ExcelDataTypes type)
    {
        _column.DataType = type;
        return this;
    }

    public ExcelColumnBuilder WithDataFormat(string format)
    {
        _column.DataFormat = format;
        return this;
    }

    public ExcelColumnBuilder WithNewLineString(string newLineString)
    {
        _column.NewLineString = newLineString;
        return this;
    }

    public ExcelColumnBuilder WithSubtotal()
    {
        if (!(_column.DataType == ExcelModelDefs.ExcelDataTypes.Number))
        {
            throw new Exception("Can't apply subtotal on a non numeric column");
        }

        _column.HasSubtotal = true;

        return this;
    }

    public ExcelColumnBuilder WithFont(ExcelModelDefs.ExcelFonts.FontType font, int fontSize, bool bold, bool italic, bool underlined, string? fontcolor)
    {
        _column.Style.Font = font;
        _column.Style.FontSize = fontSize;
        _column.Style.Bold = bold;
        _column.Style.Italic = italic;
        _column.Style.Underline = underlined;
        
        if (!String.IsNullOrEmpty(fontcolor)) 
        {
            _column.Style.FontColor = fontcolor;
        }

        return this;
    }

    public ExcelColumnBuilder WithMaxWidth(int width)
    {
        _column.MaxWidth = width;
        return this;
    }
}
