﻿using rbkApiModules.Utilities.Excel.Configurations;
using rbkApiModules.Utilities.Excel.InputModel;

namespace rbkApiModules.Utilities.Excel;

public class ExcelHeaderBuilder
{
    public ExcelTableSheetBuilder TableSheet => _builder;

    private readonly ExcelHeaderModel _header;
    private readonly ExcelTableSheetBuilder _builder;

    private ExcelHeaderBuilder(ExcelTableSheetBuilder builder, ExcelHeaderModel header)
    {
        _builder = builder;
        _header = header;
    }

    public static ExcelHeaderBuilder AddHeader(ExcelTableSheetBuilder builder, ExcelHeaderModel header)
    {
        return new ExcelHeaderBuilder(builder, header);
    }

    public ExcelHeaderBuilder SetRowHeight(int height) 
    {
        _header.RowHeight = height;
        return this;
    }

    public ExcelHeaderBuilder SetFont(ExcelModelDefs.ExcelFonts.FontType font, int fontSize, bool bold, bool italic, bool underlined, string? fontcolor)
    {
        _header.Style.Font = font;
        _header.Style.FontSize = fontSize;
        _header.Style.Bold = bold;
        _header.Style.Italic = italic;
        _header.Style.Underline = underlined;
        
        if (!String.IsNullOrEmpty(fontcolor)) 
        {
            _header.Style.FontColor = fontcolor;
        }

        return this;
    }
}
