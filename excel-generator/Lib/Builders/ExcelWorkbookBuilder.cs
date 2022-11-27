using rbkApiModules.Utilities.Excel.InputModel;

namespace rbkApiModules.Utilities.Excel;

public class ExcelWorkbookBuilder
{
    private readonly ExcelWorkbookModel _workbookModel = new ExcelWorkbookModel();
    
    private int _tabCount = 0;

    private ExcelWorkbookBuilder(string filename)
    {
        _workbookModel.FileName = filename;
    }

    public static ExcelWorkbookBuilder StartWorkbook(string filename)
    {
        return new ExcelWorkbookBuilder(filename);
    }

    #region chain builders

    public ExcelTableSheetBuilder AddTableSheet(string sheetName)
    {
        _tabCount++;
        var sheet = new ExcelTableSheetModel()
        {
            Name = sheetName,
            TabIndex = _tabCount
        };

        _workbookModel.Tables.Add(sheet);

        return ExcelTableSheetBuilder.AddSheet(this, sheet);
    }

    public Stream Build()
    {
        var lib = new SaxLib();
        var stream = lib.CreatePackage(_workbookModel);

        return stream;
    }

    #endregion

    #region Basic Workbook configuration

    public ExcelWorkbookBuilder WithGlobalDateFormat(string format)
    {
        _workbookModel.GlobalColumnBehavior.Date.Format = format;
        return this;
    }

    public ExcelWorkbookBuilder WithGlobalNewLineString(string newLineString)
    {
        _workbookModel.GlobalColumnBehavior.Text.NewLineString = newLineString;
        return this;
    }

    public ExcelWorkbookBuilder WithGlobalHtmlTagHyperlinks()
    {
        _workbookModel.GlobalColumnBehavior.Hyperlink.IsHtml = true;
        return this;
    }

    public ExcelWorkbookBuilder WithAuthor(string author)
    {
        _workbookModel.AuthoringMetadata.Author = author;
        return this;
    }
    public ExcelWorkbookBuilder WithTitle(string title)
    {
        _workbookModel.AuthoringMetadata.Title = title;
        return this;
    }
    public ExcelWorkbookBuilder WithCompany(string company)
    {
        _workbookModel.AuthoringMetadata.Company = company;
        return this;
    }
    public ExcelWorkbookBuilder WithComments(string comments)
    {
        _workbookModel.AuthoringMetadata.Comments = comments;
        return this;
    }

    #endregion
}
