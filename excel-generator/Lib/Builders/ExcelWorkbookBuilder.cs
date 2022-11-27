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

    #endregion

    #region Basic Workbook configuration

    public ExcelWorkbookBuilder SetGlobalDateFormat(string format)
    {
        _workbookModel.GlobalColumnBehavior.Date.Format = format;
        return this;
    }

    public ExcelWorkbookBuilder SetGlobalNewLineString(string newLineString)
    {
        _workbookModel.GlobalColumnBehavior.Text.NewLineString = newLineString;
        return this;
    }

    public ExcelWorkbookBuilder SetHyperlinkStyle(bool isHtml)
    {
        _workbookModel.GlobalColumnBehavior.Hyperlink.IsHtml = isHtml;
        return this;
    }

    public ExcelWorkbookBuilder SetAuthor(string author)
    {
        _workbookModel.AuthoringMetadata.Author = author;
        return this;
    }
    public ExcelWorkbookBuilder SetTitle(string title)
    {
        _workbookModel.AuthoringMetadata.Title = title;
        return this;
    }
    public ExcelWorkbookBuilder SetCompany(string company)
    {
        _workbookModel.AuthoringMetadata.Company = company;
        return this;
    }
    public ExcelWorkbookBuilder SetComments(string comments)
    {
        _workbookModel.AuthoringMetadata.Comments = comments;
        return this;
    }

    #endregion
}
