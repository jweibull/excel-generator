namespace TableExporter;

// Author/Title/Company/Comments are enforced "set at most once" via runtime checks.
// A type-level version would require 16 interfaces (one per subset of "already set" metadata)
// so each setter could return a type that no longer exposes that setter; possible but heavy.

/// <summary>
/// Workbook-level operations available after optional metadata has been set (AddTableSheet, globals, Build).
/// </summary>
public interface IExcelWorkbookBuilderCore
{
    ExcelTableSheetBuilder AddTableSheet(string sheetName);
    Stream Build();
    IExcelWorkbookBuilderCore WithGlobalDateFormat(string format);
    IExcelWorkbookBuilderCore WithGlobalNewLineSeparator(string newLineSeparator);
    IExcelWorkbookBuilderCore WithGlobalHtmlTagHyperlinks();
}

/// <summary>
/// Workbook builder with optional authoring metadata (each setter can be called at most once at runtime).
/// </summary>
public interface IExcelWorkbookBuilder : IExcelWorkbookBuilderCore
{
    IExcelWorkbookBuilder WithAuthor(string author);
    IExcelWorkbookBuilder WithTitle(string title);
    IExcelWorkbookBuilder WithCompany(string company);
    IExcelWorkbookBuilder WithComments(string comments);
}

public class ExcelWorkbookBuilder : IExcelWorkbookBuilder
{
    private readonly ExcelWorkbookModel _workbookModel = new ExcelWorkbookModel();
    private int _tabCount;
    private bool _authorSet;
    private bool _titleSet;
    private bool _companySet;
    private bool _commentsSet;

    private ExcelWorkbookBuilder(string filename)
    {
        _workbookModel.FileName = filename;
    }

    public static IExcelWorkbookBuilder StartWorkbook(string filename)
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
        var lib = new TableExporterService();
        return lib.GenerateSpreadsheetAsBase64(_workbookModel);
    }

    #endregion

    #region Basic Workbook configuration

    public ExcelWorkbookBuilder WithGlobalDateFormat(string format)
    {
        _workbookModel.GlobalColumnBehavior.Date.Format = format;
        return this;
    }

    IExcelWorkbookBuilderCore IExcelWorkbookBuilderCore.WithGlobalDateFormat(string format)
    {
        WithGlobalDateFormat(format);
        return this;
    }

    public ExcelWorkbookBuilder WithGlobalNewLineSeparator(string newLineSeparator)
    {
        _workbookModel.GlobalColumnBehavior.NewLineSeparator = newLineSeparator;
        return this;
    }

    IExcelWorkbookBuilderCore IExcelWorkbookBuilderCore.WithGlobalNewLineSeparator(string newLineSeparator)
    {
        WithGlobalNewLineSeparator(newLineSeparator);
        return this;
    }

    public ExcelWorkbookBuilder WithGlobalHtmlTagHyperlinks()
    {
        _workbookModel.GlobalColumnBehavior.Hyperlink.IsHtml = true;
        return this;
    }

    IExcelWorkbookBuilderCore IExcelWorkbookBuilderCore.WithGlobalHtmlTagHyperlinks()
    {
        WithGlobalHtmlTagHyperlinks();
        return this;
    }

    public ExcelWorkbookBuilder WithAuthor(string author)
    {
        if (_authorSet)
            throw new InvalidOperationException("Author has already been set.");
        _authorSet = true;
        _workbookModel.AuthoringMetadata.Author = author;
        return this;
    }

    IExcelWorkbookBuilder IExcelWorkbookBuilder.WithAuthor(string author)
    {
        return WithAuthor(author);
    }

    public ExcelWorkbookBuilder WithTitle(string title)
    {
        if (_titleSet)
            throw new InvalidOperationException("Title has already been set.");
        _titleSet = true;
        _workbookModel.AuthoringMetadata.Title = title;
        return this;
    }

    IExcelWorkbookBuilder IExcelWorkbookBuilder.WithTitle(string title)
    {
        return WithTitle(title);
    }

    public ExcelWorkbookBuilder WithCompany(string company)
    {
        if (_companySet)
            throw new InvalidOperationException("Company has already been set.");
        _companySet = true;
        _workbookModel.AuthoringMetadata.Company = company;
        return this;
    }

    IExcelWorkbookBuilder IExcelWorkbookBuilder.WithCompany(string company)
    {
        return WithCompany(company);
    }

    public ExcelWorkbookBuilder WithComments(string comments)
    {
        if (_commentsSet)
            throw new InvalidOperationException("Comments have already been set.");
        _commentsSet = true;
        _workbookModel.AuthoringMetadata.Comments = comments;
        return this;
    }

    IExcelWorkbookBuilder IExcelWorkbookBuilder.WithComments(string comments)
    {
        return WithComments(comments);
    }

    #endregion
}
