namespace TableExporter;

public interface ITableExporterService
{
    public MemoryStream GenerateSpreadsheetAsBase64(ExcelWorkbookModel workbookModel);
}
