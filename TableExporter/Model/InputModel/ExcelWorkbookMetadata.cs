namespace TableExporter;

/// <summary>
/// The Main Workbook Model container. Holds all data and metadata.
/// </summary>
public class ExcelWorkbookMetadata
{
    /// <summary>
    /// Authoring Metadata, Title
    /// </summary>
    public string Title { get; set; } = String.Empty;
    /// <summary>
    /// Authoring Metadata, Author name
    /// </summary>
    public string Author { get; set; } = String.Empty;

    /// <summary>
    /// Authoring Metadata, Company name
    /// </summary>
    public string Company { get; set; } = String.Empty;

    /// <summary>
    /// Authoring Metadata, Comments
    /// </summary>
    public string Comments { get; set; } = String.Empty;
}
