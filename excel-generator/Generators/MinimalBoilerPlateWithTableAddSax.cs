using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelGenerator.Generators;

public class MinimalBoilerPlateWithTableAddSax
{
    // Creates a SpreadsheetDocument.
    public void CreatePackage(string filePath)
    {
        using (SpreadsheetDocument package = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
        {
            CreateParts(package);
        }
    }

    // Adds child parts and generates content of the specified part.
    private void CreateParts(SpreadsheetDocument document)
    {
        WorkbookPart workbookPart1 = document.AddWorkbookPart();
        GenerateWorkbookPart1Content(workbookPart1);

        WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId2");
        GenerateWorkbookStylesPart1Content(workbookStylesPart1);

        WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
        GenerateWorksheetPart1Content(worksheetPart1);

        TableDefinitionPart tableDefinitionPart1 = worksheetPart1.AddNewPart<TableDefinitionPart>("rId1");
        GenerateTableDefinitionPart1Content(tableDefinitionPart1);

        SharedStringTablePart sharedStringTablePart1 = workbookPart1.AddNewPart<SharedStringTablePart>("rId3");
        GenerateSharedStringTablePart1Content(sharedStringTablePart1);

        SetPackageProperties(document);
    }

    
    // Generates content of workbookPart1.
    private void GenerateWorkbookPart1Content(WorkbookPart workbookPart1)
    {
        Workbook workbook1 = new Workbook();

        Sheets sheets1 = new Sheets();
        Sheet sheet1 = new Sheet() { Name = "Planilha1", SheetId = (UInt32Value)1U, Id = "rId1" };

        sheets1.Append(sheet1);
        
        workbook1.Append(sheets1);
        
        workbookPart1.Workbook = workbook1;
    }

    // Generates content of workbookStylesPart1.
    private void GenerateWorkbookStylesPart1Content(WorkbookStylesPart workbookStylesPart1)
    {
        Stylesheet stylesheet1 = new Stylesheet(); 

        Fonts fonts1 = new Fonts() { Count = (UInt32Value)2U, KnownFonts = true };

        Font font1 = new Font();
        FontSize fontSize1 = new FontSize() { Val = 11D };
        Color color1 = new Color() { Theme = (UInt32Value)1U };
        FontName fontName1 = new FontName() { Val = "Calibri" };
        FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
        FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

        font1.Append(fontSize1);
        font1.Append(color1);
        font1.Append(fontName1);
        font1.Append(fontFamilyNumbering1);
        font1.Append(fontScheme1);

        Font font2 = new Font();
        FontSize fontSize2 = new FontSize() { Val = 11D };
        Color color2 = new Color() { Theme = (UInt32Value)1U };
        FontName fontName2 = new FontName() { Val = "Calibri Light" };
        FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
        FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Major };

        font2.Append(fontSize2);
        font2.Append(color2);
        font2.Append(fontName2);
        font2.Append(fontFamilyNumbering2);
        font2.Append(fontScheme2);

        fonts1.Append(font1);
        fonts1.Append(font2);

        Fills fills1 = new Fills() { Count = (UInt32Value)2U };

        Fill fill1 = new Fill();
        PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

        fill1.Append(patternFill1);

        Fill fill2 = new Fill();
        PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

        fill2.Append(patternFill2);

        fills1.Append(fill1);
        fills1.Append(fill2);

        Borders borders1 = new Borders() { Count = (UInt32Value)1U };

        Border border1 = new Border();
        LeftBorder leftBorder1 = new LeftBorder();
        RightBorder rightBorder1 = new RightBorder();
        TopBorder topBorder1 = new TopBorder();
        BottomBorder bottomBorder1 = new BottomBorder();
        DiagonalBorder diagonalBorder1 = new DiagonalBorder();

        border1.Append(leftBorder1);
        border1.Append(rightBorder1);
        border1.Append(topBorder1);
        border1.Append(bottomBorder1);
        border1.Append(diagonalBorder1);

        borders1.Append(border1);

        CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
        CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

        cellStyleFormats1.Append(cellFormat1);

        CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)2U };
        CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
        CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true };

        cellFormats1.Append(cellFormat2);
        cellFormats1.Append(cellFormat3);

        CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
        CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

        cellStyles1.Append(cellStyle1);

        DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)1U };

        DifferentialFormat differentialFormat1 = new DifferentialFormat();

        Font font3 = new Font();
        Bold bold1 = new Bold() { Val = false };
        Italic italic1 = new Italic() { Val = false };
        Strike strike1 = new Strike() { Val = false };
        Condense condense1 = new Condense() { Val = false };
        Extend extend1 = new Extend() { Val = false };
        Outline outline1 = new Outline() { Val = false };
        Shadow shadow1 = new Shadow() { Val = false };
        Underline underline1 = new Underline() { Val = UnderlineValues.None };
        VerticalTextAlignment verticalTextAlignment1 = new VerticalTextAlignment() { Val = VerticalAlignmentRunValues.Baseline };
        FontSize fontSize3 = new FontSize() { Val = 11D };
        Color color3 = new Color() { Theme = (UInt32Value)1U };
        FontName fontName3 = new FontName() { Val = "Calibri Light" };
        FontScheme fontScheme3 = new FontScheme() { Val = FontSchemeValues.Major };

        font3.Append(bold1);
        font3.Append(italic1);
        font3.Append(strike1);
        font3.Append(condense1);
        font3.Append(extend1);
        font3.Append(outline1);
        font3.Append(shadow1);
        font3.Append(underline1);
        font3.Append(verticalTextAlignment1);
        font3.Append(fontSize3);
        font3.Append(color3);
        font3.Append(fontName3);
        font3.Append(fontScheme3);

        differentialFormat1.Append(font3);

        differentialFormats1.Append(differentialFormat1);
        TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

        stylesheet1.Append(fonts1);
        stylesheet1.Append(fills1);
        stylesheet1.Append(borders1);
        stylesheet1.Append(cellStyleFormats1);
        stylesheet1.Append(cellFormats1);
        stylesheet1.Append(cellStyles1);
        stylesheet1.Append(differentialFormats1);
        stylesheet1.Append(tableStyles1);
        
        workbookStylesPart1.Stylesheet = stylesheet1;
    }

    // Generates content of worksheetPart1.
    private void GenerateWorksheetPart1Content(WorksheetPart worksheetPart1)
    {
        Worksheet worksheet1 = new Worksheet(); 

        Columns columns1 = new Columns();
        Column column1 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 12D, CustomWidth = true };

        columns1.Append(column1);

        SheetData sheetData1 = new SheetData();

        Row row1 = new Row() { RowIndex = (UInt32Value)1U };

        Cell cell1 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
        CellValue cellValue1 = new CellValue();
        cellValue1.Text = "0";

        cell1.Append(cellValue1);

        row1.Append(cell1);

        Row row2 = new Row() { RowIndex = (UInt32Value)2U };

        Cell cell2 = new Cell() { CellReference = "A2", DataType = CellValues.SharedString };
        CellValue cellValue2 = new CellValue();
        cellValue2.Text = "1";

        cell2.Append(cellValue2);

        row2.Append(cell2);

        Row row3 = new Row() { RowIndex = (UInt32Value)3U };

        Cell cell3 = new Cell() { CellReference = "A3", DataType = CellValues.SharedString };
        CellValue cellValue3 = new CellValue();
        cellValue3.Text = "2";

        cell3.Append(cellValue3);

        row3.Append(cell3);

        sheetData1.Append(row1);
        sheetData1.Append(row2);
        sheetData1.Append(row3);
        
        TableParts tableParts1 = new TableParts() { Count = (UInt32Value)1U };
        TablePart tablePart1 = new TablePart() { Id = "rId1" };

        tableParts1.Append(tablePart1);

        worksheet1.Append(columns1);
        worksheet1.Append(sheetData1);
        worksheet1.Append(tableParts1);

        worksheetPart1.Worksheet = worksheet1;
    }

    // Generates content of tableDefinitionPart1.
    private void GenerateTableDefinitionPart1Content(TableDefinitionPart tableDefinitionPart1)
    {
        Table table1 = new Table() { Id = (UInt32Value)1U, Name = "Table1", DisplayName = "Table1", Reference = "A1:A3", TotalsRowShown = false, HeaderRowFormatId = (UInt32Value)0U };
        AutoFilter autoFilter1 = new AutoFilter() { Reference = "A1:A3" };

        TableColumns tableColumns1 = new TableColumns() { Count = (UInt32Value)1U };
        TableColumn tableColumn1 = new TableColumn() { Id = (UInt32Value)1U, Name = "Header" };

        tableColumns1.Append(tableColumn1);
        TableStyleInfo tableStyleInfo1 = new TableStyleInfo() { Name = "TableStyleLight3", ShowFirstColumn = false, ShowLastColumn = false, ShowRowStripes = true, ShowColumnStripes = false };

        table1.Append(autoFilter1);
        table1.Append(tableColumns1);
        table1.Append(tableStyleInfo1);

        tableDefinitionPart1.Table = table1;
    }

    // Generates content of sharedStringTablePart1.
    private void GenerateSharedStringTablePart1Content(SharedStringTablePart sharedStringTablePart1)
    {
        

        using (var writer = OpenXmlWriter.Create(sharedStringTablePart1))
        {
            // Change this based on real data count
            SharedStringTable sharedStringTable1 = new SharedStringTable() { Count = (UInt32Value)3U, UniqueCount = (UInt32Value)3U };
            writer.WriteStartElement(sharedStringTable1);


            //write the row start element with the row index attribute
            writer.WriteStartElement(new SharedStringItem());

            //write the text value
            writer.WriteElement(new Text("Header"));

            // write the end sharedItem element
            writer.WriteEndElement();

            //write the row start element with the row index attribute
            writer.WriteStartElement(new SharedStringItem());

            //write the text value
            writer.WriteElement(new Text("A"));

            // write the end sharedItem element
            writer.WriteEndElement();

            //write the row start element with the row index attribute
            writer.WriteStartElement(new SharedStringItem());

            //write the text value
            writer.WriteElement(new Text("B"));

            // write the end sharedItem element
            writer.WriteEndElement();


            // write the end SharedStringTable element
            writer.WriteEndElement();

            writer.Close();
        }
    }

    private void SetPackageProperties(OpenXmlPackage document)
    {
        document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2022-08-05T10:18:15Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        document.PackageProperties.LastModifiedBy = "Jon Karl Weibull";
    }

    //A simple helper to get the column name from the column index. This is not well tested!
    private string GetColumnName(int columnIndex)
    {
        int dividend = columnIndex;
        string columnName = String.Empty;
        int modifier;

        while (dividend > 0)
        {
            modifier = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modifier).ToString() + columnName;
            dividend = (int)((dividend - modifier) / 26);
        }

        return columnName;
    }

}
