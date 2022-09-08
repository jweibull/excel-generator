using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelGenerator.ForBenchmarking;

[MemoryDiagnoser]
public class DOMBasedGenerator
{
    [Benchmark]
    public void DOM()
    {
        //string path = Directory.GetCurrentDirectory();
        //path = Path.Combine(path, "..", "..", "output");
        //for (int i = 0; i < 1; i++)
        //{
        //    var nameCounter = 1;
        //    var baseFilename = "output";
        //    var filename = baseFilename;
        //    while (File.Exists(Path.Combine(path, filename + ".xlsx")))
        //    {
        //        filename = baseFilename + nameCounter++;
        //    }
        //    filename = Path.Combine(path, filename + ".xlsx");

        //    CreatePackage();
        //}

        CreatePackage();
    }


    public void CreatePackage()
    {
        using (var stream = new MemoryStream())
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                //this list of attributes will be used when writing a start element
                WorkbookPart workbookpart = document.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();
                if (workbookpart != null)
                {
                    WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    SheetData sheetData = new SheetData();
                    worksheetPart.Worksheet = new Worksheet(sheetData);

                    Sheets sheets = workbookpart.Workbook.AppendChild<Sheets>(new Sheets());

                    Sheet sheet = new Sheet()
                    {
                        Id = workbookpart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "mySheet"
                    };
                    sheets.Append(sheet);

                    Row row;
                    Cell cell;
                    CellValue cellValue;
                    for (int rowNum = 1; rowNum <= 100000; rowNum++)
                    {
                        row = new Row() { RowIndex = (UInt32)rowNum };
                        sheetData.Append(row);
                        for (int columnNum = 1; columnNum <= 50; columnNum++)
                        {
                            cell = new Cell() { CellReference = string.Format("{0}{1}", GetColumnName(columnNum), rowNum), DataType = CellValues.String };
                            cellValue = new CellValue(string.Format("This is Row {0}, Column {1}", rowNum, columnNum));
                            cell.Append(cellValue);
                            sheetData.Append(cell);
                        }
                    }

                    document.Save();

                    document.Close();
                }
            }
        }
    }



    //A simple helper to get the column name from the column index. This is not well tested!
    private string GetColumnName(int columnIndex)
    {
        int dividend = columnIndex;
        string columnName = string.Empty;
        int modifier;

        while (dividend > 0)
        {
            modifier = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modifier).ToString() + columnName;
            dividend = (dividend - modifier) / 26;
        }

        return columnName;
    }
}
