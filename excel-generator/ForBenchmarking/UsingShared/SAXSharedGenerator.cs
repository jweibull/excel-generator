using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace ExcelGenerator.ForBenchmarking;

[MemoryDiagnoser]
public class SAXSharedGenerator
{
    private readonly int rowSize = 100000;
    private readonly int columnSize = 50;
    private readonly string[,] _allData;
    private readonly Dictionary<string, string> _sharedStringsToIndex;
    private int _sharedStringsCount;
    private int _sharedStringsUniqueCount;

    public SAXSharedGenerator()
    {
        _allData = new string[rowSize, columnSize];
        _sharedStringsToIndex = new Dictionary<string, string>();
        GenerateSharedData();
    }

    private void GenerateSharedData()
    {
        for (int rowNum = 1; rowNum <= 100000; rowNum++)
        {
            for (int columnNum = 1; columnNum <= 50; columnNum++)
            {
                var rowNumShared = rowNum + rowNum % 2;
                _allData[rowNum - 1, columnNum - 1] = string.Format("This is Row {0}, Column {1}", rowNumShared, columnNum);
            }
        }

        AddToSharedStringDictionary(_allData);
        _sharedStringsCount = _sharedStringsToIndex.Count();
    }

    private void AddToSharedStringDictionary(string[,] sharedStrings)
    {
        var count = 0;
        for (int rowIndex = 0; rowIndex < sharedStrings.GetLength(0); rowIndex++)
        {
            for (int columnIndex = 0; columnIndex < sharedStrings.GetLength(1); columnIndex++)
                if (_sharedStringsToIndex.ContainsKey(sharedStrings[rowIndex, columnIndex]))
            {
                count++;
            }
            else
            {
                count++;
                _sharedStringsToIndex.Add(sharedStrings[rowIndex, columnIndex], _sharedStringsToIndex.Count().ToString());
            }
        }
        _sharedStringsCount += count;
    }

    [Benchmark]
    public void SAX_SHARED()
    {
        CreatePackage();
    }


    public void CreatePackage()
    {
        using (var stream = new MemoryStream())
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                //this list of attributes will be used when writing a start element
                OpenXmlWriter writer;
                document.AddWorkbookPart();
                if (document.WorkbookPart != null)
                {
                    SharedStringTablePart sharedStringTablePart = document.WorkbookPart.AddNewPart<SharedStringTablePart>("ssId1");

                    GenerateSharedStringsTable(sharedStringTablePart);

                    WorksheetPart workSheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();

                    writer = OpenXmlWriter.Create(workSheetPart);
                    writer.WriteStartElement(new Worksheet());
                    writer.WriteStartElement(new SheetData());

                    var cell = new Cell();
                    var value = new CellValue();
                    for (int rowNum = 1; rowNum <= 100000; rowNum++)
                    {
                        //write the row start element with the row index attribute
                        writer.WriteStartElement(new Row() { RowIndex = (uint)rowNum });

                        for (int columnNum = 1; columnNum <= 50; columnNum++)
                        {
                            //write the cell start element with the type and reference attributes
                            //writer.WriteStartElement(new Cell() { CellReference = string.Format("{0}{1}", GetColumnName(columnNum), rowNum), DataType = CellValues.String });
                            cell.CellReference = string.Format("{0}{1}", GetColumnName(columnNum), rowNum);
                            cell.DataType = CellValues.SharedString;
                            writer.WriteStartElement(cell);
                            //write the cell value
                            //writer.WriteElement(new CellValue(string.Format("This is Row {0}, Cell {1}", rowNum, columnNum)));
                            value.Text = _sharedStringsToIndex[_allData[rowNum - 1, columnNum - 1]];
                            writer.WriteElement(value);
                            // write the end cell element
                            writer.WriteEndElement();
                        }

                        // write the end row element
                        writer.WriteEndElement();
                    }

                    // write the end SheetData element
                    writer.WriteEndElement();
                    // write the end Worksheet element
                    writer.WriteEndElement();
                    writer.Close();

                    writer = OpenXmlWriter.Create(document.WorkbookPart);
                    writer.WriteStartElement(new Workbook());
                    writer.WriteStartElement(new Sheets());

                    writer.WriteElement(new Sheet()
                    {
                        Name = "Large Sheet",
                        SheetId = 1,
                        Id = document.WorkbookPart.GetIdOfPart(workSheetPart)
                    });

                    // End Sheets
                    writer.WriteEndElement();
                    // End Workbook
                    writer.WriteEndElement();

                    writer.Close();

                    document.Save();

                    document.Close();
                }
            }
        }
    }

    private void GenerateSharedStringsTable(SharedStringTablePart sharedStringTablePart)
    {
        using (var writer = OpenXmlWriter.Create(sharedStringTablePart))
        {
            writer.WriteStartElement(new SharedStringTable() { Count = (UInt32)_sharedStringsCount, UniqueCount = (UInt32)_sharedStringsUniqueCount });

            foreach (var key in _sharedStringsToIndex.Keys)
            {
                writer.WriteStartElement(new SharedStringItem());
                writer.WriteElement(new Text(key));
                writer.WriteEndElement();
            }

            writer.WriteEndElement();

            writer.Close();
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
