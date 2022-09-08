using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;


namespace ExcelGenerator.ForBenchmarking;

[MemoryDiagnoser]
public class ClosedXMLBasedGenerator
{
    [Benchmark]
    public void ClosedXML()
    {
        CreatePackage();
    }


    public void CreatePackage()
    {
        using (var stream = new MemoryStream())
        {
            using var wbook = new XLWorkbook();

            var ws = wbook.Worksheets.Add("Sheet1");
            
            for (int rowNum = 1; rowNum <= 100000; rowNum++)
            {
                for (int columnNum = 1; columnNum <= 50; columnNum++)
                {
                    ws.Cell(string.Format("{0}{1}", GetColumnName(columnNum), rowNum)).Value = string.Format("This is Row {0}, Column {1}", rowNum, columnNum);
                }
            }

            wbook.SaveAs(stream);
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
