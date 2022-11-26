using BenchmarkDotNet.Attributes;
using OfficeOpenXml;

namespace ExcelGenerator.ForBenchmarking;

[MemoryDiagnoser]
public class EPPlusStyledBasedGenerator
{
    [Benchmark]
    public void EPPlusFree()
    {
        CreatePackage();
    }


    public void CreatePackage()
    {
        using (var stream = new MemoryStream())
        {
            ExcelPackage excel = new ExcelPackage();

            var ws = excel.Workbook.Worksheets.Add("Sheet1");

            for (int rowNum = 1; rowNum <= 100000; rowNum++)
            {
                for (int columnNum = 1; columnNum <= 50; columnNum++)
                {
                    ws.Cells[string.Format("{0}{1}", GetColumnName(columnNum), rowNum)].Value = string.Format("This is Row {0}, Column {1}", rowNum, columnNum);
                }
            }

            excel.SaveAs(stream);
        }    
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
            dividend = (dividend - modifier) / 26;
        }

        return columnName;
    }
}
