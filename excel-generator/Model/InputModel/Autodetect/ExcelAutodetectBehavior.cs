using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelGenerator.Excel;

/// <summary>
/// Class describing the rules needed when auto detecting a data type on a column
/// </summary>
public class ExcelAutodetectBehavior
{
    public ExcelDateAutodetect Date { get; set; }
    public ExcelHyperlinkAutodetect Hyperlink { get; set; }
}

