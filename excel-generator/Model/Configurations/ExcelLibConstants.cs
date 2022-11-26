using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace rbkApiModules.Utilities.Excel;

public static class ExcelLibConstants
{
    public static class Configuration
    {
        public const int NumLengthSamples = 50;
        public const string ColorPattern = @"^([A-Fa-f0-9]{8})$";
    }

    public static class StyleContants
    {
        public static UInt32 StartIndex
        {
            get
            {
                return 164;
            }
        }
    }
}

