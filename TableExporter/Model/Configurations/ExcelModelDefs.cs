﻿namespace TableExporter;

public static class ExcelModelDefs
{
    public enum ExcelSheetTypes
        {
            Table = 0,
            Chart = 1
        }

    public enum ExcelThemes
    {
        None = 0,
        TableStyleLight1 = 1,
        TableStyleLight2 = 2,
        TableStyleLight3 = 3,
        TableStyleLight4 = 4,
        TableStyleLight5 = 5,
        TableStyleLight6 = 6,
        TableStyleLight7 = 7,
        TableStyleLight8 = 8,
        TableStyleLight9 = 9,
        TableStyleLight10 = 10,
        TableStyleLight11 = 11,
        TableStyleLight12 = 12,
        TableStyleLight13 = 13,
        TableStyleLight14 = 14,
        TableStyleLight15 = 15,
        TableStyleLight16 = 16,
        TableStyleLight17 = 17,
        TableStyleLight18 = 18,
        TableStyleLight19 = 19,
        TableStyleLight20 = 20,
        TableStyleLight21 = 21,
        TableStyleMedium1 = 22,
        TableStyleMedium2 = 23,
        TableStyleMedium3 = 24,
        TableStyleMedium4 = 25,
        TableStyleMedium5 = 26,
        TableStyleMedium6 = 27,
        TableStyleMedium7 = 28,
        TableStyleMedium8 = 29,
        TableStyleMedium9 = 30,
        TableStyleMedium10 = 31,
        TableStyleMedium11 = 32,
        TableStyleMedium12 = 33,
        TableStyleMedium13 = 34,
        TableStyleMedium14 = 35,
        TableStyleMedium15 = 36,
        TableStyleMedium16 = 37,
        TableStyleMedium17 = 38,
        TableStyleMedium18 = 39,
        TableStyleMedium19 = 40,
        TableStyleMedium20 = 41,
        TableStyleMedium21 = 42,
        TableStyleDark1 = 43,
        TableStyleDark2 = 44,
        TableStyleDark3 = 45,
        TableStyleDark4 = 46,
        TableStyleDark5 = 47,
        TableStyleDark6 = 48,
        TableStyleDark7 = 49,
        TableStyleDark8 = 50,
        TableStyleDark9 = 51,
        TableStyleDark10 = 52,
        TableStyleDark11 = 53
    }    
    

    public static class ExcelFonts
    {
        public enum FontType
        {
            Arial = 0,
            Calibri = 1,
            CalibriLight = 2,
            CourierNew = 3,
            TimesNewRoman = 4
        }

        public static double GetFontSizeFactor(FontType fontType)
        {
            switch ((int)fontType)
            {
                case 0: return 6.5D;
                case 1: return 7D;
                case 2: return 7D;
                case 3: return 7D;
                case 4: return 7D;
                default: return 7D;
            }
        }
    }

    public enum ExcelDataTypes
    {
        Text = 0,
        Number = 1,
        DateTime = 2,
        Hyperlink = 3,
        AutoDetect = 4,
        Sheetlink = 5
    }
    
}
